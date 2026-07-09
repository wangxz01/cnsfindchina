"""Cell 期刊 issue 爬虫。

用法:
    1. 把要爬的 issue URL 放进同目录 urls.txt（每行一个，# 开头为注释）
    2. python scraper.py [--out output.xlsx]

流程:
    1. 用 Playwright 持久化浏览器逐个打开 urls.txt 里的 issue URL
    2. 检测 Cloudflare，若被挑战则提示用户手动过验证后回终端按 Enter
    3. 抽取每个 issue 中所有分类下的文章链接（不过滤 section）
    4. 逐篇打开文章页，从首次加载的 HTML 内嵌 JSON 提取
       标题/DOI/作者/首条 affiliation；若 affiliation 缺失，回退 JS 点击 #show-more-btn
    5. 每个 issue 一个 sheet 写入 Excel；多次运行从零开始，不做断点续爬
"""
from __future__ import annotations

import argparse
import json
import random
import re
import sys
import time
from datetime import date
from pathlib import Path
from urllib.parse import urljoin

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from excel_writer import write_excel, NATURE_COLUMNS
# 共用辅助（避免与 scraper_common 循环 import：scraper_common 用 TYPE_CHECKING 引用本模块）
from scraper_common import (
    cache_path as _common_cache_path,
    load_from_cache as _common_load_cache,
    save_to_cache as _common_save_cache,
    load_urls as _common_load_urls,
    default_out_path as _common_default_out,
    safe_goto as _safe_goto,
    count_real_articles,
    DATA_DIR,
)


# 国别判定 lazy import（避免与 nature_scraper 循环 import）
def _country_helpers():
    try:
        from nature_scraper import parse_country, is_china_country
        return parse_country, is_china_country
    except ImportError:
        return (lambda aff: ""), (lambda c: False)

# ---------- 常量 ----------

BASE = "https://www.sciencedirect.com"
ISSUE_URL_EXAMPLE = "https://www.sciencedirect.com/journal/cell/vol/189/issue/10"

PROFILE_DIR = DATA_DIR / "browser_profile"
URLS_FILE = DATA_DIR / "urls.txt"
CACHE_DIR = DATA_DIR / "cache"


# ---------- Callbacks ----------
# 把 print/input 抽象成 hook，CLI 用默认实现，Web 服务子类化覆盖。


class ScraperCallbacks:
    """默认实现：终端 print + input。Web 模式请子类化覆盖 log/cf_wait/is_cancelled/on_state。"""

    def log(self, msg: str) -> None:
        """普通日志输出。"""
        print(msg)

    def cf_wait(self, target_url: str, current_url: str,
                attempt: int, max_attempts: int,
                check_clear=None) -> str:
        """CF 触发时阻塞等待用户处理。

        返回：
          'user_resumed' —— 用户确认已处理
          'auto_cleared' —— 页面已自动消退（CF 自消，无需用户操作）
          'skip'         —— 用户跳过此文章（记为 BLOCKED 继续下一篇）
          'cancelled'    —— 用户中止整个任务

        check_clear：可选回调，调用方应在轮询中调用，返回 True 表示页面已恢复。
        CLI 实现可忽略（终端 input 等用户）；Web 实现用于自动消退检测。
        """
        input(f">>> 完成后回到此终端按 Enter（第 {attempt}/{max_attempts} 次）: ")
        return "user_resumed"

    def is_cancelled(self) -> bool:
        """是否被外部取消（如 Web 端按了停止）。"""
        return False

    def on_state(self, state: dict) -> None:
        """结构化状态变更（phase + 当前进度）。CLI 默认忽略；Web 端转发给前端。"""
        pass

# ---------- 随机化辅助（避免访问节律过于规律导致 IP 被封） ----------


def rwait(min_s: float, max_s: float) -> None:
    """在 [min_s, max_s] 秒间随机停顿。"""
    time.sleep(random.uniform(min_s, max_s))


def human_pause(page, min_s: float = 0.4, max_s: float = 1.6) -> None:
    """页面内的随机短停顿（用 wait_for_timeout，避免与 goto 冲突）。"""
    page.wait_for_timeout(int(random.uniform(min_s, max_s) * 1000))


def human_scroll(page, steps: int | None = None) -> None:
    """模拟阅读式慢速滚动：分若干段向下滚，每段之间随机停顿。"""
    try:
        if steps is None:
            steps = random.randint(4, 8)
        for i in range(steps):
            # 每次滚动一段随机距离
            delta = random.randint(250, 700)
            page.mouse.wheel(0, delta)
            time.sleep(random.uniform(0.35, 1.1))
        # 偶尔回滚一点（像在读细节）
        if random.random() < 0.4:
            page.mouse.wheel(0, -random.randint(100, 300))
            time.sleep(random.uniform(0.3, 0.8))
    except Exception:
        pass


def random_mouse_jitter(page, moves: int | None = None) -> None:
    """在视口内随机移动鼠标若干次，模拟人类光标行为。"""
    try:
        if moves is None:
            moves = random.randint(2, 5)
        w, h = 1366, 900
        for _ in range(moves):
            x = random.randint(100, w - 100)
            y = random.randint(100, h - 100)
            page.mouse.move(x, y, steps=random.randint(8, 20))
            time.sleep(random.uniform(0.08, 0.35))
    except Exception:
        pass


# ---------- Cloudflare ----------


# 直接扫描 HTML 内容判断 CF / 拦截页（不含 cookie 同意墙的通用文字——
# 那些文字在用户已同意后仍可能留在 DOM 仅 CSS 隐藏，raw scan 会误判。
# Cookie 墙的检测见 is_cloudflare() 中的可见性检查。）
CF_HTML_MARKERS = (
    # Cloudflare
    "just a moment",
    "are you a robot",
    "checking if the site connection is secure",
    "verifying you are human",
    "cf-chl-bypass",
    "cf-browser-verification",
    "/cdn-cgi/challenge-platform/",
    "cf-turnstile",
    "sorry, you have been blocked",
    "enable javascript and cookies to continue",
    "attention required! | cloudflare",
    "cf-error-code",
    "challenge-platform",
    "_cf_chl_opt",
    "cf-mitigated: challenge",
    # Nature interstitial（"Thank you for visiting nature.com ..." 整页拦截）
    "thank you for visiting nature.com",
    "you are using a browser version with limited support",
    # 通用反爬
    "access denied",
    "request blocked",
    "to continue, please complete the security check",
    "your activity is a little suspicious",
)


# Cookie 同意墙常见选择器（用 element visibility 判断，避免误判已隐藏的横幅）
COOKIE_BANNER_SELECTORS = (
    "#cookie-policy-banner",       # Springer/Nature
    "#cookie-banner",              # Science.org / 通用
    "#cookie-notification",
    ".cc-banner",
    ".cookie-banner",
    ".cookie-disclaimer",
    ".cookie-message",
    '[id*="cookie-consent"]',
    '[class*="cookie-consent"]',
    '[class*="CookieConsent"]',
)


def detect_cf_in_html(html: str) -> tuple[bool, str]:
    """扫描原始 HTML 判断是否 CF 拦截页；返回 (是否 CF, 命中标记)。"""
    if not html:
        return False, ""
    low = html[:30000].lower()  # 只看头部 30KB，CF 标记都在前面
    # 1) <title> 内容
    m = re.search(r"<title[^>]*>([^<]+)</title>", low)
    if m:
        t = m.group(1).strip()
        if any(x in t for x in ("just a moment", "are you a robot", "attention required")):
            return True, f"title={t!r}"
    # 2) 正文标记
    for marker in CF_HTML_MARKERS:
        if marker in low:
            return True, f"marker={marker!r}"
    return False, ""


def is_cloudflare(page) -> bool:
    """扫描 page.content() 判断 CF/拦截；另检查 cookie 同意墙元素是否可见。

    返回 True 表示"页面被某种墙挡住，需要用户介入"。
    网络异常时返回 True（保守处理）。
    """
    try:
        html = page.content()
    except Exception:
        return True
    is_cf, _ = detect_cf_in_html(html)
    if is_cf:
        return True
    # URL 兜底
    if "cdn-cgi/challenge" in (page.url or "").lower():
        return True
    # Cookie 同意墙：用元素可见性判断（raw HTML 里的文字即便同意后仍在）
    try:
        blocked = page.evaluate(
            """
            (selectors) => {
                for (const sel of selectors) {
                    const els = document.querySelectorAll(sel);
                    for (const el of els) {
                        if (!el.offsetParent && !el.getClientRects().length) continue;
                        // 排除明确 aria-hidden 或 display:none 的
                        const style = window.getComputedStyle(el);
                        if (style.display === 'none' || style.visibility === 'hidden') continue;
                        if (parseFloat(style.opacity) < 0.1) continue;
                        return true;
                    }
                }
                return false;
            }
            """,
            list(COOKIE_BANNER_SELECTORS),
        )
        if blocked:
            return True
    except Exception:
        pass
    return False


def wait_until_cf_clear(page, target_url: str | None = None,
                        max_attempts: int = 3, cb: ScraperCallbacks | None = None,
                        force: bool = False) -> bool:
    """检测并处理 CF；最多提示用户 max_attempts 次。返回是否已通过。

    设计要点：
    - 不自动 reload/goto（用户反馈自动重导航会再次触发 CF 造成循环）
    - 只检测 + 提示 + 等用户手动操作 + 再检测
    - 用户应在浏览器中手动让页面停在目标 URL 上
    - force=True：跳过 is_cloudflare 首检，强制进入"提示 + cf_wait"流程
      —— 用于 title 异常但 is_cloudflare 未识别的挑战页（如 Elsevier 的
      "Are you a robot?"：标记在 citation_title meta 里，<title> 不命中）。
    """
    cb = cb or ScraperCallbacks()
    if not force and not is_cloudflare(page):
        return True

    for attempt in range(1, max_attempts + 1):
        cb.log(f"\n[Cloudflare] 检测到人机验证挑战（第 {attempt}/{max_attempts} 次）。")
        if target_url:
            cb.log(f"[Cloudflare] 目标 URL: {target_url}")
        try:
            cur = page.url
        except Exception:
            cur = "(unknown)"
        cb.log(f"[Cloudflare] 当前 URL: {cur}")
        cb.log("[Cloudflare] 请在浏览器窗口中：")
        cb.log("  1) 完成人机验证（勾选复选框或等自动放行）")
        cb.log("  2) 必要时手动把地址栏改成目标 URL 并回车")
        cb.log("  3) 确认浏览器停在目标文章页（非 CF 挑战页）")
        cb.log("  （挑战页常会自动消退，无需操作——程序会每 0.3s 检测一次）")
        cb.on_state({
            "phase": "cf_blocked",
            "target": target_url or "",
            "current": cur,
            "attempt": attempt,
            "max_attempts": max_attempts,
        })

        def _check_clear():
            try:
                return not is_cloudflare(page)
            except Exception:
                return False

        result = cb.cf_wait(target_url or "", cur, attempt, max_attempts, _check_clear)
        if result == "cancelled":
            return False
        if result == "skip":
            return False  # 调用方据此记为 [CF BLOCKED]，继续下一篇
        if result == "auto_cleared":
            cb.log("[Cloudflare] 检测到挑战页已自动消退，继续。")
            return True  # check_clear 已确认页面正常，跳过重复检测

        # result == "user_resumed"：给页面稳定时间 + 再检测
        time.sleep(1.5)
        try:
            page.wait_for_load_state("domcontentloaded", timeout=15000)
        except Exception:
            pass

        if not is_cloudflare(page):
            cb.log("[Cloudflare] 已通过。")
            return True

    cb.log("[Cloudflare] 多次尝试后仍未通过。")
    return False


# ---------- 提取 + 挑战页兜底（Cell/Nature/Science 共用） ----------


def _is_challenge_title(title: str) -> bool:
    """标题是否疑似 CF / 机器人挑战页（强制重提取的依据）。

    空 title 视为异常（强制让用户介入），避免把空标题静默写入缓存。
    """
    if not title:
        return True
    low = title.lower()
    return any(m in low for m in ("are you a robot", "just a moment", "attention required"))


def extract_with_cf_retry(page, url: str, cb, section: str, extract_fn,
                          human_pause_fn, max_retries: int = 5) -> dict | None:
    """提取 → 若疑似挑战页 → 强制等用户介入 → 重提取，循环至正常或达上限。

    返回正常 fields；用户取消或达到 max_retries 仍异常时返回 None
    （调用方记为 [CF BLOCKED]，不写缓存）。

    用于 Cell/Nature/Science 三个 scraper 的 process_issue：goto 通过 CF 检查后，
    extract_fields 拿到的 title 仍可能是 "Are you a robot?" 等——is_cloudflare
    未识别这种 Elsevier 挑战页，需要靠 title 兜底。
    """
    fields = extract_fn(page, url)
    if section:
        fields["type"] = section
    for attempt in range(1, max_retries + 1):
        if not _is_challenge_title(fields.get("title", "")):
            return fields
        cb.log(f"        [warn] 第 {attempt}/{max_retries} 次提取疑似挑战页 "
               f"(title={fields.get('title')!r})，请手动处理后继续")
        # force=True 跳过 is_cloudflare 首检；max_attempts=1 让本函数外层循环控制总次数
        if not wait_until_cf_clear(page, target_url=url, cb=cb,
                                   max_attempts=1, force=True):
            return None
        human_pause_fn(1.2, 2.8)
        fields = extract_fn(page, url)
        if section:
            fields["type"] = section
    return None


# ---------- issue 列表抽取 ----------


def extract_article_list(page) -> list[tuple[str, str, str]]:
    """返回 [(section, article_url, list_title)]，保持文档顺序、去重。

    section 归属：每篇文章链接 → 文档顺序上最近的前置 section-title（兼容 h2/h3）。
    """
    data = page.evaluate(
        """
        () => {
            const titles = Array.from(
                document.querySelectorAll("h2.section-title, h3.section-title, .section-title")
            );
            const articles = Array.from(
                document.querySelectorAll("a.article-content-title[href*='/science/article/pii/']")
            );
            const out = [];
            for (const a of articles) {
                let section = "";
                for (const t of titles) {
                    // t 在 a 之前（a 相对 t 是 FOLLOWING）则更新；遇到 t 在 a 之后即停
                    if (t.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING) {
                        section = (t.textContent || "").trim();
                    } else {
                        break;
                    }
                }
                const titleEl = a.querySelector(".js-article-title");
                out.push({
                    section,
                    href: a.getAttribute("href"),
                    title: titleEl ? titleEl.textContent.trim() : a.textContent.trim()
                });
            }
            return out;
        }
        """
    )
    seen = set()
    cleaned = []
    for item in data:
        href = item.get("href") or ""
        if not href:
            continue
        url = urljoin(BASE, href)
        m = re.search(r"/pii/(S\d+)", url)
        if not m:
            continue
        pii = m.group(1)
        if pii in seen:
            continue
        seen.add(pii)
        section = item.get("section", "").strip()
        title = item.get("title", "").strip()
        # 过滤掉勘误（Author Correction / Publisher Correction），不算专业论文
        if re.match(r"^(Author|Publisher)\s+Correction\s*:", title, re.IGNORECASE):
            continue
        cleaned.append((section, url, title))
    return cleaned


def count_expected(page) -> int:
    """独立计数：regex 扫描 page.content() 原始 HTML 中的所有 PII，去重后返回。

    与 extract_article_list 的 DOM 选择器（querySelectorAll）路径独立——
    直接扫描 HTML 字符串，能捕获 DOM 选择器遗漏的文章。
    """
    html = page.content()
    piis = set(re.findall(r'/science/article/pii/(S\d{16})', html))
    return len(piis)


# ---------- 文章字段抽取 ----------


def _extract_from_html(html: str) -> dict:
    # 标题
    m = re.search(r'<meta name="citation_title" content="([^"]+)"', html)
    title = m.group(1) if m else ""

    # DOI：优先 citation_doi meta；否则第一个 cell DOI
    m = re.search(r'<meta name="citation_doi" content="([^"]+)"', html)
    if not m:
        m = re.search(r"(10\.1016/j\.cell\.[^\s\"<>]+)", html)
    doi = m.group(1).rstrip(".,;)") if m else ""

    # 作者：按 author-id 去重（同一文章 JSON 在页面里会被嵌入多次）
    authors = []
    seen_aid = set()
    for m in re.finditer(r'"#name":"author","\$":\{([^}]*)\}', html):
        header = m.group(1)
        aid_m = re.search(r'"author-id":"([^"]+)"', header)
        aid = aid_m.group(1) if aid_m else ""
        # 在该 author 块后续窗口内找 given-name/surname
        tail = html[m.end():m.end() + 800]
        nxt = tail.find('"#name":"author"')
        if nxt > 0:
            tail = tail[:nxt]
        g = re.search(r'"#name":"given-name","_":"([^"]+)"', tail)
        s = re.search(r'"#name":"surname","_":"([^"]+)"', tail)
        name = ((g.group(1) if g else "") + " " + (s.group(1) if s else "")).strip()
        if aid and aid in seen_aid:
            continue
        if aid:
            seen_aid.add(aid)
        if name:
            authors.append(name)

    # 首条 affiliation：定位第一个 "id":"aff1" 之后的 textfn
    # （页面前部可能有"Preview"等非单位 textfn 干扰）
    first_aff = ""
    pos = html.find('"id":"aff1"')
    if pos < 0:
        pos = html.find('"id":"aff')
    if pos >= 0:
        m = re.search(r'"#name":"textfn","_":"([^"]+)"', html[pos:])
        if m:
            first_aff = m.group(1)

    # 一作 + 国别（与 Nature/Science 同口径）
    first_author = authors[0] if authors else ""
    parse_country, is_china_country = _country_helpers()
    country = parse_country(first_aff)
    is_china = is_china_country(country)

    return {
        "title": title, "doi": doi, "type": "", "authors": authors, "first_aff": first_aff,
        "first_author": first_author,
        "first_author_country": country,
        "is_china": is_china,
    }


def _click_show_more_js(page) -> bool:
    """用 JS 触发 #show-more-btn 的 click 事件，避免 Playwright 鼠标点击落到错误位置。

    原生 page.click() 会做 actionability 检查 + 滚动 + 鼠标事件模拟，
    遇到 sticky header / overlay / 按钮部分被遮挡时容易点错位置。
    JS btn.click() 直接派发 click 事件，行为确定。
    """
    try:
        return bool(page.evaluate(
            """
            () => {
                const btn = document.querySelector("#show-more-btn");
                if (!btn) return false;
                const text = (btn.textContent || "").toLowerCase();
                if (!text.includes("show more") && !text.includes("show full")) return false;
                try { btn.scrollIntoView({block: "center", inline: "center"}); } catch (e) {}
                btn.click();
                return true;
            }
            """
        ))
    except Exception:
        return False


def extract_fields(page, article_url: str) -> dict:
    html = page.content()
    fields = _extract_from_html(html)
    fields["url"] = article_url

    # affiliation 缺失 -> 回退：JS 点击 #show-more-btn
    if not fields["first_aff"]:
        # 点击前先随机鼠标抖动，降低节律性
        random_mouse_jitter(page, moves=random.randint(2, 4))
        human_pause(page, 0.3, 1.0)
        if _click_show_more_js(page):
            human_pause(page, 2.0, 4.0)  # 原 wait_for_timeout(2500)
            html = page.content()
            # 与主路径一致：定位 aff1 之后的 textfn
            pos = html.find('"id":"aff1"')
            if pos < 0:
                pos = html.find('"id":"aff')
            if pos >= 0:
                m = re.search(r'"#name":"textfn","_":"([^"]+)"', html[pos:])
                if m:
                    fields["first_aff"] = m.group(1)

    # 再兜底：DOM 选择器
    if not fields["first_aff"]:
        try:
            el = page.query_selector(".author-affiliation dd, .affiliation dd, dl.affiliation dd")
            if el:
                fields["first_aff"] = el.inner_text().strip()
        except Exception:
            pass

    # 标题兜底：DOM h1
    if not fields["title"]:
        try:
            el = page.query_selector("h1.article-title, .CoreArticle-header h1, h1")
            if el:
                fields["title"] = el.inner_text().strip()
        except Exception:
            pass

    return fields


# ---------- 主流程 ----------


# ---- 按 PII 缓存（断点续爬） ----
# 每篇成功抓取的文章以 cache/<PII>.json 存盘。
# 重跑时命中缓存就不再访问网络，但仍会写入 Excel，所以输出永远完整。
# 删除 cache/<PII>.json 可强制重抓某篇；删除整个 cache 目录则全部重抓。


def cache_path(pii: str) -> Path:
    return _common_cache_path(CACHE_DIR, pii)


def load_from_cache(pii: str) -> dict | None:
    return _common_load_cache(CACHE_DIR, pii)


def save_to_cache(pii: str, fields: dict) -> None:
    _common_save_cache(CACHE_DIR, pii, fields)


def load_urls() -> list[str]:
    """从同目录 urls.txt 读取 issue URL 列表，# 开头为注释，空行跳过。"""
    return _common_load_urls(URLS_FILE)


def process_issue(page, issue_url: str, use_cache: bool = True,
                  cb: ScraperCallbacks | None = None) -> list:
    """处理单个 issue：返回该 issue 的 results 列表 [(section, url, fields), ...]。"""
    cb = cb or ScraperCallbacks()
    cb.log(f"\n[issue] 打开: {issue_url}")
    cb.on_state({"phase": "issue_start", "issue_url": issue_url})
    if not _safe_goto(page, issue_url, cb, max_retries=2):
        cb.log("[error] issue 页 goto 多次重试失败，跳过此 issue")
        return []
    if not wait_until_cf_clear(page, target_url=issue_url, cb=cb):
        cb.log("[error] issue 页 CF 未通过，跳过此 issue")
        return []

    try:
        page.wait_for_selector("a.article-content-title", timeout=20000)
    except PWTimeout:
        cb.log("[warn] 未找到文章链接，可能 CF 未真正通过或页面结构变化。")
        return []

    all_articles = extract_article_list(page)
    targets = all_articles
    # 数量检验：独立 regex 计数 vs DOM 提取计数（避免循环论证）
    expected = count_expected(page)
    cb.on_state({"phase": "count_check", "issue_url": issue_url,
                 "expected": expected, "actual": len(targets)})
    if expected != len(targets):
        cb.log(f"[warn] 数量检验不一致：DOM 提取 {len(targets)} 篇 vs 页面 regex {expected} 篇")
    else:
        cb.log(f"[check] 数量检验通过：{len(targets)} 篇 == 页面 {expected} 篇")
    cb.log(f"[*] 共 {len(targets)} 篇文章（全部 section，不过滤）")
    for sec, url, ttl in targets:
        cb.log(f"      - [{sec}] {ttl[:60]}")
    cb.on_state({"phase": "issue_plan", "issue_url": issue_url,
                 "targets": [(s, u, t) for s, u, t in targets]})

    results = []
    total = len(targets)
    cache_hits = 0
    for i, (section, url, list_title) in enumerate(targets, start=1):
        if cb.is_cancelled():
            cb.log("[*] 收到取消信号，停止当前 issue")
            break
        m = re.search(r"/pii/(S\d+)", url)
        pii = m.group(1) if m else ""

        cb.on_state({"phase": "article_start", "issue_url": issue_url,
                     "article_idx": i, "article_total": total,
                     "section": section, "url": url, "title": list_title})

        # 缓存命中：直接用，不访问网络
        if use_cache and pii:
            cached = load_from_cache(pii)
            if cached:
                cache_hits += 1
                cb.log(f"\n[{i}/{total}] [cache] {list_title[:70]}")
                results.append((section, url, cached))
                cb.on_state({"phase": "article_done", "url": url,
                             "fields": cached, "cached": True})
                # 小延迟让前端进度条来得及渲染
                time.sleep(0.05)
                continue

        cb.log(f"\n[{i}/{total}] 打开: {url}")
        cb.log(f"        标题(list): {list_title[:80]}")
        # 进文章前随机停顿 + 鼠标抖动
        rwait(1.5, 4.0)
        random_mouse_jitter(page, moves=random.randint(1, 3))
        if not _safe_goto(page, url, cb, max_retries=2):
            cb.log("        [error] goto 重试均失败，跳过此篇")
            fields = {
                "url": url, "title": "[GOTO FAILED]", "doi": pii, "type": section,
                "first_author": "", "first_aff": "", "first_author_country": "",
                "is_china": False, "authors": [],
            }
            results.append((section, url, fields))
            cb.on_state({"phase": "article_done", "url": url,
                         "fields": fields, "blocked": True})
            rwait(5.0, 10.0)
            continue

        # 文章页 CF 检测交给 extract_with_cf_retry 兜底（基于提取结果判定，更精准）；
        # 这里直接进 SPA 渲染等待 + 提取流程

        # 等待 SPA 渲染
        try:
            page.wait_for_selector("body", timeout=10000)
        except PWTimeout:
            pass
        human_pause(page, 1.2, 2.8)

        # 模拟阅读：缓慢滚动
        human_scroll(page)
        try:
            page.evaluate("() => window.scrollTo({top: 0, behavior: 'instant'})")
        except Exception:
            pass
        human_pause(page, 0.5, 1.5)

        fields = extract_with_cf_retry(page, url, cb, section,
                                       extract_fields, human_pause, max_retries=5)
        if fields is None:
            cb.log("        [error] 多次重试仍是挑战页，记为 [CF BLOCKED]")
            fields = {
                "url": url, "title": "[CF BLOCKED]", "doi": pii, "type": section,
                "first_author": "", "first_aff": "", "first_author_country": "",
                "is_china": False, "authors": [],
            }
        results.append((section, url, fields))

        cb.log(f"        标题(art): {fields['title'][:80]}")
        cb.log(f"        DOI:       {fields['doi']}")
        cb.log(f"        作者数:    {len(fields['authors'])}  "
               f"{'; '.join(fields['authors'][:3])}{' ...' if len(fields['authors']) > 3 else ''}")
        cb.log(f"        首条单位:  {fields['first_aff'][:100]}")

        # 让 cache 携带 section + issue_url，供"立即导出"按 issue 分组、按 section 分类
        fields["section"] = section
        fields["issue_url"] = issue_url
        # 只在拿到真实数据时写缓存；CF BLOCKED 不写
        if pii and fields.get("title") and fields["title"] != "[CF BLOCKED]":
            save_to_cache(pii, fields)

        cb.on_state({"phase": "article_done", "url": url,
                     "fields": fields, "cached": False})

        # 文章间随机停顿；每 5-8 篇插入一次长歇
        if i % random.randint(5, 8) == 0:
            pause = random.uniform(45.0, 120.0)
            cb.log(f"        [*] 第 {i} 篇完成，长歇 {pause:.0f}s 模拟阅读间歇...")
            for _ in range(int(pause)):
                if cb.is_cancelled():
                    break
                time.sleep(1)
        else:
            rwait(6.0, 15.0)

    if cache_hits:
        cb.log(f"\n[issue] 缓存命中 {cache_hits}/{total} 篇（未访问网络）")
    cb.on_state({"phase": "issue_done", "issue_url": issue_url, "count": len(results)})
    return results


def default_out_path() -> str:
    """默认输出文件名按日期：cell_YYYY-MM-DD.xlsx"""
    return _common_default_out("cell")


def run_scraper(urls: list[str], out_path: str | Path,
                cb: ScraperCallbacks | None = None,
                use_cache: bool = True, headless: bool = False) -> list:
    """主抓取流程（CLI 与 Web 共用）。

    返回 all_issues: [(issue_url, [(section, url, fields), ...]), ...]
    """
    cb = cb or ScraperCallbacks()
    out_path = str(out_path)
    all_issues: list[tuple[str, list]] = []

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=headless,
            viewport={"width": 1366, "height": 900},
            args=["--disable-blink-features=AutomationControlled"],
        )
        ctx.add_init_script(
            "() => { Object.defineProperty(navigator, 'webdriver', {get: () => undefined}); }"
        )
        page = ctx.new_page()

        for idx, issue_url in enumerate(urls, start=1):
            if cb.is_cancelled():
                cb.log("[*] 收到取消信号，停止整个抓取流程")
                break
            cb.log(f"\n========== issue {idx}/{len(urls)} ==========")
            cb.on_state({"phase": "issue_progress", "issue_idx": idx,
                         "issue_total": len(urls), "issue_url": issue_url})
            results = process_issue(page, issue_url, use_cache=use_cache, cb=cb)
            all_issues.append((issue_url, results))
            actual_out = write_excel(out_path, all_issues, columns=NATURE_COLUMNS,
                        col_widths=[55, 55, 25, 14, 18, 60, 14, 10, 50])
            if actual_out != out_path:
                cb.log(f"[warn] 原文件被占用，实际写入: {actual_out}")
                out_path = actual_out
            cb.log(f"\n[issue {idx}] 完成，共 {len(results)} 篇；已写入 {out_path}")
            cb.on_state({"phase": "excel_written", "out_path": out_path,
                         "issue_url": issue_url})
            if idx < len(urls):
                pause = random.uniform(20.0, 60.0)
                cb.log(f"[*] issue 间长歇 {pause:.0f}s ...")
                for _ in range(int(pause)):
                    if cb.is_cancelled():
                        break
                    time.sleep(1)

        ctx.close()

    # 统计实际拿到的（非 BLOCKED/FAILED）文章数；为 0 视为异常
    real_count = count_real_articles(all_issues)
    if real_count == 0:
        cb.log(f"\n[warn] 全部完成但 0 篇成功（可能 CF/cookie 墙未过或结构变化）")
        cb.on_state({"phase": "all_skipped", "out_path": out_path,
                     "reason": "0 篇文章抓取成功"})
    else:
        cb.log(f"\n[done] 全部完成（{real_count} 篇），结果写入 {out_path}")
        cb.on_state({"phase": "all_done", "out_path": out_path})
    return all_issues


def main() -> int:
    ap = argparse.ArgumentParser(description="Cell journal issue scraper")
    ap.add_argument("--out", default=None,
                    help=f"output xlsx path (默认: {default_out_path()})")
    ap.add_argument("--headless", action="store_true", help="headless mode (NOT recommended; CF needs manual)")
    ap.add_argument("--fresh", action="store_true",
                    help="忽略缓存，全部重抓（等价于 rm -rf cache）")
    args = ap.parse_args()

    out_path = Path(args.out if args.out else default_out_path()).resolve()
    urls = load_urls()
    print(f"[*] 从 urls.txt 读到 {len(urls)} 个 issue URL；use_cache={not args.fresh}")
    run_scraper(urls, out_path, use_cache=not args.fresh, headless=args.headless)
    return 0


if __name__ == "__main__":
    sys.exit(main())
