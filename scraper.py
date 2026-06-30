"""Cell 期刊 issue 爬虫。

用法:
    1. 把要爬的 issue URL 放进同目录 urls.txt（每行一个，# 开头为注释）
    2. python scraper.py [--out output.xlsx]

流程:
    1. 用 Playwright 持久化浏览器逐个打开 urls.txt 里的 issue URL
    2. 检测 Cloudflare，若被挑战则提示用户手动过验证后回终端按 Enter
    3. 抽取每个 issue 中 Articles / Short Articles / Resources 三个分类下的文章链接
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

# 只爬这三个 section
WANTED_SECTIONS = {"Articles", "Short Articles", "Resources"}

PROFILE_DIR = Path(__file__).parent / "browser_profile"
URLS_FILE = Path(__file__).parent / "urls.txt"
CACHE_DIR = Path(__file__).parent / "cache"


# ---------- Callbacks ----------
# 把 print/input 抽象成 hook，CLI 用默认实现，Web 服务子类化覆盖。


class ScraperCallbacks:
    """默认实现：终端 print + input。Web 模式请子类化覆盖 log/cf_wait/is_cancelled/on_state。"""

    def log(self, msg: str) -> None:
        """普通日志输出。"""
        print(msg)

    def cf_wait(self, target_url: str, current_url: str,
                attempt: int, max_attempts: int) -> bool:
        """CF 触发时阻塞等待用户处理。返回 True 表示用户已处理；False 表示中止。"""
        input(f">>> 完成后回到此终端按 Enter（第 {attempt}/{max_attempts} 次）: ")
        return True

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


# 直接扫描 HTML 内容判断 CF，不依赖 title() / innerText（CF 挑战页常把它们清空）
CF_HTML_MARKERS = (
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
    """扫描 page.content() 判断 CF；网络异常时返回 True（保守处理）。"""
    try:
        html = page.content()
    except Exception:
        return True
    is_cf, _ = detect_cf_in_html(html)
    if is_cf:
        return True
    # URL 兜底
    return "cdn-cgi/challenge" in (page.url or "").lower()


def wait_until_cf_clear(page, target_url: str | None = None,
                        max_attempts: int = 3, cb: ScraperCallbacks | None = None) -> bool:
    """检测并处理 CF；最多提示用户 max_attempts 次。返回是否已通过。

    设计要点：
    - 不自动 reload/goto（用户反馈自动重导航会再次触发 CF 造成循环）
    - 只检测 + 提示 + 等用户手动操作 + 再检测
    - 用户应在浏览器中手动让页面停在目标 URL 上
    """
    cb = cb or ScraperCallbacks()
    if not is_cloudflare(page):
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
        cb.on_state({
            "phase": "cf_blocked",
            "target": target_url or "",
            "current": cur,
            "attempt": attempt,
            "max_attempts": max_attempts,
        })
        if not cb.cf_wait(target_url or "", cur, attempt, max_attempts):
            return False
        if cb.is_cancelled():
            return False

        # 给页面一点时间稳定
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
        cleaned.append((item.get("section", "").strip(), url, item.get("title", "").strip()))
    return cleaned


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
        "title": title, "doi": doi, "authors": authors, "first_aff": first_aff,
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
    return CACHE_DIR / f"{pii}.json"


def load_from_cache(pii: str) -> dict | None:
    p = cache_path(pii)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def save_to_cache(pii: str, fields: dict) -> None:
    CACHE_DIR.mkdir(exist_ok=True)
    cache_path(pii).write_text(
        json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def load_urls() -> list[str]:
    """从同目录 urls.txt 读取 issue URL 列表，# 开头为注释，空行跳过。"""
    if not URLS_FILE.exists():
        print(f"[error] 未找到 {URLS_FILE}。请在该文件中每行写一个 issue URL。")
        sys.exit(2)
    urls = []
    for ln in URLS_FILE.read_text(encoding="utf-8").splitlines():
        s = ln.strip()
        if not s or s.startswith("#"):
            continue
        urls.append(s)
    if not urls:
        print(f"[error] {URLS_FILE} 中没有有效 URL。")
        sys.exit(2)
    return urls


def process_issue(page, issue_url: str, use_cache: bool = True,
                  cb: ScraperCallbacks | None = None) -> list:
    """处理单个 issue：返回该 issue 的 results 列表 [(section, url, fields), ...]。"""
    cb = cb or ScraperCallbacks()
    cb.log(f"\n[issue] 打开: {issue_url}")
    cb.on_state({"phase": "issue_start", "issue_url": issue_url})
    try:
        page.goto(issue_url, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        cb.log(f"[warn] issue 页 goto 异常: {str(e)[:120]}")
    if not wait_until_cf_clear(page, target_url=issue_url, cb=cb):
        cb.log("[error] issue 页 CF 未通过，跳过此 issue")
        return []

    try:
        page.wait_for_selector("a.article-content-title", timeout=20000)
    except PWTimeout:
        cb.log("[warn] 未找到文章链接，可能 CF 未真正通过或页面结构变化。")
        return []

    all_articles = extract_article_list(page)
    cb.log(f"[*] 共发现 {len(all_articles)} 篇文章（全部 section）")
    targets = [t for t in all_articles if t[0] in WANTED_SECTIONS]
    cb.log(f"[*] 过滤到 Articles/Short Articles/Resources：{len(targets)} 篇")
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
                continue

        cb.log(f"\n[{i}/{total}] 打开: {url}")
        cb.log(f"        标题(list): {list_title[:80]}")
        # 进文章前随机停顿 + 鼠标抖动
        rwait(1.5, 4.0)
        random_mouse_jitter(page, moves=random.randint(1, 3))
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
        except PWTimeout:
            cb.log("        [warn] goto 超时，继续尝试解析")
        except Exception as e:
            msg = str(e)
            if "ERR_ABORTED" in msg or "net::ERR_" in msg:
                cb.log(f"        [warn] goto 被中断（{msg[:80]}），疑似 CF 拦截，进入手动处理")
            else:
                cb.log(f"        [warn] goto 异常: {msg[:120]}")

        # CF 检测 + 用户手动处理
        if not wait_until_cf_clear(page, target_url=url, cb=cb):
            cb.log(f"        [error] CF 未通过，本篇记为 [CF BLOCKED]（不写缓存，下次重试）")
            fields = {"title": "[CF BLOCKED]", "doi": "", "authors": [],
                      "first_aff": "", "url": url}
            results.append((section, url, fields))
            cb.on_state({"phase": "article_done", "url": url,
                         "fields": fields, "blocked": True})
            rwait(5.0, 10.0)
            continue

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

        fields = extract_fields(page, url)
        # 疑似 CF 残留页：再过一次并重新解析
        if not fields["title"] or "are you a robot" in fields["title"].lower() \
           or "just a moment" in fields["title"].lower():
            cb.log("        [warn] 解析失败（疑似 CF 页），再次进入手动处理")
            if wait_until_cf_clear(page, target_url=url, cb=cb):
                human_pause(page, 1.2, 2.8)
                fields = extract_fields(page, url)
        results.append((section, url, fields))

        cb.log(f"        标题(art): {fields['title'][:80]}")
        cb.log(f"        DOI:       {fields['doi']}")
        cb.log(f"        作者数:    {len(fields['authors'])}  "
               f"{'; '.join(fields['authors'][:3])}{' ...' if len(fields['authors']) > 3 else ''}")
        cb.log(f"        首条单位:  {fields['first_aff'][:100]}")

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
    return f"cell_{date.today().isoformat()}.xlsx"


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
            write_excel(out_path, all_issues, columns=NATURE_COLUMNS,
                        col_widths=[55, 55, 25, 14, 18, 60, 14, 10, 50])
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

    cb.log(f"\n[done] 全部完成，结果写入 {out_path}")
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
