"""Science 期刊 issue 爬虫（判断一作是否在中国）。

用法:
    1. 把 issue URL 放进 urls_science.txt（每行一个，# 注释）
    2. python science_scraper.py [--out science_YYYY-MM-DD.xlsx]

流程:
    1. Playwright 持久化浏览器逐个打开 issue URL
    2. Cookie 同意墙 / 反爬触发时暂停等用户手动处理（复用 Cell 的 CF 检测）
    3. 抽取所有分类的文章（按 h5.to-section 分组，不过滤 section）
    4. 逐篇打开文章页，提取：
       标题 / DOI / 类型 / 一作 / 一作单位 (#con1_content 内首个 affiliation) /
       一作国家 / 是否中国 / 作者列表
    5. 每个 issue 一个 sheet 写入 Excel（同 Cell 的"链接-》文章"布局）
    6. 按 DOI 缓存到 cache_science/ 支持断点续爬

Affiliation 通常已在 DOM 中（display:none），直取即可；个别文章需点击
"Authors Info & Affiliations" + "Expand All" 才注入 DOM，作为兜底。
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

# 复用基础设施
from scraper import (
    ScraperCallbacks,
    rwait, human_pause, human_scroll, random_mouse_jitter,
    is_cloudflare, wait_until_cf_clear,
    extract_with_cf_retry,
)
from nature_scraper import parse_country, is_china_country  # 国别判定逻辑完全一致
from excel_writer import write_excel, NATURE_COLUMNS  # 列定义相同
# 共用辅助（cache / urls / 默认输出 / 计数）；safe_goto 通过 _safe_goto 包装器转发
from scraper_common import (
    cache_path as _common_cache_path,
    load_from_cache as _common_load_cache,
    save_to_cache as _common_save_cache,
    load_urls as _common_load_urls,
    default_out_path as _common_default_out,
    count_real_articles,
    DATA_DIR,
)

BASE = "https://www.science.org"

PROFILE_DIR = DATA_DIR / "browser_profile_science"
URLS_FILE = DATA_DIR / "urls_science.txt"
CACHE_DIR = DATA_DIR / "cache_science"


# ---------- issue 列表抽取 ----------


def extract_article_list(page) -> list[tuple[str, str, str]]:
    """返回 [(section, article_url, list_title)]，文档顺序、去重。

    section 由最近前驱 h5.to-section 决定（compareDocumentPosition）。
    """
    data = page.evaluate(
        """
        () => {
            const headings = Array.from(document.querySelectorAll('h5.to-section'));
            const links = Array.from(document.querySelectorAll('a[href*="/doi/"]'));
            const filtered = links.filter(a => {
                const href = a.getAttribute('href') || '';
                return /\\/doi\\/(?:abs\\/|full\\/|pdf\\/|epdf\\/)?10\\.1126\\/science\\./.test(href);
            });

            const out = [];
            for (const a of filtered) {
                // 跳过 RELATED 跨链框
                if (a.closest('.card-related')) continue;
                const rawHref = a.getAttribute('href');
                const href = rawHref.replace(/^\\/doi\\/(abs|full|pdf|epdf)\\//, '/doi/');
                let section = '';
                for (const h of headings) {
                    if (h.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING) {
                        section = (h.textContent || '').trim();
                    } else {
                        break;
                    }
                }
                const text = (a.textContent || '').trim();
                // 跳过明显非标题锚文本（Abstract / Download PDF / +N authors）
                if (/^(Abstract|Download PDF|PDF|\\+\\d+ authors|fewer|Expand All)$/i.test(text)) continue;
                out.push({href, section, title: text.slice(0, 200)});
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
        m = re.search(r"/doi/(10\.1126/science\.[a-z0-9]+)", url)
        if not m:
            continue
        doi = m.group(1)
        if doi in seen:
            continue
        seen.add(doi)
        cleaned.append((item.get("section", "").strip(), url, item.get("title", "").strip()))
    return cleaned


def count_expected(page) -> int:
    """独立计数：regex 扫描 page.content() 原始 HTML 中的所有 science DOI，去重后返回。

    与 extract_article_list 的 DOM 选择器路径独立。
    只匹配 /doi/ URL 路径里的 DOI，避免 issue DOI (10.1126/science.2026.392.issue-XXXX) 干扰。
    """
    html = page.content()
    ids = set(re.findall(r'/doi/(?:abs/|full/|pdf/|epdf/)?(10\.1126/science\.[a-z0-9]+)', html))
    return len(ids)


# ---------- 文章字段抽取 ----------


def _strip_html(s: str) -> str:
    s = re.sub(r"<[^>]+>", "", s)
    s = re.sub(r"&nbsp;", " ", s)
    s = re.sub(r"&amp;", "&", s)
    return re.sub(r"\s+", " ", s).strip()


def _extract_fields_from_html(html: str, article_url: str) -> dict:
    # 标题
    m = re.search(r'<meta name="dc\.Title"\s+content="([^"]+)"', html)
    title = m.group(1) if m else ""
    if not title:
        m = re.search(r'<meta name="citation_title"\s+content="([^"]+)"', html)
        title = m.group(1) if m else ""

    # DOI：优先 publication_doi meta，否则 URL 里取
    m = re.search(r'<meta name="publication_doi"\s+content="([^"]+)"', html)
    if not m:
        m = re.search(r'<meta name="DOI"\s+content="([^"]+)"', html)
    if not m:
        m = re.search(r'(10\.1126/science\.[a-z0-9]+)', article_url)
    doi = (m.group(1) if m else "").replace("doi:", "").strip()

    # 类型
    m = re.search(r'<meta name="dc\.Type"\s+content="([^"]+)"', html)
    article_type = (m.group(1) if m else "").strip()

    # 作者：dc.Creator，顺序敏感
    authors = re.findall(r'<meta name="dc\.Creator"\s+content="([^"]+)"', html)
    norm_authors = [a.strip() for a in authors]
    first_author = norm_authors[0] if norm_authors else ""

    # 一作 Aff1：定位 #con1_content 内首个 affiliation 的 <span property="name">
    first_aff = ""
    pos = html.find('id="con1_content"')
    if pos < 0:
        # 兜底：找第一个 contributor 的 content
        pos = html.find('id="con1"')
    if pos >= 0:
        chunk = html[pos:pos + 4000]
        # 第一个 <span property="name">...</span>
        m = re.search(
            r'<div property="affiliation"[^>]*>\s*<span property="name">([^<]+)</span>',
            chunk,
        )
        if m:
            first_aff = _strip_html(m.group(1)).rstrip(".")
        else:
            # 再兜底：任何 affiliation span
            m = re.search(r'<span property="name">([^<]+)</span>', chunk)
            if m:
                first_aff = _strip_html(m.group(1)).rstrip(".")

    # 一作兜底：若无 dc.Creator，从 #con1 的 heading 取
    if not first_author:
        m = re.search(
            r'id="con1"[^>]*>.*?<span property="givenName">([^<]+)</span>\s*'
            r'<span property="familyName">([^<]+)</span>',
            html, re.DOTALL,
        )
        if m:
            first_author = f"{m.group(1).strip()} {m.group(2).strip()}"

    country = parse_country(first_aff)
    is_china = is_china_country(country)

    return {
        "url": article_url,
        "title": title,
        "doi": doi,
        "type": article_type,
        "first_author": first_author,
        "first_aff": first_aff,
        "first_author_country": country,
        "is_china": is_china,
        "authors": norm_authors,
    }


def extract_fields(page, article_url: str) -> dict:
    html = page.content()
    fields = _extract_fields_from_html(html, article_url)
    if not fields["title"]:
        try:
            el = page.query_selector("h1.core-self-citation__title, h1.highwire-cite-title, h1")
            if el:
                fields["title"] = el.inner_text().strip()
        except Exception:
            pass

    # Affiliation 兜底：直取失败时点击 "Authors Info & Affiliations" + "Expand All"
    if not fields["first_aff"]:
        try:
            # 滚到 contributors 区，触发可能的 lazy-load
            page.evaluate(
                "() => { const el = document.getElementById('tab-contributors'); "
                "if (el) el.scrollIntoView({block:'center'}); }"
            )
            page.wait_for_timeout(500)
            # 尝试点 Expand All（如果存在）
            clicked = page.evaluate(
                """
                () => {
                    const btn = document.querySelector('button[data-expandable="all"]');
                    if (!btn) return false;
                    btn.click();
                    return true;
                }
                """
            )
            if clicked:
                page.wait_for_timeout(2000)
                html = page.content()
                fields = _extract_fields_from_html(html, article_url)
                if not fields["title"]:
                    # 重新查 h1
                    try:
                        el = page.query_selector("h1")
                        if el:
                            fields["title"] = el.inner_text().strip()
                    except Exception:
                        pass
        except Exception:
            pass

    return fields


def _safe_goto(page, url: str, cb: ScraperCallbacks, max_retries: int = 2) -> bool:
    """goto 重试：转发到 scraper_common（避免在文件内重复实现）。"""
    from scraper_common import safe_goto
    return safe_goto(page, url, cb, max_retries)


# ---------- 缓存（薄包装到 scraper_common） ----------


def cache_path(doi: str) -> Path:
    return _common_cache_path(CACHE_DIR, doi)


def load_from_cache(doi: str) -> dict | None:
    return _common_load_cache(CACHE_DIR, doi)


def save_to_cache(doi: str, fields: dict) -> None:
    _common_save_cache(CACHE_DIR, doi, fields)


# ---------- issue 处理 ----------


def process_issue(page, issue_url: str, use_cache: bool = True,
                  cb: ScraperCallbacks | None = None) -> list:
    cb = cb or ScraperCallbacks()
    cb.log(f"\n[issue] 打开: {issue_url}")
    cb.on_state({"phase": "issue_start", "issue_url": issue_url})
    if not _safe_goto(page, issue_url, cb, max_retries=2):
        cb.log("[error] issue 页 goto 多次重试失败，跳过此 issue")
        return []
    if not wait_until_cf_clear(page, target_url=issue_url, cb=cb):
        cb.log("[error] issue 页 CF/cookie 墙未通过，跳过此 issue")
        return []

    # 强制等 JS 完全渲染（issue 页是 lazy-load，wait_for_selector 早返会漏文章）
    page.wait_for_timeout(5000)
    try:
        page.wait_for_selector('h5.to-section, a[href*="/doi/10.1126/science."]',
                               timeout=20000)
    except PWTimeout:
        cb.log("[warn] 未找到 section/文章选择器，可能页面结构变化或被拦截。")

    all_articles = extract_article_list(page)
    cb.log(f"[*] 共发现 {len(all_articles)} 个 article DOI（全 section）")
    if not all_articles:
        # 诊断：dump 实际 DOM 状态帮排查
        diag = page.evaluate("""
            () => ({
                h5_count: document.querySelectorAll('h5.to-section').length,
                h5_texts: Array.from(document.querySelectorAll('h5.to-section')).map(h => h.textContent.trim()).slice(0, 15),
                doi_link_count: document.querySelectorAll('a[href*="/doi/10.1126/science."]').length,
                url: location.href,
                title: document.title,
            })
        """)
        cb.log(f"[diag] DOM 状态: h5={diag['h5_count']}, doi链接={diag['doi_link_count']}")
        cb.log(f"[diag] h5.to-section 文本: {diag['h5_texts']}")
        cb.log(f"[diag] 当前 URL: {diag['url']}; title: {diag['title']}")
        cb.log("[error] 0 篇文章，请检查上方 DOM 状态（可能是 cookie 墙未过 / 反爬 / 结构变化）")
        return []

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
        m = re.search(r"/doi/(10\.1126/science\.[a-z0-9]+)", url)
        doi = m.group(1) if m else ""

        cb.on_state({"phase": "article_start", "issue_url": issue_url,
                     "article_idx": i, "article_total": total,
                     "section": section, "url": url, "title": list_title})

        if use_cache and doi:
            cached = load_from_cache(doi)
            if cached:
                cache_hits += 1
                cb.log(f"\n[{i}/{total}] [cache] {list_title[:70]}")
                results.append((section, url, cached))
                cb.on_state({"phase": "article_done", "url": url,
                             "fields": cached, "cached": True})
                # 小延迟让前端进度条/列表来得及渲染（全命中时 otherwise 瞬时跳到 100%）
                time.sleep(0.05)
                continue

        cb.log(f"\n[{i}/{total}] 打开: {url}")
        cb.log(f"        标题(list): {list_title[:80]}")
        # 进文章前随机停顿（拉长，避免 Science 限速）
        rwait(3.0, 6.0)
        random_mouse_jitter(page, moves=2)
        if not _safe_goto(page, url, cb, max_retries=2):
            cb.log("        [error] goto 重试均失败，跳过此篇")
            fields = {
                "url": url, "title": "[GOTO FAILED]", "doi": doi, "type": section,
                "first_author": "", "first_aff": "", "first_author_country": "",
                "is_china": False, "authors": [],
            }
            results.append((section, url, fields))
            rwait(15.0, 25.0)
            continue

        # 文章页 CF 检测交给 extract_with_cf_retry 兜底（基于提取结果判定，更精准）

        # 等渲染
        try:
            page.wait_for_selector("#con1_content, .core-authors, h1", timeout=15000)
        except PWTimeout:
            page.wait_for_timeout(2000)
        human_pause(page, 1.5, 3.0)
        human_scroll(page)

        fields = extract_with_cf_retry(page, url, cb, section,
                                       extract_fields, human_pause, max_retries=5)
        if fields is None:
            cb.log("        [error] 多次重试仍是挑战页，记为 [CF BLOCKED]")
            fields = {
                "url": url, "title": "[CF BLOCKED]", "doi": doi, "type": section,
                "first_author": "", "first_aff": "", "first_author_country": "",
                "is_china": False, "authors": [],
            }
        results.append((section, url, fields))

        cb.log(f"        标题: {fields['title'][:80]}")
        cb.log(f"        DOI:  {fields['doi']}")
        cb.log(f"        类型: {fields['type']}  一作: {fields['first_author']}")
        cb.log(f"        单位: {fields['first_aff'][:120]}")
        cb.log(f"        国家: {fields['first_author_country']}  "
               f"是否中国: {'是' if fields['is_china'] else '否'}")

        # 让 cache 携带 section + issue_url，供"立即导出"按 issue 分组、按 section 分类
        fields["section"] = section
        fields["issue_url"] = issue_url
        if doi and fields.get("title") and fields["title"] not in ("[CF BLOCKED]", "[GOTO FAILED]"):
            save_to_cache(doi, fields)

        cb.on_state({"phase": "article_done", "url": url,
                     "fields": fields, "cached": False})

        # 长间隔（Science 比 Nature 更敏感，拉长到 12-25s，每 4 篇长歇）
        if i % 4 == 0:
            sleep_dur = random.randint(60, 120)
            cb.log(f"        [*] 第 {i} 篇完成，长歇 {sleep_dur}s...")
            for _ in range(sleep_dur):
                if cb.is_cancelled():
                    break
                time.sleep(1)
        else:
            rwait(12.0, 25.0)

    if cache_hits:
        cb.log(f"\n[issue] 缓存命中 {cache_hits}/{total} 篇（未访问网络）")
    cb.on_state({"phase": "issue_done", "issue_url": issue_url, "count": len(results)})
    return results


# ---------- 主流程 ----------


def default_out_path() -> str:
    return _common_default_out("science")


def load_urls() -> list[str]:
    return _common_load_urls(URLS_FILE)


def run_scraper(urls: list[str], out_path: str | Path,
                cb: ScraperCallbacks | None = None,
                use_cache: bool = True, headless: bool = False) -> list:
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
                out_path = actual_out  # 后续 issue / all_done 用新路径
            cb.log(f"\n[issue {idx}] 完成，共 {len(results)} 篇；已写入 {out_path}")
            cb.on_state({"phase": "excel_written", "out_path": out_path,
                         "issue_url": issue_url})
            if idx < len(urls):
                pause = random.randint(30, 60)
                cb.log(f"[*] issue 间长歇 {pause}s ...")
                for _ in range(pause):
                    if cb.is_cancelled():
                        break
                    time.sleep(1)

        ctx.close()

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
    ap = argparse.ArgumentParser(description="Science journal issue scraper")
    ap.add_argument("--out", default=None,
                    help=f"output xlsx path (默认: {default_out_path()})")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--fresh", action="store_true", help="忽略缓存，全部重抓")
    args = ap.parse_args()

    out_path = Path(args.out if args.out else default_out_path()).resolve()
    urls = load_urls()
    print(f"[*] 从 urls_science.txt 读到 {len(urls)} 个 issue URL；use_cache={not args.fresh}")
    run_scraper(urls, out_path, use_cache=not args.fresh, headless=args.headless)
    return 0


if __name__ == "__main__":
    sys.exit(main())
