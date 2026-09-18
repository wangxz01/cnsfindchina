"""Nature 期刊 issue 爬虫（判断一作是否在中国）。

用法:
    1. 把要爬的 issue URL 放进 urls_nature.txt（每行一个，# 注释）
    2. python nature_scraper.py [--out nature_YYYY-MM-DD.xlsx]

流程:
    1. Playwright 持久化浏览器逐个打开 issue URL
    2. Cookie 同意墙 / 反爬触发时暂停等用户手动处理（复用 Cell 的 CF 检测）
    3. 抽取所有分类的 s41586- 研究论文（不过滤 section）
    4. 逐篇打开文章页，提取：
       标题 / DOI / 类型 / 一作 / 一作单位 (Aff1) / 一作国家 / 是否中国 / 作者列表
    5. 每个 issue 一个 sheet 写入 Excel（列定义见 excel_writer.NATURE_COLUMNS）
    6. 按 PII 缓存到 cache_nature/ 支持断点续爬

判定一作国别：解析 Aff1 文本末段，匹配国家名清单；是否中国 = 国家含 "China"。
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

# 复用 Cell scraper 的通用基础设施
from scraper import (
    ScraperCallbacks,
    rwait, human_pause, human_scroll, random_mouse_jitter,
    detect_cf_in_html, is_cloudflare, wait_until_cf_clear,
    extract_with_cf_retry,
)
from excel_writer import write_excel, NATURE_COLUMNS, column_widths
from article_metadata import enrich_fields, finalize_fields, publication_dates, expected_article_ids, count_check, china_label
from scraper_common import run_source, cancellable_sleep, merge_cached_fields, ScrapeCancelled

# 共用辅助
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


# ---------- 常量 ----------

BASE = "https://www.nature.com"

# Cell 与 Nature 用不同 profile / cache，避免 cookie 互染
PROFILE_DIR = DATA_DIR / "browser_profile_nature"
URLS_FILE = DATA_DIR / "urls_nature.txt"
CACHE_DIR = DATA_DIR / "cache_nature"


# ---------- 国别识别 ----------

# 常见国家别名 → 规范名（用于 Nature Aff1 末段匹配）
from countries import parse_country, is_china_country


# ---------- issue 列表抽取 ----------


def extract_article_list(page) -> list[tuple[str, str, str]]:
    """返回 [(section, article_url, list_title)]，文档顺序、去重、只含 s41586-。

    section 由最近前驱 h3.c-section-heading 决定（compareDocumentPosition）。
    """
    data = page.evaluate(
        """
        () => {
            const headings = Array.from(document.querySelectorAll('h3.c-section-heading'));
            const cards = Array.from(document.querySelectorAll(
                'article.c-card, article[itemtype*="ScholarlyArticle"]'
            ));
            const out = [];
            for (const card of cards) {
                if (card.closest('header, footer, nav, aside, .card-related')) continue;
                const a = card.querySelector('a[href*="/articles/"]');
                if (!a) continue;
                const href = a.getAttribute('href') || '';
                if (!/\\/articles\\/s41586-/.test(href)) continue;
                let section = '';
                for (const h of headings) {
                    if (h.compareDocumentPosition(card) & Node.DOCUMENT_POSITION_FOLLOWING) {
                        section = (h.textContent || '').trim();
                    } else {
                        break;
                    }
                }
                const titleEl = card.querySelector('.c-card__title');
                const title = titleEl ? titleEl.textContent.trim() : a.textContent.trim();
                out.push({section, href, title: title.slice(0, 200)});
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
        m = re.search(r"/articles/(s41586-\d{3}-\d{4,7}-[a-z0-9]+)", url)
        if not m:
            continue
        article_id = m.group(1)
        if article_id in seen:
            continue
        seen.add(article_id)
        section = (item.get("section") or "").strip()
        title = item.get("title", "").strip()
        # 过滤掉勘误（Author Correction / Publisher Correction），不算专业论文
        if re.match(r"^(Author|Publisher)\s+Correction\s*:", title, re.IGNORECASE):
            continue
        cleaned.append((section, url, title))
    return cleaned


def count_expected(page) -> int:
    return len(expected_article_ids(page.content(), 'nature'))


# ---------- 文章字段抽取 ----------


def _strip_html(s: str) -> str:
    s = re.sub(r"<[^>]+>", "", s)
    s = re.sub(r"&nbsp;", " ", s)
    s = re.sub(r"&amp;", "&", s)
    return re.sub(r"\s+", " ", s).strip()


def _extract_fields_from_html(html: str, article_url: str) -> dict:
    # 标题
    m = re.search(r'<meta name="dc.title"\s+content="([^"]+)"', html)
    title = m.group(1) if m else ""
    if not title:
        m = re.search(r'<meta name="citation_title"\s+content="([^"]+)"', html)
        title = m.group(1) if m else ""

    # DOI
    m = re.search(r'<meta name="prism.doi"\s+content="([^"]+)"', html)
    if not m:
        m = re.search(r'<meta name="doi"\s+content="([^"]+)"', html)
    if not m:
        m = re.search(r'(10\.1038/s41586[^\s"<>]+)', html)
    doi = (m.group(1) if m else "").replace("doi:", "").strip()

    # 类型（Article / Perspective）
    m = re.search(r'<meta name="dc.type"\s+content="([^"]+)"', html)
    if not m:
        m = re.search(r'<meta name="citation_article_type"\s+content="([^"]+)"', html)
    article_type = (m.group(1) if m else "").strip() or "Article"

    # 作者：优先 dc.creator（顺序敏感）；大型合作组无 dc.creator 时回退到 article 头部 DOM
    authors = re.findall(r'<meta name="dc.creator"\s+content="([^"]+)"', html)
    if not authors:
        # 回退：抓 article 作者列表
        m = re.search(
            r'<ul[^>]*class="[^"]*c-article-author-list[^"]*"[^>]*>(.*?)</ul>',
            html, re.DOTALL,
        )
        if m:
            authors = re.findall(
                r'<meta[^>]*content="([^"]+)"[^>]*itemprop="name"',
                m.group(1),
            )
        if not authors:
            # 再兜底：JSON-LD 或 data-test
            m = re.search(r'data-test="author-list"[^>]*>(.*?)</p>', html, re.DOTALL)
            if m:
                authors = [a.strip() for a in _strip_html(m.group(1)).split(",") if a.strip()]
    # 规整 dc.creator 的 "Lastname, Firstname" 格式 → "Firstname Lastname"
    norm_authors = []
    for a in authors:
        if "," in a and a.count(",") == 1:
            last, first = [p.strip() for p in a.split(",", 1)]
            if first and last:
                norm_authors.append(f"{first} {last}")
                continue
        norm_authors.append(a.strip())

    # 一作名字：第一个作者；如果没有作者列表，下面从 Aff1 兜底
    first_author = norm_authors[0] if norm_authors else ""

    # Aff1（第一单位）的地址与作者列表
    first_aff = ""
    aff1_authors = ""
    pos = html.find('id="Aff1"')
    if pos < 0:
        pos = html.find('id="aff1"')
    if pos >= 0:
        # 在 Aff1 <li> 范围内找 address 和 authors-list（截到下一个 Aff 或 </li> 或 </ol>）
        end = html.find("</li>", pos)
        if end < 0:
            end = pos + 3000
        chunk = html[pos:end]
        m_addr = re.search(
            r'class="[^"]*c-article-author-affiliation__address[^"]*"[^>]*>(.*?)</p>',
            chunk, re.DOTALL,
        )
        if m_addr:
            first_aff = _strip_html(m_addr.group(1))
        m_auths = re.search(
            r'class="[^"]*c-article-author-affiliation__authors-list[^"]*"[^>]*>(.*?)</p>',
            chunk, re.DOTALL,
        )
        if m_auths:
            aff1_authors = _strip_html(m_auths.group(1))

    # 一作兜底：从 Aff1 authors-list 的第一个名字取
    if not first_author and aff1_authors:
        # 用 & 或 , 分隔，取第一个
        first = re.split(r"[,&]", aff1_authors)[0].strip()
        # 去掉脚注数字
        first = re.sub(r"[\d\*\u200b]+", "", first).strip()
        if first:
            first_author = first

    # 国别判定（最后一一步）
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
    # DOM 兜底：如果 meta 抓不到标题
    if not fields["title"]:
        try:
            el = page.query_selector("h1.c-article-magazine-title, h1.article-title, h1")
            if el:
                fields["title"] = el.inner_text().strip()
        except Exception:
            pass
    return enrich_fields(fields, html, 'nature', parse_country)


# ---------- 缓存 ----------


def cache_path(article_id: str) -> Path:
    return _common_cache_path(CACHE_DIR, article_id)


def load_from_cache(article_id: str) -> dict | None:
    return _common_load_cache(CACHE_DIR, article_id)


def save_to_cache(article_id: str, fields: dict) -> None:
    _common_save_cache(CACHE_DIR, article_id, fields)


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

    # Nature issue 列表需要 JS 渲染，等一下
    try:
        page.wait_for_selector('a[href*="/articles/s41586-"]', timeout=20000)
    except PWTimeout:
        cb.log("[warn] 未找到 s41586- 文章链接，可能页面结构变化或被拦截。")
        return []

    all_articles = extract_article_list(page)
    targets = all_articles
    check = count_check(page.content(), targets, 'nature', issue_url)
    cb.on_state(check)
    if not check['matched']:
        cb.log(f"[warn] 数量检验：{check['actual']}/{check['expected']}；漏项 {check['missing_ids']}；多项 {check['extra_ids']}")
    else:
        cb.log(f"[check] 数量检验通过：{check['actual']} 篇")
    cb.log(f"[*] 共 {len(targets)} 篇 s41586 文章（全部 section，不过滤）")
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
        m = re.search(r"/articles/(s41586-\d{3}-\d{4,7}-[a-z0-9]+)", url)
        article_id = m.group(1) if m else ""

        cb.on_state({"phase": "article_start", "issue_url": issue_url,
                     "article_idx": i, "article_total": total,
                     "section": section, "url": url, "title": list_title})

        # 缓存命中
        if use_cache and article_id:
            cached = load_from_cache(article_id)
            if cached:
                cached = dict(cached, issue_url=issue_url, section=section)
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
        rwait(1.5, 4.0)
        random_mouse_jitter(page, moves=2)
        if not _safe_goto(page, url, cb, max_retries=2):
            cb.log("        [error] goto 重试均失败，跳过此篇")
            fields = {
                "url": url, "title": "[GOTO FAILED]", "doi": "", "type": section,
                "first_author": "", "first_aff": "", "first_author_country": "",
                "is_china": None, "authors": [],
            }
            results.append((section, url, fields))
            cb.on_state({"phase": "article_done", "url": url,
                         "fields": fields, "blocked": True})
            rwait(5.0, 10.0)
            continue

        # 文章页 CF 检测交给 extract_with_cf_retry 兜底（基于提取结果判定，更精准）

        # 等 SPA 渲染
        try:
            page.wait_for_selector("#Aff1, .c-article-author-affiliation__address", timeout=10000)
        except PWTimeout:
            page.wait_for_timeout(1500)
        human_pause(page, 1.2, 2.8)
        human_scroll(page)

        fields = extract_with_cf_retry(page, url, cb, section,
                                       extract_fields, human_pause, max_retries=5)
        if cb.is_cancelled():
            raise ScrapeCancelled()
        if fields is None:
            cb.log("        [error] 多次重试仍是挑战页，记为 [CF BLOCKED]")
            fields = {
                "url": url, "title": "[CF BLOCKED]", "doi": article_id, "type": section,
                "first_author": "", "first_aff": "", "first_author_country": "",
                "is_china": None, "authors": [],
            }
        if use_cache:
            fields = merge_cached_fields(CACHE_DIR, article_id, fields)
        results.append((section, url, fields))

        cb.log(f"        标题: {fields['title'][:80]}")
        cb.log(f"        DOI:  {fields['doi']}")
        cb.log(f"        类型: {fields['type']}  一作: {fields['first_author']}")
        cb.log(f"        单位: {fields['first_aff'][:120]}")
        cb.log(f"        国家: {fields['first_author_country']}  "
               f"是否中国: {china_label(fields['is_china'])}")

        # 让 cache 携带 section + issue_url，供"立即导出"按 issue 分组、按 section 分类
        fields["section"] = section
        fields["issue_url"] = issue_url
        # 只在拿到真实数据时写缓存
        if article_id and fields.get("title") and fields["title"] not in ("[CF BLOCKED]", "[GOTO FAILED]"):
            save_to_cache(article_id, fields)

        cb.on_state({"phase": "article_done", "url": url,
                     "fields": fields, "cached": False})

        # 文章间随机停顿 + 长歇
        if i % 5 == 0:
            pause = time.time()
            sleep_dur = 45 + (hash(article_id) % 60)
            cb.log(f"        [*] 第 {i} 篇完成，长歇 {sleep_dur}s 模拟阅读间歇...")
            for _ in range(int(sleep_dur)):
                if cb.is_cancelled():
                    break
                time.sleep(1)
        else:
            rwait(6.0, 14.0)

    if cache_hits:
        cb.log(f"\n[issue] 缓存命中 {cache_hits}/{total} 篇（未访问网络）")
    cb.on_state({"phase": "issue_done", "issue_url": issue_url, "count": len(results)})
    return results


# ---------- 主流程 ----------


def default_out_path() -> str:
    return _common_default_out("nature")


def load_urls() -> list[str]:
    return _common_load_urls(URLS_FILE)


def run_scraper(urls: list[str], out_path: str | Path,
                cb: ScraperCallbacks | None = None,
                use_cache: bool = True, headless: bool = False) -> list:
    return run_source(urls, out_path, cb or ScraperCallbacks(), use_cache, headless,
                      source='nature', profile_dir=PROFILE_DIR,
                      playwright_factory=sync_playwright, process_issue=process_issue,
                      columns=NATURE_COLUMNS, col_widths=column_widths(NATURE_COLUMNS))



def main() -> int:
    ap = argparse.ArgumentParser(description="Nature journal issue scraper")
    ap.add_argument("--out", default=None,
                    help=f"output xlsx path (默认: {default_out_path()})")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--fresh", action="store_true", help="忽略缓存，全部重抓")
    args = ap.parse_args()

    out_path = Path(args.out if args.out else default_out_path()).resolve()
    urls = load_urls()
    print(f"[*] 从 urls_nature.txt 读到 {len(urls)} 个 issue URL；use_cache={not args.fresh}")
    run_scraper(urls, out_path, use_cache=not args.fresh, headless=args.headless)
    return 0


if __name__ == "__main__":
    sys.exit(main())
