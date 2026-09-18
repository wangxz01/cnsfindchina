"""三套 scraper 共用的辅助函数。

历史上 scraper.py / nature_scraper.py / science_scraper.py 各自维护了一份
近乎一致的 cache_path / load_from_cache / save_to_cache / load_urls /
default_out_path / _safe_goto，~150 行重复。本模块集中这些与具体期刊无关的
工具；三个 scraper 文件改成调用这里。

保留各 scraper 的 process_issue / run_scraper / extract_article_list /
_extract_fields_from_html 不变——它们包含期刊特定的选择器和节奏参数。
"""
from __future__ import annotations

import json
import re
import sys
import time
import os
import tempfile
import threading
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING

from playwright.sync_api import TimeoutError as PWTimeout
from article_metadata import finalize_fields, FAILED_TITLES

if TYPE_CHECKING:
    # 仅类型注解用；运行时不 import 以避免 scraper.py ↔ scraper_common.py 循环
    from scraper import ScraperCallbacks


# 强制 stdout/stderr 用 utf-8，避免 Windows 默认 cp932/cp936 终端
# print 中文日志时抛 UnicodeEncodeError 或乱码。Python 3.7+ 支持 reconfigure。
# 放在 scraper_common 而非各 scraper，因为所有模块都 import 它，一处覆盖全部。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass


# ---------- 项目目录布局 ----------
# 项目根 = 本文件所在目录的上一级（src/ 的父目录）
PROJECT_ROOT = Path(__file__).parent.parent
# 所有运行时 I/O（urls / 缓存 / 浏览器 profile / 输出 xlsx）都放 data/
DATA_DIR = PROJECT_ROOT / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
CACHE_VERSION = 2
_active = threading.local()


class ScrapeCancelled(Exception):
    pass


def cancellable_sleep(seconds, cb=None):
    cb = cb or getattr(_active, 'cb', None)
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if cb and cb.is_cancelled():
            raise ScrapeCancelled()
        time.sleep(min(0.2, max(0, deadline - time.monotonic())))


def atomic_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(dir=path.parent, suffix='.tmp')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
        os.replace(name, path)
    finally:
        if os.path.exists(name):
            os.unlink(name)


# ---------- 缓存（按文章 ID；PII 或 DOI） ----------


def cache_path(cache_dir: Path, article_id: str) -> Path:
    """返回 cache_dir/<id>.json；id 含 '/'（如 DOI）自动替换为 '_'。"""
    safe = article_id.replace("/", "_")
    return cache_dir / f"{safe}.json"


def load_from_cache(cache_dir: Path, article_id: str) -> dict | None:
    p = cache_path(cache_dir, article_id)
    if not p.exists():
        return None
    try:
        fields = json.loads(p.read_text(encoding="utf-8"))
        if not isinstance(fields, dict) or fields.get('cache_version') != CACHE_VERSION:
            return None
        source = fields.get('source')
        if source not in ('cell', 'nature', 'science'):
            return None
        finalize_fields(fields, source)
        return fields if fields['extraction_status'] == 'complete' else None
    except Exception:
        return None


def save_to_cache(cache_dir: Path, article_id: str, fields: dict) -> None:
    # Partial records remain exportable, but are never accepted as complete cache hits.
    fields['cache_version'] = CACHE_VERSION
    atomic_json(cache_path(cache_dir, article_id), fields)


def merge_cached_fields(cache_dir, article_id, fields):
    """Retry partial articles without losing previously verified metadata."""
    if fields.get('title') in FAILED_TITLES:
        return fields
    try:
        old = json.loads(cache_path(cache_dir, article_id).read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return fields
    if not isinstance(old, dict) or old.get('cache_version') != CACHE_VERSION:
        return fields
    for key in ('title', 'doi', 'authors', 'first_author', 'available_online', 'version_of_record', 'published_date'):
        if not fields.get(key) and old.get(key):
            fields[key] = old[key]
    fields['date_evidence'] = {**old.get('date_evidence', {}), **fields.get('date_evidence', {})}
    if not fields.get('affiliation_verified') and old.get('affiliation_verified') and fields.get('first_author') == old.get('first_author'):
        for key in ('first_aff', 'first_author_affiliations', 'affiliation_verified', 'first_author_countries', 'first_author_country', 'is_china'):
            fields[key] = old.get(key)
    return finalize_fields(fields, fields['source'])


# ---------- urls 文件加载 ----------


def load_urls(urls_file: Path) -> list[str]:
    """读 urls 文件，# 注释和空行跳过；空文件 sys.exit(2)。"""
    if not urls_file.exists():
        print(f"[error] 未找到 {urls_file}。")
        sys.exit(2)
    urls = []
    for ln in urls_file.read_text(encoding="utf-8").splitlines():
        s = ln.strip()
        if not s or s.startswith("#"):
            continue
        urls.append(s)
    if not urls:
        print(f"[error] {urls_file} 中没有有效 URL。")
        sys.exit(2)
    return urls


# ---------- 默认输出路径 ----------


def default_out_path(prefix: str) -> str:
    """按日期命名，放在 data/ 目录下：data/<prefix>_YYYY-MM-DD.xlsx"""
    return str(DATA_DIR / f"{prefix}_{date.today().isoformat()}.xlsx")


# ---------- goto 重试（Science 启用，其它可选用） ----------


def safe_goto(page, url: str, cb: ScraperCallbacks, max_retries: int = 2) -> bool:
    """goto 重试：连续访问容易被限速/超时。成功返回 True。"""
    for attempt in range(max_retries):
        if cb.is_cancelled():
            raise ScrapeCancelled()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            return True
        except PWTimeout:
            cb.log(f"        [warn] goto 超时（尝试 {attempt+1}/{max_retries}）")
        except Exception as e:
            msg = str(e)
            if "ERR_ABORTED" in msg or "net::ERR_" in msg:
                cb.log(f"        [warn] goto 中断（尝试 {attempt+1}/{max_retries}, {msg[:60]}）")
            else:
                cb.log(f"        [warn] goto 异常（尝试 {attempt+1}/{max_retries}）: {msg[:100]}")
        if attempt < max_retries - 1:
            wait_s = 15 * (attempt + 1)
            cb.log(f"        [*] 等待 {wait_s}s 后重试...")
            cancellable_sleep(wait_s, cb)
    return False


# ---------- 一作不在 wanted_sections 时的诊断小工具 ----------


def count_real_articles(all_issues: list) -> int:
    """统计非 [CF BLOCKED] / [GOTO FAILED] 的真实文章数。

    供 run_scraper 判定 all_done vs all_skipped 使用。
    """
    return sum(
        1 for _, results in all_issues
        for _, _, f in results
        if f.get("title") and f["title"] not in ("[CF BLOCKED]", "[GOTO FAILED]")
    )


def run_source(urls, out_path, cb, use_cache, headless, *, source, profile_dir,
               playwright_factory, process_issue, columns, col_widths):
    """Shared lifecycle, preserving partial rows and a truthful terminal status."""
    from excel_writer import write_excel
    out_path = str(out_path)
    all_issues, issue_meta = [], {}
    current_rows = []
    current_url = ''
    _active.cb = cb

    class TrackingCallbacks:
        def __getattr__(self, name):
            return getattr(cb, name)

        def on_state(self, state):
            phase = state.get('phase')
            meta = issue_meta.setdefault(current_url, {})
            if phase == 'issue_plan':
                meta['expected'] = len(state['targets'])
            elif phase == 'count_check':
                meta['count_check'] = state
            elif phase == 'article_done':
                fields = state['fields']
                finalize_fields(fields, source)
                fields['issue_url'] = current_url
                current_rows.append((fields.get('section') or fields.get('type', ''), state['url'], fields))
                state['blocked'] = fields['extraction_status'] == 'failed'
            cb.on_state(state)

    tracker = TrackingCallbacks()
    try:
        with playwright_factory() as p:
            ctx = p.chromium.launch_persistent_context(
                user_data_dir=str(profile_dir), headless=headless,
                viewport={'width': 1366, 'height': 900},
                args=['--disable-blink-features=AutomationControlled'])
            try:
                page = ctx.new_page()
                for idx, current_url in enumerate(urls, 1):
                    if cb.is_cancelled():
                        break
                    current_rows = []
                    issue_meta[current_url] = {}
                    tracker.on_state(dict(phase='issue_progress', issue_idx=idx, issue_total=len(urls), issue_url=current_url))
                    try:
                        rows = process_issue(page, current_url, use_cache=use_cache, cb=tracker)
                        current_rows = rows
                    except ScrapeCancelled:
                        pass
                    except Exception as exc:
                        issue_meta[current_url]['error'] = str(exc)[:300]
                        cb.log(f'[error] issue 抓取失败，保留已完成文章: {exc}')
                    if not current_rows and not cb.is_cancelled():
                        issue_meta[current_url].setdefault('error', '未取得文章列表或文章数据')
                    all_issues.append((current_url, current_rows))
                    out_path = write_excel(out_path, all_issues, columns=columns,
                                           col_widths=col_widths, issue_meta=issue_meta)
                    tracker.on_state(dict(phase='excel_written', out_path=out_path))
                    if idx < len(urls) and not cb.is_cancelled():
                        try:
                            cancellable_sleep(30, cb)
                        except ScrapeCancelled:
                            break
            finally:
                ctx.close()
    finally:
        _active.cb = None
    rows = [fields for _, results in all_issues for _, _, fields in results]
    real = sum(bool(f.get('title')) and f['title'] not in FAILED_TITLES for f in rows)
    incomplete = (len(all_issues) < len(urls) or any(f.get('extraction_status') != 'complete' for f in rows)
                  or any(m.get('error') or not m.get('count_check', {}).get('matched', False) for m in issue_meta.values()))
    phase = ('cancelled' if cb.is_cancelled() else 'all_skipped' if not real else
             'partial_done' if incomplete else 'all_done')
    cb.on_state(dict(phase=phase, out_path=out_path if all_issues else '',
                     reason='部分数据缺失或数量校验未通过' if incomplete else ''))
    cb.log(f'[{phase}] 已保存 {len(rows)} 篇记录，成功访问 {real} 篇')
    return all_issues
