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
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING

from playwright.sync_api import TimeoutError as PWTimeout

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
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


def save_to_cache(cache_dir: Path, article_id: str, fields: dict) -> None:
    cache_dir.mkdir(exist_ok=True)
    cache_path(cache_dir, article_id).write_text(
        json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8"
    )


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
            time.sleep(wait_s)
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
