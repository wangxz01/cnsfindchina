"""统一爬虫 Web 服务：Cell / Nature / Science 三源隔离，可并行运行。

启动：
    python unified_web.py [--host 127.0.0.1] [--port 8000]

浏览器打开 http://127.0.0.1:8000/，顶部切换 source（Cell / Nature / Science）。
三个 source 完全隔离：各自的线程、CF 等待事件、状态、缓存、浏览器 profile，
互不阻塞，可同时启动三个任务并行跑。
"""
from __future__ import annotations

import argparse
import json
import queue
import threading
from pathlib import Path
from typing import Callable

from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

# 三个 source 的具体实现
import scraper as cell_mod
import nature_scraper as nature_mod
import science_scraper as science_mod
from scraper import ScraperCallbacks


# ---------- Source 配置 ----------

class SourceConfig:
    def __init__(self, key: str, label: str,
                 urls_file: Path, cache_dir: Path,
                 default_out: Callable[[], str],
                 run_scraper: Callable, columns, col_widths):
        self.key = key
        self.label = label
        self.urls_file = urls_file
        self.cache_dir = cache_dir
        self.default_out = default_out
        self.run_scraper = run_scraper
        self.columns = columns
        self.col_widths = col_widths


# 9 列 NATURE_COLUMNS 的统一列宽（URL/标题/DOI/类型/一作/单位/国家/中国/作者列表）
DEFAULT_COL_WIDTHS = [55, 55, 25, 14, 18, 60, 14, 10, 50]


SOURCES: dict[str, SourceConfig] = {
    "cell": SourceConfig(
        key="cell", label="Cell",
        urls_file=cell_mod.URLS_FILE,
        cache_dir=cell_mod.CACHE_DIR,
        default_out=cell_mod.default_out_path,
        run_scraper=cell_mod.run_scraper,
        columns=cell_mod.NATURE_COLUMNS,
        col_widths=DEFAULT_COL_WIDTHS,
    ),
    "nature": SourceConfig(
        key="nature", label="Nature",
        urls_file=nature_mod.URLS_FILE,
        cache_dir=nature_mod.CACHE_DIR,
        default_out=nature_mod.default_out_path,
        run_scraper=nature_mod.run_scraper,
        columns=nature_mod.NATURE_COLUMNS,
        col_widths=DEFAULT_COL_WIDTHS,
    ),
    "science": SourceConfig(
        key="science", label="Science",
        urls_file=science_mod.URLS_FILE,
        cache_dir=science_mod.CACHE_DIR,
        default_out=science_mod.default_out_path,
        run_scraper=science_mod.run_scraper,
        columns=science_mod.NATURE_COLUMNS,
        col_widths=DEFAULT_COL_WIDTHS,
    ),
}


# ---------- 事件总线（所有 session 的事件经此转发给 SSE 客户端） ----------


class EventBus:
    def __init__(self):
        self._clients: list[queue.Queue] = []
        self._lock = threading.Lock()

    def subscribe(self) -> queue.Queue:
        q: queue.Queue = queue.Queue()
        # 先把所有 session 当前快照推过去
        for k, s in SESSIONS.items():
            q.put({"type": "state", "source": k, "data": s.snapshot()})
        with self._lock:
            self._clients.append(q)
        return q

    def unsubscribe(self, q) -> None:
        with self._lock:
            if q in self._clients:
                self._clients.remove(q)

    def broadcast(self, event: dict) -> None:
        with self._lock:
            for q in self._clients:
                try:
                    q.put_nowait(event)
                except queue.Full:
                    pass


# ---------- Session（每个 source 独立一份，互不影响） ----------


class Session:
    def __init__(self, source_key: str):
        self.source_key = source_key
        self.lock = threading.Lock()
        self.thread: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.cf_event = threading.Event()

        self.status: str = "idle"
        self.current_issue: str = ""
        self.issue_idx: int = 0
        self.issue_total: int = 0
        self.article_idx: int = 0
        self.article_total: int = 0
        self.current_article_url: str = ""
        self.current_article_title: str = ""
        self.cf_info: dict = {}
        self.out_path: str = ""
        self.error_msg: str = ""

    def set_status(self, status: str, **extra) -> None:
        with self.lock:
            self.status = status
            for k, v in extra.items():
                setattr(self, k, v)
        # 通过 bus 广播（bus 在下方定义）
        bus.broadcast({"type": "state", "source": self.source_key,
                       "data": self.snapshot()})

    def snapshot(self) -> dict:
        with self.lock:
            return {
                "source": self.source_key,
                "status": self.status,
                "current_issue": self.current_issue,
                "issue_idx": self.issue_idx,
                "issue_total": self.issue_total,
                "article_idx": self.article_idx,
                "article_total": self.article_total,
                "current_article_url": self.current_article_url,
                "current_article_title": self.current_article_title,
                "cf_info": self.cf_info,
                "out_path": self.out_path,
                "error_msg": self.error_msg,
                "running": self.thread is not None and self.thread.is_alive(),
            }


# 先建 bus 与 SESSIONS（注意顺序：Session.set_status 引用 bus，需 bus 先于 Session 实例化）
bus = EventBus()
SESSIONS: dict[str, Session] = {k: Session(k) for k in ("cell", "nature", "science")}


# ---------- WebCallbacks：桥接 scraper 事件 → session 状态 + bus ----------


class WebCallbacks(ScraperCallbacks):
    def __init__(self, source_key: str):
        self.source_key = source_key
        self.session = SESSIONS[source_key]

    def log(self, msg: str) -> None:
        print(f"[{self.source_key}] {msg}")
        bus.broadcast({"type": "log", "source": self.source_key, "data": msg})

    def cf_wait(self, target_url: str, current_url: str,
                attempt: int, max_attempts: int) -> bool:
        self.session.set_status(
            "cf_blocked",
            cf_info={"target": target_url, "current": current_url,
                     "attempt": attempt, "max_attempts": max_attempts},
        )
        self.session.cf_event.clear()
        # 0.3s 轮询 stop_event，让"停止"按钮尽快生效（之前是 2s）
        while not self.session.cf_event.wait(timeout=0.3):
            if self.session.stop_event.is_set():
                return False
        return not self.session.stop_event.is_set()

    def is_cancelled(self) -> bool:
        return self.session.stop_event.is_set()

    def on_state(self, state: dict) -> None:
        phase = state.get("phase")
        if phase == "issue_progress":
            self.session.set_status(
                "running",
                current_issue=state.get("issue_url", ""),
                issue_idx=state.get("issue_idx", 0),
                issue_total=state.get("issue_total", 0),
            )
        elif phase == "cf_blocked":
            # wait_until_cf_clear 已在 set_status 处理；这里无需重复
            pass
        elif phase == "article_start":
            self.session.set_status(
                "running",
                current_article_url=state.get("url", ""),
                current_article_title=state.get("title", ""),
                article_idx=state.get("article_idx", 0),
                article_total=state.get("article_total", 0),
                current_issue=state.get("issue_url", self.session.current_issue),
            )
        elif phase == "article_done":
            f = state.get("fields", {})
            bus.broadcast({
                "type": "article_done",
                "source": self.source_key,
                "data": {
                    "url": state.get("url", ""),
                    "title": f.get("title", ""),
                    "doi": f.get("doi", ""),
                    "type": f.get("type", ""),
                    "first_author": f.get("first_author", ""),
                    "first_aff": f.get("first_aff", ""),
                    "first_author_country": f.get("first_author_country", ""),
                    "is_china": f.get("is_china", False),
                    "authors": f.get("authors", []),
                    "cached": state.get("cached", False),
                    "blocked": state.get("blocked", False),
                },
            })
        elif phase == "excel_written":
            self.session.set_status("running", out_path=state.get("out_path", ""))
        elif phase == "all_done":
            self.session.set_status("done", out_path=state.get("out_path", ""),
                                    error_msg="")
        elif phase == "all_skipped":
            # 0 篇成功：标 error 让用户知道有问题，但 out_path 仍设上（可能含 BLOCKED 占位行）
            self.session.set_status(
                "error",
                out_path=state.get("out_path", ""),
                error_msg=state.get("reason", "0 篇文章抓取成功（CF/cookie 墙或结构变化）"),
            )


def _worker(source_key: str, urls, out_path, use_cache):
    cb = WebCallbacks(source_key)
    cfg = SOURCES[source_key]
    session = SESSIONS[source_key]
    try:
        cfg.run_scraper(urls, out_path, cb=cb, use_cache=use_cache, headless=False)
    except Exception as e:
        session.set_status("error", error_msg=str(e)[:300])
        cb.log(f"[error] 异常: {e!r}")
    finally:
        with session.lock:
            session.thread = None
        # 任务结束，再广播一次最终状态（running=False）
        bus.broadcast({"type": "state", "source": source_key,
                       "data": session.snapshot()})


# ---------- FastAPI ----------

BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(title="CNS Scraper Unified")
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def index():
    idx = STATIC_DIR / "unified_index.html"
    if not idx.exists():
        return HTMLResponse("<h1>static/unified_index.html 不存在</h1>", status_code=500)
    return FileResponse(str(idx))


@app.get("/favicon.ico")
def favicon():
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">'
        '<rect width="64" height="64" rx="10" fill="#1d1d1f"/>'
        '<text x="32" y="42" font-size="22" text-anchor="middle" '
        'fill="#fff" font-family="sans-serif" font-weight="bold">CNS</text>'
        '</svg>'
    ).encode("utf-8")
    return Response(content=svg, media_type="image/svg+xml")


@app.get("/api/sources")
def list_sources():
    return {"sources": [
        {"key": k, "label": v.label} for k, v in SOURCES.items()
    ]}


def _resolve_source(source: str | None) -> SourceConfig:
    if source not in SOURCES:
        raise HTTPException(status_code=400, detail=f"未知 source: {source}")
    return SOURCES[source]


@app.get("/api/urls.txt")
def get_urls(source: str = Query(...)):
    cfg = _resolve_source(source)
    if not cfg.urls_file.exists():
        return {"content": ""}
    return {"content": cfg.urls_file.read_text(encoding="utf-8")}


@app.post("/api/urls.txt")
def save_urls(payload: dict, source: str = Query(...)):
    cfg = _resolve_source(source)
    cfg.urls_file.write_text(payload.get("content", ""), encoding="utf-8")
    return {"ok": True}


def _parse_urls(text):
    return [ln.strip() for ln in text.splitlines()
            if ln.strip() and not ln.strip().startswith("#")]


@app.post("/api/start")
def start(payload: dict | None = None, source: str = Query(...)):
    cfg = _resolve_source(source)
    session = SESSIONS[source]
    with session.lock:
        if session.thread is not None and session.thread.is_alive():
            raise HTTPException(
                status_code=409,
                detail=f"{source} 已有任务在运行"
            )

    if payload and payload.get("urls"):
        urls = list(payload["urls"])
    else:
        text = cfg.urls_file.read_text(encoding="utf-8") if cfg.urls_file.exists() else ""
        urls = _parse_urls(text)

    if not urls:
        raise HTTPException(status_code=400, detail=f"{source} 没有有效的 issue URL")

    out_path = (payload or {}).get("out") or cfg.default_out()
    out_path = str(Path(out_path).resolve())
    use_cache = not bool((payload or {}).get("fresh", False))

    session.stop_event.clear()
    session.cf_event.clear()
    session.set_status(
        "running",
        current_issue="", issue_idx=0, issue_total=len(urls),
        article_idx=0, article_total=0,
        current_article_url="", current_article_title="",
        cf_info={}, out_path=out_path, error_msg="",
    )

    t = threading.Thread(target=_worker, args=(source, urls, out_path, use_cache), daemon=True)
    with session.lock:
        session.thread = t
    t.start()
    return {"ok": True, "out_path": out_path, "issue_total": len(urls), "source": source}


@app.post("/api/cf_resumed")
def cf_resumed(source: str = Query(...)):
    """前端按 source 触发：只唤醒对应 session 的 CF 等待。"""
    _resolve_source(source)
    SESSIONS[source].cf_event.set()
    return {"ok": True}


@app.post("/api/stop")
def stop(source: str | None = Query(None)):
    """停止指定 source；不传 source 则停止所有正在运行的。"""
    targets = [source] if source else list(SESSIONS.keys())
    stopped = []
    for k in targets:
        s = SESSIONS[k]
        s.stop_event.set()
        s.cf_event.set()  # 唤醒可能的 cf_wait
        s.set_status("cancelled")
        stopped.append(k)
    return {"ok": True, "stopped": stopped}


@app.get("/api/status")
def status(source: str | None = Query(None)):
    """不传 source 返回所有；传了返回单个。"""
    if source:
        _resolve_source(source)
        return SESSIONS[source].snapshot()
    return {k: v.snapshot() for k, v in SESSIONS.items()}


@app.get("/api/stream")
def stream():
    """SSE：所有 source 的事件都带 source 字段，前端按需过滤。"""
    q = bus.subscribe()

    def gen():
        try:
            while True:
                try:
                    event = q.get(timeout=15)
                except queue.Empty:
                    yield ": ping\n\n"
                    continue
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
        finally:
            bus.unsubscribe(q)

    return StreamingResponse(gen(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.get("/api/download")
def download(source: str = Query(...)):
    _resolve_source(source)
    session = SESSIONS[source]
    out = session.out_path
    if not out or not Path(out).exists():
        # 兜底：找该 source 的默认文件
        fallback = Path(SOURCES[source].default_out())
        if fallback.exists():
            out = str(fallback)
            session.out_path = out
        else:
            raise HTTPException(status_code=404, detail=f"{source} 输出文件还未生成")
    return FileResponse(out, filename=Path(out).name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/reset_cache")
def reset_cache(source: str = Query(...)):
    cfg = _resolve_source(source)
    if cfg.cache_dir.exists():
        for f in cfg.cache_dir.glob("*.json"):
            try: f.unlink()
            except Exception: pass
    return {"ok": True}


@app.post("/api/export")
def export_now(source: str = Query(...)):
    """立即从 cache 重建 Excel 并返回路径（中途手动导出）。

    遍历 cache 目录所有 JSON，按 issue URL 分组（每个 cache 文件含 url 字段），
    写到该 source 默认输出路径。
    """
    cfg = _resolve_source(source)
    session = SESSIONS[source]
    if not cfg.cache_dir.exists() or not any(cfg.cache_dir.glob("*.json")):
        raise HTTPException(status_code=404, detail=f"{source} 缓存为空，无可导出数据")

    # 按 issue_url 分组成多 sheet（与正常 run_scraper 输出一致）；
    # section 优先 fields["section"]，回退 fields["type"]（Nature/Science 已有 type=section），
    # 再回退空。旧 cache 既无 section 也无 issue_url → 全部落到 fallback 单 sheet。
    from collections import defaultdict
    from excel_writer import write_excel
    groups: dict[str, list] = defaultdict(list)
    for f in sorted(cfg.cache_dir.glob("*.json")):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        issue = d.get("issue_url") or ""
        section = d.get("section") or d.get("type") or ""
        groups[issue].append((section, d.get("url", ""), d))

    fallback_label = f"export ({source} cache)"
    all_issues = [(iu or fallback_label, items) for iu, items in groups.items()]

    out_path = session.out_path or cfg.default_out()
    out_path = str(Path(out_path).resolve())
    actual = write_excel(out_path, all_issues,
                         columns=cfg.columns, col_widths=cfg.col_widths)
    session.out_path = actual
    bus.broadcast({"type": "state", "source": source, "data": session.snapshot()})
    return {"ok": True, "out_path": actual, "count": sum(len(it) for it in groups.values())}


def main():
    ap = argparse.ArgumentParser(description="CNS unified scraper web UI")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8000)
    args = ap.parse_args()

    import uvicorn
    print(f"[unified] 启动：http://{args.host}:{args.port}/  (Ctrl+C 退出)")
    print(f"[unified] 三个 source 完全隔离，可并行：{list(SOURCES.keys())}")
    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
