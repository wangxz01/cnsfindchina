"""Cell 期刊爬虫 Web 服务。

启动：
    python web.py [--host 127.0.0.1] [--port 8000]

然后在浏览器打开 http://127.0.0.1:8000/，所有操作通过网页完成。
CLI 模式（python scraper.py）行为不变。
"""
from __future__ import annotations

import argparse
import json
import queue
import threading
import time
from datetime import date
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

from scraper import (
    ScraperCallbacks,
    URLS_FILE,
    default_out_path,
    run_scraper,
)


# ---------- 全局状态（单用户单进程工具，简单化） ----------


class Session:
    def __init__(self):
        self.lock = threading.Lock()
        self.thread: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.cf_event = threading.Event()
        self.cf_signal_time = 0.0  # 上次 cf_wait 起始时间

        # 状态字段
        self.status: str = "idle"  # idle | running | cf_blocked | done | error | cancelled
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

        # 日志队列：所有 SSE 客户端从这里消费
        # 用 list 维护多个客户端的独立队列，每来一条日志广播给所有
        self._clients: list[queue.Queue] = []
        self._clients_lock = threading.Lock()

    # ---- 客户端管理 ----
    def subscribe(self) -> queue.Queue:
        q: queue.Queue = queue.Queue()
        # 先把当前快照发过去
        q.put({"type": "state", "data": self.snapshot()})
        with self._clients_lock:
            self._clients.append(q)
        return q

    def unsubscribe(self, q: queue.Queue) -> None:
        with self._clients_lock:
            if q in self._clients:
                self._clients.remove(q)

    def _broadcast(self, event: dict) -> None:
        with self._clients_lock:
            for q in self._clients:
                try:
                    q.put_nowait(event)
                except queue.Full:
                    pass

    # ---- 状态变更 ----
    def set_status(self, status: str, **extra) -> None:
        with self.lock:
            self.status = status
            for k, v in extra.items():
                setattr(self, k, v)
        self._broadcast({"type": "state", "data": self.snapshot()})

    def snapshot(self) -> dict:
        with self.lock:
            return {
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
            }


session = Session()
BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"


# ---------- WebCallbacks：把 scraper 事件桥接到 session ----------


class WebCallbacks(ScraperCallbacks):
    def log(self, msg: str) -> None:
        # 也打到 stdout 方便调试
        print(msg)
        session._broadcast({"type": "log", "data": msg})

    def cf_wait(self, target_url: str, current_url: str,
                attempt: int, max_attempts: int) -> bool:
        session.set_status(
            "cf_blocked",
            cf_info={"target": target_url, "current": current_url,
                     "attempt": attempt, "max_attempts": max_attempts},
        )
        # 阻塞等待，直到前端 /api/cf_resumed 触发 cf_event；同时检查取消
        session.cf_event.clear()
        while not session.cf_event.wait(timeout=2.0):
            if session.stop_event.is_set():
                return False
        return not session.stop_event.is_set()

    def is_cancelled(self) -> bool:
        return session.stop_event.is_set()

    def on_state(self, state: dict) -> None:
        phase = state.get("phase")
        if phase == "issue_progress":
            session.set_status(
                "running",
                current_issue=state.get("issue_url", ""),
                issue_idx=state.get("issue_idx", 0),
                issue_total=state.get("issue_total", 0),
            )
        elif phase == "issue_plan":
            # targets 信息可选广播（量大时跳过）
            pass
        elif phase == "article_start":
            session.set_status(
                "running",
                current_article_url=state.get("url", ""),
                current_article_title=state.get("title", ""),
                article_idx=state.get("article_idx", 0),
                article_total=state.get("article_total", 0),
                current_issue=state.get("issue_url", session.current_issue),
            )
        elif phase == "article_done":
            fields = state.get("fields", {})
            session._broadcast({
                "type": "article_done",
                "data": {
                    "url": state.get("url", ""),
                    "title": fields.get("title", ""),
                    "doi": fields.get("doi", ""),
                    "authors": fields.get("authors", []),
                    "first_aff": fields.get("first_aff", ""),
                    "cached": state.get("cached", False),
                    "blocked": state.get("blocked", False),
                },
            })
        elif phase == "excel_written":
            session.set_status("running", out_path=state.get("out_path", ""))
        elif phase == "all_done":
            session.set_status("done", out_path=state.get("out_path", ""))


# ---------- 后台抓取线程 ----------


def _worker(urls: list[str], out_path: str, use_cache: bool):
    cb = WebCallbacks()
    try:
        run_scraper(urls, out_path, cb=cb, use_cache=use_cache, headless=False)
    except Exception as e:
        session.set_status("error", error_msg=str(e)[:300])
        cb.log(f"[error] 异常: {e!r}")
    finally:
        # 标记线程结束
        with session.lock:
            session.thread = None


# ---------- FastAPI ----------


app = FastAPI(title="Cell Scraper")
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def index():
    idx = STATIC_DIR / "index.html"
    if not idx.exists():
        return HTMLResponse("<h1>static/index.html 不存在</h1>", status_code=500)
    return FileResponse(str(idx))


@app.get("/favicon.ico")
def favicon():
    # 内联 SVG 图标（避免浏览器每次请求都 404 刷屏）
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">'
        '<rect width="64" height="64" rx="10" fill="#0071e3"/>'
        '<text x="32" y="44" font-size="38" text-anchor="middle" '
        'fill="#fff" font-family="sans-serif" font-weight="bold">C</text>'
        '</svg>'
    ).encode("utf-8")
    return Response(content=svg, media_type="image/svg+xml")


@app.get("/api/urls.txt")
def get_urls():
    """读取当前 urls.txt 内容。"""
    if not URLS_FILE.exists():
        return {"content": ""}
    return {"content": URLS_FILE.read_text(encoding="utf-8")}


@app.post("/api/urls.txt")
def save_urls(payload: dict):
    """覆盖写入 urls.txt。"""
    content = payload.get("content", "")
    URLS_FILE.write_text(content, encoding="utf-8")
    return {"ok": True}


def _parse_urls_from_text(text: str) -> list[str]:
    urls = []
    for ln in text.splitlines():
        s = ln.strip()
        if not s or s.startswith("#"):
            continue
        urls.append(s)
    return urls


@app.post("/api/start")
def start(payload: dict | None = None):
    """启动抓取。可选 payload: {urls: [...], out: "...", fresh: false}"""
    with session.lock:
        if session.thread is not None and session.thread.is_alive():
            raise HTTPException(status_code=409, detail="已有抓取任务在运行")
        session.thread = None

    # 取 URLs：优先 payload，否则读 urls.txt
    if payload and payload.get("urls"):
        urls = list(payload["urls"])
    else:
        text = URLS_FILE.read_text(encoding="utf-8") if URLS_FILE.exists() else ""
        urls = _parse_urls_from_text(text)

    if not urls:
        raise HTTPException(status_code=400, detail="没有有效的 issue URL")

    out_path = (payload or {}).get("out") or default_out_path()
    out_path = str(Path(out_path).resolve())
    use_cache = not bool((payload or {}).get("fresh", False))

    # 重置 session
    session.stop_event.clear()
    session.cf_event.clear()
    session.set_status(
        "running",
        current_issue="", issue_idx=0, issue_total=len(urls),
        article_idx=0, article_total=0,
        current_article_url="", current_article_title="",
        cf_info={}, out_path=out_path, error_msg="",
    )

    t = threading.Thread(target=_worker, args=(urls, out_path, use_cache), daemon=True)
    with session.lock:
        session.thread = t
    t.start()
    return {"ok": True, "out_path": out_path, "issue_total": len(urls)}


@app.post("/api/cf_resumed")
def cf_resumed():
    """用户在前端点了"我已过 CF 验证"。"""
    session.cf_event.set()
    return {"ok": True}


@app.post("/api/stop")
def stop():
    """请求取消（不会立即停，会在下一个 cb.is_cancelled 检查点生效）。"""
    session.stop_event.set()
    # 如果正卡在 cf_wait，也唤醒它让它退出
    session.cf_event.set()
    session.set_status("cancelled")
    return {"ok": True}


@app.get("/api/status")
def status():
    return session.snapshot()


@app.get("/api/stream")
def stream():
    """SSE 日志+状态流。"""
    q = session.subscribe()

    def gen():
        try:
            while True:
                try:
                    event = q.get(timeout=15)
                except queue.Empty:
                    # 发心跳保活
                    yield ": ping\n\n"
                    continue
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
        finally:
            session.unsubscribe(q)

    return StreamingResponse(gen(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache",
                                      "X-Accel-Buffering": "no"})


@app.get("/api/download")
def download():
    if not session.out_path or not Path(session.out_path).exists():
        # 兜底：找今天的默认文件
        fallback = Path(default_out_path())
        if fallback.exists():
            session.out_path = str(fallback)
        else:
            raise HTTPException(status_code=404, detail="输出文件还未生成")
    name = Path(session.out_path).name
    return FileResponse(session.out_path, filename=name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/reset_cache")
def reset_cache():
    """清空 cache/ 目录。"""
    from scraper import CACHE_DIR
    if CACHE_DIR.exists():
        for f in CACHE_DIR.glob("*.json"):
            try:
                f.unlink()
            except Exception:
                pass
    return {"ok": True}


def main():
    ap = argparse.ArgumentParser(description="Cell scraper web UI")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8000)
    args = ap.parse_args()

    import uvicorn
    print(f"[web] 启动：http://{args.host}:{args.port}/  (Ctrl+C 退出)")
    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
