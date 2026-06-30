"""Nature 期刊爬虫 Web 服务（镜像 web.py，但调用 nature_scraper）。

启动：
    python nature_web.py [--host 127.0.0.1] [--port 8001]

然后在浏览器打开 http://127.0.0.1:8001/。
"""
from __future__ import annotations

import argparse
import json
import queue
import threading
import time
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

import nature_scraper
from nature_scraper import (
    URLS_FILE, CACHE_DIR, default_out_path,
    process_issue, run_scraper,
)
from scraper import ScraperCallbacks


# ---------- Session（单用户单进程） ----------


class Session:
    def __init__(self):
        self.lock = threading.Lock()
        self.thread = None
        self.stop_event = threading.Event()
        self.cf_event = threading.Event()
        self.status = "idle"
        self.current_issue = ""
        self.issue_idx = 0
        self.issue_total = 0
        self.article_idx = 0
        self.article_total = 0
        self.current_article_url = ""
        self.current_article_title = ""
        self.cf_info = {}
        self.out_path = ""
        self.error_msg = ""
        self._clients: list[queue.Queue] = []
        self._clients_lock = threading.Lock()

    def subscribe(self) -> queue.Queue:
        q = queue.Queue()
        q.put({"type": "state", "data": self.snapshot()})
        with self._clients_lock:
            self._clients.append(q)
        return q

    def unsubscribe(self, q):
        with self._clients_lock:
            if q in self._clients:
                self._clients.remove(q)

    def _broadcast(self, event):
        with self._clients_lock:
            for q in self._clients:
                try:
                    q.put_nowait(event)
                except queue.Full:
                    pass

    def set_status(self, status, **extra):
        with self.lock:
            self.status = status
            for k, v in extra.items():
                setattr(self, k, v)
        self._broadcast({"type": "state", "data": self.snapshot()})

    def snapshot(self):
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


class WebCallbacks(ScraperCallbacks):
    def log(self, msg):
        print(msg)
        session._broadcast({"type": "log", "data": msg})

    def cf_wait(self, target_url, current_url, attempt, max_attempts):
        session.set_status(
            "cf_blocked",
            cf_info={"target": target_url, "current": current_url,
                     "attempt": attempt, "max_attempts": max_attempts},
        )
        session.cf_event.clear()
        while not session.cf_event.wait(timeout=2.0):
            if session.stop_event.is_set():
                return False
        return not session.stop_event.is_set()

    def is_cancelled(self):
        return session.stop_event.is_set()

    def on_state(self, state):
        phase = state.get("phase")
        if phase == "issue_progress":
            session.set_status("running",
                current_issue=state.get("issue_url", ""),
                issue_idx=state.get("issue_idx", 0),
                issue_total=state.get("issue_total", 0))
        elif phase == "article_start":
            session.set_status("running",
                current_article_url=state.get("url", ""),
                current_article_title=state.get("title", ""),
                article_idx=state.get("article_idx", 0),
                article_total=state.get("article_total", 0),
                current_issue=state.get("issue_url", session.current_issue))
        elif phase == "article_done":
            f = state.get("fields", {})
            session._broadcast({
                "type": "article_done",
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
            session.set_status("running", out_path=state.get("out_path", ""))
        elif phase == "all_done":
            session.set_status("done", out_path=state.get("out_path", ""))


def _worker(urls, out_path, use_cache):
    cb = WebCallbacks()
    try:
        run_scraper(urls, out_path, cb=cb, use_cache=use_cache, headless=False)
    except Exception as e:
        session.set_status("error", error_msg=str(e)[:300])
        cb.log(f"[error] 异常: {e!r}")
    finally:
        with session.lock:
            session.thread = None


app = FastAPI(title="Nature Scraper")
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
def index():
    idx = STATIC_DIR / "nature_index.html"
    if not idx.exists():
        fallback = STATIC_DIR / "index.html"
        if fallback.exists():
            return FileResponse(str(fallback))
        return HTMLResponse("<h1>static/index.html 不存在</h1>", status_code=500)
    return FileResponse(str(idx))


@app.get("/favicon.ico")
def favicon():
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">'
        '<rect width="64" height="64" rx="10" fill="#1d1d1f"/>'
        '<text x="32" y="44" font-size="38" text-anchor="middle" '
        'fill="#fff" font-family="serif" font-weight="bold" font-style="italic">N</text>'
        '</svg>'
    ).encode("utf-8")
    return Response(content=svg, media_type="image/svg+xml")


@app.get("/api/urls.txt")
def get_urls():
    if not URLS_FILE.exists():
        return {"content": ""}
    return {"content": URLS_FILE.read_text(encoding="utf-8")}


@app.post("/api/urls.txt")
def save_urls(payload: dict):
    URLS_FILE.write_text(payload.get("content", ""), encoding="utf-8")
    return {"ok": True}


def _parse_urls(text):
    return [ln.strip() for ln in text.splitlines()
            if ln.strip() and not ln.strip().startswith("#")]


@app.post("/api/start")
def start(payload: dict | None = None):
    with session.lock:
        if session.thread is not None and session.thread.is_alive():
            raise HTTPException(status_code=409, detail="已有抓取任务在运行")

    if payload and payload.get("urls"):
        urls = list(payload["urls"])
    else:
        text = URLS_FILE.read_text(encoding="utf-8") if URLS_FILE.exists() else ""
        urls = _parse_urls(text)

    if not urls:
        raise HTTPException(status_code=400, detail="没有有效的 issue URL")

    out_path = (payload or {}).get("out") or default_out_path()
    out_path = str(Path(out_path).resolve())
    use_cache = not bool((payload or {}).get("fresh", False))

    session.stop_event.clear()
    session.cf_event.clear()
    session.set_status("running",
        current_issue="", issue_idx=0, issue_total=len(urls),
        article_idx=0, article_total=0,
        current_article_url="", current_article_title="",
        cf_info={}, out_path=out_path, error_msg="")

    t = threading.Thread(target=_worker, args=(urls, out_path, use_cache), daemon=True)
    with session.lock:
        session.thread = t
    t.start()
    return {"ok": True, "out_path": out_path, "issue_total": len(urls)}


@app.post("/api/cf_resumed")
def cf_resumed():
    session.cf_event.set()
    return {"ok": True}


@app.post("/api/stop")
def stop():
    session.stop_event.set()
    session.cf_event.set()
    session.set_status("cancelled")
    return {"ok": True}


@app.get("/api/status")
def status():
    return session.snapshot()


@app.get("/api/stream")
def stream():
    q = session.subscribe()
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
            session.unsubscribe(q)
    return StreamingResponse(gen(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.get("/api/download")
def download():
    if not session.out_path or not Path(session.out_path).exists():
        fallback = Path(default_out_path())
        if fallback.exists():
            session.out_path = str(fallback)
        else:
            raise HTTPException(status_code=404, detail="输出文件还未生成")
    return FileResponse(session.out_path, filename=Path(session.out_path).name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/api/reset_cache")
def reset_cache():
    if CACHE_DIR.exists():
        for f in CACHE_DIR.glob("*.json"):
            try: f.unlink()
            except Exception: pass
    return {"ok": True}


def main():
    ap = argparse.ArgumentParser(description="Nature scraper web UI")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8001)
    args = ap.parse_args()

    import uvicorn
    print(f"[nature-web] 启动：http://{args.host}:{args.port}/  (Ctrl+C 退出)")
    uvicorn.run(app, host=args.host, port=args.port, log_level="info")


if __name__ == "__main__":
    main()
