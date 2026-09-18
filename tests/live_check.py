"""Visible-browser check of one real article, with manual verification handoff.

Run: uv run python tests/live_check.py cell
Publisher HTML and results stay in ignored data/live_checks/ for diagnosis.
"""
from __future__ import annotations

import argparse
import json
import sys
import queue
import threading
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))

from playwright.sync_api import sync_playwright
import scraper
import nature_scraper
import science_scraper
from scraper_common import DATA_DIR, atomic_json
from excel_writer import write_excel

SAMPLES = {
    'cell': (scraper, 'https://www.sciencedirect.com/science/article/pii/S0092867426003946'),
    'nature': (nature_scraper, 'https://www.nature.com/articles/s41586-026-10614-4'),
    'science': (science_scraper, 'https://www.science.org/doi/10.1126/science.aeb5171'),
}


def wait_for_command(page):
    """Keep Playwright processing browser events during the manual handoff."""
    commands = queue.Queue()

    def read():
        try:
            commands.put(input('Enter=extract; q=close > ').strip().lower())
        except EOFError:
            commands.put('q')

    threading.Thread(target=read, daemon=True).start()
    while True:
        try:
            return commands.get_nowait()
        except queue.Empty:
            try:
                page.wait_for_timeout(200)
            except Exception:
                print('BROWSER_CLOSED: 测试浏览器已关闭。', flush=True)
                return 'q'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('source', choices=SAMPLES)
    parser.add_argument('--url')
    parser.add_argument('--auto', action='store_true', help='Extract immediately if no challenge is detected')
    parser.add_argument('--show-more', action='store_true', help='Exercise the Cell Show more control')
    args = parser.parse_args()
    module, default_url = SAMPLES[args.source]
    url = args.url or default_url
    folder = DATA_DIR / 'live_checks' / args.source
    folder.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(module.PROFILE_DIR), headless=False,
            viewport={'width': 1366, 'height': 900},
            args=['--disable-blink-features=AutomationControlled'])
        try:
            page = context.new_page()
            try:
                page.goto(url, wait_until='domcontentloaded', timeout=60000)
            except Exception as exc:
                print(f'Navigation notice: {exc}', flush=True)
            print(f'BROWSER_READY {args.source}: {url}', flush=True)
            auto = args.auto and not scraper.is_cloudflare(page) and not scraper._is_challenge_title(page.title())
            if not auto:
                print('HUMAN_CHECK_REQUIRED: 请在弹出的 Chromium 窗口通过验证；确认文章可见后继续。', flush=True)
            while True:
                command = '' if auto else wait_for_command(page)
                auto = False
                if command == 'q':
                    break
                (folder / 'before.html').write_text(page.content(), encoding='utf-8')
                if args.source == 'cell' and args.show_more:
                    print(f'SHOW_MORE_CLICKED {module._click_show_more_js(page)}', flush=True)
                    page.wait_for_timeout(1000)
                fields = module.extract_fields(page, url)
                (folder / 'after.html').write_text(page.content(), encoding='utf-8')
                atomic_json(folder / 'fields.json', fields)
                out = write_excel(str(folder / 'result.xlsx'),
                                  [(url, [(fields.get('type') or '文章实测', url, fields)])],
                                  columns=module.NATURE_COLUMNS)
                print('EXTRACTED ' + json.dumps({key: fields.get(key) for key in (
                    'title', 'first_author', 'first_aff', 'is_china', 'available_online',
                    'version_of_record', 'published_date', 'date_evidence',
                    'extraction_status', 'missing_fields')}, ensure_ascii=False), flush=True)
                print(f'OUTPUT {out}', flush=True)
        finally:
            try:
                context.close()
            except Exception:
                pass  # The user may already have closed the visible window.


if __name__ == '__main__':
    main()
