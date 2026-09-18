"""Read Cell's actual archive links, including paginated years and lazy panels."""
from __future__ import annotations

import re
import threading
import time
from copy import deepcopy
from urllib.parse import urljoin, urlsplit

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from article_metadata import Document, clean
from scraper import is_cloudflare

ARCHIVE_URL = 'https://www.sciencedirect.com/journal/cell/issues'
YEAR_RE = re.compile(r'^\s*(\d{4})\s*[—–-]\s*Volumes?\s', re.I)
CELL_PATH = r'/journal/cell/vol/(\d+)/(?:issue/(\d+(?:-\d+)*)|suppl/([A-Za-z0-9-]+))/?'


def canonical_issue_url(value):
    parts = urlsplit(urljoin(ARCHIVE_URL, value))
    if parts.scheme == 'https' and parts.netloc == 'www.sciencedirect.com' and re.fullmatch(CELL_PATH, parts.path):
        return f'https://{parts.netloc}{parts.path.rstrip("/")}'
    return ''


def parse_archive(html, page_url=ARCHIVE_URL):
    root = Document(html).root
    years, issues, unsupported = [], {}, []
    nodes = list(root.walk())
    by_id = {n.attrs['id']: n for n in nodes if n.attrs.get('id')}
    for button in nodes:
        match = YEAR_RE.match(clean(button.text())) if button.tag == 'button' else None
        if not match:
            continue
        year = int(match[1])
        years.append(dict(year=year, label=clean(button.text()), page_url=page_url))
        panel = by_id.get(button.attrs.get('aria-controls'))
        if panel is None:
            continue
        rows, seen = [], set()
        for anchor in panel.walk():
            if anchor.tag != 'a' or not anchor.has_class('js-issue-item-link'):
                continue
            url = canonical_issue_url(anchor.attrs.get('href', ''))
            if not url:
                unsupported.append(f'{year}: {clean(anchor.text())}')
                continue
            if url in seen:
                continue
            seen.add(url)
            match = re.fullmatch(CELL_PATH, urlsplit(url).path)
            item = anchor
            while item.parent and item is not panel and not item.has_class('issue-item'):
                item = item.parent
            label, detail = clean(anchor.text()), clean(item.text())
            if detail.startswith(label):
                detail = detail[len(label):].strip()
            # Keep the catalogue's date text, including historical month-only dates.
            rows.append(dict(url=url, year=year, volume=int(match[1]),
                             issue=match[2] or f'Supplement {match[3]}',
                             label=label, detail=detail))
        if rows:
            issues[str(year)] = rows
    return years, issues, unsupported


class CatalogCancelled(Exception):
    pass


def _wait(page, condition, stop, update, timeout=35):
    """Pump Playwright while waiting, so manual challenges stay responsive."""
    deadline = time.monotonic() + timeout
    challenge_deadline = None
    while True:
        if stop.is_set():
            raise CatalogCancelled()
        blocked = is_cloudflare(page)
        if blocked:
            if challenge_deadline is None:
                challenge_deadline = time.monotonic() + 300
                update(status='waiting', message='请在打开的 Cell 浏览器中完成人机验证或 Cookie 提示；通过后会自动继续（最多等待 5 分钟）。')
            if time.monotonic() > challenge_deadline:
                raise RuntimeError('等待验证超时，请重新读取目录。')
            deadline = time.monotonic() + timeout
        else:
            if challenge_deadline is not None:
                update(status='reading', message='验证已通过，继续读取目录…')
                challenge_deadline = None
            result = condition()
            if result:
                return result
            if time.monotonic() > deadline:
                raise RuntimeError('目录内容未按预期加载，请重试；若持续失败，网站结构可能已变化。')
        page.wait_for_timeout(300)


def _read_year(page, button, year, target, stop, update):
    clicks, last_click = 0, float('-inf')

    def loaded():
        nonlocal clicks, last_click
        _, rows, unsupported = parse_archive(page.content(), target)
        unsupported = [label for label in unsupported if label.startswith(f'{year}:')]
        if str(year) in rows or unsupported:
            return rows, unsupported
        # Hydration can discard an early click. Only retry after observing that
        # this exact year is still collapsed; never toggle an expanded panel shut.
        if (button.get_attribute('aria-expanded') != 'true'
                and clicks < 3 and time.monotonic() - last_click >= 2):
            button.click()
            clicks += 1
            last_click = time.monotonic()
        return None

    rows, unsupported = _wait(page, loaded, stop, update)
    if unsupported:
        raise RuntimeError('目录包含暂不支持的期次，请到官网核对：' + '；'.join(unsupported[:5]))
    return rows[str(year)]


def read_catalog(mode, selected_years, known_years, profile_dir, stop, update):
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir), headless=False,
            viewport={'width': 1366, 'height': 900},
            args=['--disable-blink-features=AutomationControlled'])
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.set_default_timeout(10000)
            buttons = page.locator('.accordion-panel-title').filter(has_text=YEAR_RE)

            def navigate(url):
                if stop.is_set():
                    raise CatalogCancelled()
                try:
                    page.goto(url, wait_until='domcontentloaded', timeout=30000)
                except PlaywrightTimeout:
                    pass  # A challenge may still be loading; let the user complete it.
                _wait(page, lambda: buttons.count() and buttons.first.is_visible(), stop, update)

            if mode == 'years':
                navigate(ARCHIVE_URL)
                found = {}
                for page_number in range(1, 11):
                    years, _, _ = parse_archive(page.content(), page.url)
                    if not years or all(y['year'] in found for y in years):
                        raise RuntimeError('目录翻页未取得新年份，请重试。')
                    found.update({y['year']: y for y in years})
                    update(years=sorted(found.values(), key=lambda y: y['year'], reverse=True),
                           message=f'已读取 {len(found)} 个年份，正在检查更早年份…')
                    nxt = page.get_by_role('button', name='Next page', exact=True)
                    if not nxt.count() or not nxt.is_visible() or nxt.is_disabled():
                        update(years_complete=True, message=f'已读取目录中的 {len(found)} 个年份，请勾选年份后读取期号。')
                        break
                    # The observed archive pagination uses ?page=2, ?page=3, etc.
                    # Navigate directly after checking Next exists; a click can be
                    # ignored before the server-rendered React page is hydrated.
                    navigate(f'{ARCHIVE_URL}?page={page_number + 1}')
                else:
                    raise RuntimeError('目录页数超过预期，已保留读到的年份，请核对官网。')
            else:
                by_year = {y['year']: y for y in known_years}
                current_url = None
                for year in sorted(selected_years, reverse=True):
                    update(message=f'正在读取 {year} 年期号…')
                    target = by_year[year]['page_url']
                    if current_url != target:
                        navigate(target)
                        current_url = target
                    button = buttons.filter(has_text=re.compile(rf'^\s*{year}\s*[—–-]'))
                    if button.count() != 1:
                        raise RuntimeError(f'{year} 年的目录位置已变化，请重新读取年份目录。')
                    rows = _read_year(page, button, year, target, stop, update)
                    update(issue_year=str(year), issue_rows=rows)
                update(message=f'已读取所选 {len(selected_years)} 个年份的期号，请勾选后追加。')
        finally:
            ctx.close()


class CatalogSession:
    def __init__(self):
        self.lock = threading.RLock()
        self.stop_event = threading.Event()
        self.thread = None
        self.data = dict(status='idle', message='', error='', years=[], years_complete=False, issues={})

    def snapshot(self):
        with self.lock:
            return dict(deepcopy(self.data), running=self.thread is not None)

    def update(self, **values):
        with self.lock:
            if 'issue_year' in values:
                self.data['issues'][values.pop('issue_year')] = values.pop('issue_rows')
            self.data.update(values)

    def start(self, mode, years, profile_dir):
        # The web caller serializes admission with scraper starts under bus._lock.
        with self.lock:
            if self.thread is not None:
                raise RuntimeError('目录正在读取中')
            self.stop_event.clear()
            self.update(status='reading', error='', message='正在打开 Cell 官方目录…')
            if mode == 'years':
                self.update(years_complete=False)
            known = deepcopy(self.data['years'])

            def worker():
                try:
                    read_catalog(mode, years, known, profile_dir, self.stop_event, self.update)
                    self.update(status='done')
                except CatalogCancelled:
                    self.update(status='cancelled', message='已取消读取；已读到的目录可继续使用。')
                except Exception as exc:
                    self.update(status='error', error=str(exc)[:500], message='目录读取失败；已读到的内容已保留，可重试。')
                finally:
                    with self.lock:
                        if self.stop_event.is_set():
                            self.update(status='cancelled', message='已取消读取；已读到的目录可继续使用。', error='')
                        self.thread = None

            self.thread = threading.Thread(target=worker, daemon=True)
            try:
                self.thread.start()
            except Exception:
                self.thread = None
                self.update(status='error', error='无法启动目录读取线程')
                raise

    def cancel(self):
        with self.lock:
            if self.thread is not None:
                self.stop_event.set()
                self.update(status='stopping', message='正在取消读取并关闭目录浏览器…')
