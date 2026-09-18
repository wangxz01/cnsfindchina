"""Read publisher archive links, including paginated years and lazy panels."""
from __future__ import annotations

import re
import threading
import time
from copy import deepcopy
from urllib.parse import urljoin, urlsplit

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from article_metadata import Document, clean, normalize_date
from scraper import is_cloudflare

ARCHIVE_URL = 'https://www.sciencedirect.com/journal/cell/issues'
ARCHIVE_URLS = dict(cell=ARCHIVE_URL, nature='https://www.nature.com/nature/volumes',
                    science='https://www.science.org/loi/science')
LABELS = dict(cell='Cell', nature='Nature', science='Science')
YEAR_RE = re.compile(r'^\s*(\d{4})\s*[—–-]\s*Volumes?\s', re.I)
CELL_PATH = r'/journal/cell/vol/(\d+)/(?:issue/(\d+(?:-\d+)*)|suppl/([A-Za-z0-9-]+))/?'


ISSUE_PATHS = dict(cell=CELL_PATH, nature=r'/nature/volumes/(\d+)/issues/(\d+)/?',
                   science=r'/toc/science/((?:os-)?\d+)/(\d+)/?')


def canonical_issue_url(value, source='cell'):
    parts = urlsplit(urljoin(ARCHIVE_URLS[source], value))
    if parts.scheme == 'https' and parts.netloc == urlsplit(ARCHIVE_URLS[source]).netloc and re.fullmatch(ISSUE_PATHS[source], parts.path):
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


def parse_nature_years(html):
    found = {}
    heading_year = None
    for node in Document(html).root.walk():
        if node.tag == 'h3' and re.fullmatch(r'\d{4}', clean(node.text())):
            heading_year = int(clean(node.text()))
        if node.tag != 'a':
            continue
        parts = urlsplit(urljoin(ARCHIVE_URLS['nature'], node.attrs.get('href', '')))
        match = re.fullmatch(r'/nature/volumes/(\d+)/?', parts.path)
        if parts.netloc != 'www.nature.com' or parts.scheme != 'https' or not match:
            continue
        period = clean(node.parent.text())
        # Old volumes span two calendar years (e.g. Nov 1869 - Apr 1870).
        # Associate the volume with both, then filter issues by their actual date.
        dates = sorted({int(y) for y in re.findall(r'\b(?:18|19|20|21)\d{2}\b', period)})
        if not dates and heading_year:
            dates = [heading_year]
        if not dates:
            raise RuntimeError(f'无法确定 Nature 卷 {match[1]} 的年份')
        for year in range(dates[0], dates[-1] + 1):
            item = found.setdefault(year, dict(year=year, label=f'{year} 年', page_url=ARCHIVE_URLS['nature'], volumes=[]))
            volume = dict(volume=int(match[1]), url=f'https://www.nature.com{parts.path.rstrip("/")}', period=period)
            if not any(v['url'] == volume['url'] for v in item['volumes']):
                item['volumes'].append(volume)
    return sorted(found.values(), key=lambda y: y['year'], reverse=True)


def parse_nature_issues(html, expected_volume):
    root = Document(html).root
    container = next((n for n in root.walk() if n.attrs.get('id') == 'issue-list'), None)
    rows, seen = [], set()
    if container is None:
        return rows
    for node in container.walk():
        if node.tag != 'a' or '/issues/' not in node.attrs.get('href', ''):
            continue
        url = canonical_issue_url(node.attrs['href'], 'nature')
        if not url:
            raise RuntimeError('Nature 目录出现未知期号链接，请核对官网。')
        match = re.fullmatch(ISSUE_PATHS['nature'], urlsplit(url).path)
        if int(match[1]) != int(expected_volume):
            raise RuntimeError('Nature 返回的卷号与请求不一致，请重新读取目录。')
        if url in seen:
            continue
        date = normalize_date(clean(node.text()))
        if not date:
            raise RuntimeError(f'Nature 第 {match[2]} 期缺少可识别的期次日期，无法按年份归类。')
        seen.add(url)
        rows.append(dict(url=url, year=int(date[:4]), volume=int(match[1]), issue=match[2],
                         label=f'Volume {match[1]}, Issue {match[2]}', detail=clean(node.text())))
    return rows


def parse_science_years(html):
    years = {}
    for node in Document(html).root.walk():
        if node.tag != 'a':
            continue
        parts = urlsplit(urljoin(ARCHIVE_URLS['science'], node.attrs.get('href', '')))
        match = re.fullmatch(r'/loi/science/group/d\d{4}\.y(\d{4})', parts.path)
        if parts.scheme == 'https' and parts.netloc == 'www.science.org' and match:
            year = int(match[1])
            years[year] = dict(year=year, label=f'{year} 年', page_url=f'https://www.science.org{parts.path}')
    return sorted(years.values(), key=lambda y: y['year'], reverse=True)


def parse_science_issues(html, year):
    root = Document(html).root
    rows, seen = [], set()
    for panel in root.walk():
        if panel.attrs.get('role') != 'tabpanel' or not panel.attrs.get('id', '').endswith(f'-y{year}'):
            continue
        for node in panel.walk():
            if node.tag != 'a' or not node.has_class('past-issue'):
                continue
            url = canonical_issue_url(node.attrs.get('href', ''), 'science')
            if not url:
                raise RuntimeError('Science 目录出现未知期号链接，请核对官网。')
            if url in seen:
                continue
            seen.add(url)
            match = re.fullmatch(ISSUE_PATHS['science'], urlsplit(url).path)
            rows.append(dict(url=url, year=year, volume=match[1], issue=match[2],
                             label=f'Volume {match[1]}, Issue {match[2]}', detail=clean(node.text())))
    return rows


def read_other_catalog(page, source, mode, selected_years, known_years, stop, update):
    def navigate(url, parse):
        if stop.is_set():
            raise CatalogCancelled()
        try:
            page.goto(url, wait_until='domcontentloaded', timeout=30000)
        except PlaywrightTimeout:
            pass
        return _wait(page, lambda: parse(page.content()), stop, update)

    if mode == 'years':
        parser = parse_nature_years if source == 'nature' else parse_science_years
        years = navigate(ARCHIVE_URLS[source], parser)
        update(years=years, years_complete=True,
               message=f'已读取目录中的 {len(years)} 个年份，请勾选年份后读取期号。')
        return

    known = {y['year']: y for y in known_years}
    volumes = {}
    for year in sorted(selected_years, reverse=True):
        if stop.is_set():
            raise CatalogCancelled()
        if source == 'science':
            update(message=f'正在读取 Science {year} 年期号…')
            rows = navigate(known[year]['page_url'], lambda html: parse_science_issues(html, year))
        else:
            rows = []
            expected = known[year]['volumes']
            for index, volume in enumerate(expected, 1):
                if stop.is_set():
                    raise CatalogCancelled()
                update(message=f'正在读取 Nature {year} 年 · Volume {volume["volume"]}（{index}/{len(expected)} 卷）…')
                if volume['url'] not in volumes:
                    volumes[volume['url']] = navigate(volume['url'], lambda html: parse_nature_issues(html, volume['volume']))
                rows.extend(row for row in volumes[volume['url']] if row['year'] == year)
            rows = list({row['url']: row for row in rows}.values())
            if not rows:
                raise RuntimeError(f'未读到 Nature {year} 年的期号，请刷新目录后重试。')
        update(issue_year=str(year), issue_rows=rows)
    update(message=f'已读取所选 {len(selected_years)} 个年份的期号，请勾选后追加。')


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
                update(status='waiting', message='请在该期刊打开的浏览器中完成人机验证或 Cookie 提示；通过后会自动继续（最多等待 5 分钟）。')
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


def read_catalog(mode, selected_years, known_years, profile_dir, stop, update, source='cell'):
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir), headless=False,
            viewport={'width': 1366, 'height': 900},
            args=['--disable-blink-features=AutomationControlled'])
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.set_default_timeout(10000)
            if source != 'cell':
                read_other_catalog(page, source, mode, selected_years, known_years, stop, update)
                return
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
    def __init__(self, source='cell'):
        self.source = source
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
            self.update(status='reading', error='', message=f'正在打开 {LABELS[self.source]} 官方目录…')
            if mode == 'years':
                self.update(years_complete=False)
            known = deepcopy(self.data['years'])

            def worker():
                try:
                    read_catalog(mode, years, known, profile_dir, self.stop_event, self.update, source=self.source)
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
