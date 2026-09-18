import sys
import threading
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
import issue_catalog as catalog
import unified_web as web
from excel_writer import _sheet_name_from_url


class CatalogTests(unittest.TestCase):
    def fixture(self, name):
        return (Path(__file__).parent / 'fixtures' / name).read_text(encoding='utf-8')

    def test_nature_cross_year_volume_is_available_to_both_years(self):
        years = catalog.parse_nature_years(self.fixture('nature_archive.html'))
        self.assertEqual([y['year'] for y in years], [2026, 1870, 1869])
        self.assertEqual([v['volume'] for v in years[1]['volumes']], [2, 1])
        rows = catalog.parse_nature_issues(self.fixture('nature_archive_volume.html'), 1)
        self.assertEqual([(r['year'], r['issue']) for r in rows], [(1870, '26'), (1869, '9')])
        with self.assertRaises(RuntimeError): catalog.parse_nature_issues(self.fixture('nature_archive_volume.html'), 2)
        with self.assertRaises(RuntimeError):
            catalog.parse_nature_issues('<ul id="issue-list"><a href="/nature/volumes/1/issues/1">No. 1</a></ul>', 1)

    def test_science_scopes_issues_to_year_and_preserves_old_series(self):
        html = self.fixture('science_archive.html')
        self.assertEqual([y['year'] for y in catalog.parse_science_years(html)], [2026, 1880])
        rows = catalog.parse_science_issues(html, 1880)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['volume'], 'os-1')
        self.assertEqual(web.validate_urls([rows[0]['url']], 'science'), [rows[0]['url']])
        self.assertEqual(_sheet_name_from_url(rows[0]['url']), 'vos-1-i27')
        self.assertEqual(catalog.parse_science_issues(html, 1881), [])

    def test_source_urls_cannot_cross_journals(self):
        samples = {'nature':'https://www.nature.com/nature/volumes/657/issues/8132',
                   'science':'https://www.science.org/toc/science/393/6817'}
        for source, url in samples.items():
            self.assertEqual(catalog.canonical_issue_url(url, source), url)
            self.assertEqual(catalog.canonical_issue_url(url, 'cell'), '')
            with self.assertRaises(web.HTTPException): web.validate_urls([url], 'cell')

    def test_nature_reader_filters_cross_year_and_reuses_shared_volume(self):
        from types import SimpleNamespace
        years = catalog.parse_nature_years(self.fixture('nature_archive.html'))
        years[1]['volumes'] = years[1]['volumes'][1:]
        requests, updates = [], []
        html = self.fixture('nature_archive_volume.html')
        page = SimpleNamespace(goto=lambda url, **kwargs: requests.append(url), content=lambda:html)
        with patch.object(catalog, 'is_cloudflare', return_value=False):
            catalog.read_other_catalog(page, 'nature', 'issues', [1869,1870], years,
                                       threading.Event(), lambda **v:updates.append(v))
        self.assertEqual(len(requests), 1)
        completed = {v['issue_year']:v['issue_rows'] for v in updates if 'issue_year' in v}
        self.assertEqual([r['issue'] for r in completed['1869']], ['9'])
        self.assertEqual([r['issue'] for r in completed['1870']], ['26'])

    def test_actual_links_scoped_to_year_dedup_and_month_only_date(self):
        html = (Path(__file__).parent / 'fixtures/cell_archive.html').read_text(encoding='utf-8')
        years, issues, unsupported = catalog.parse_archive(html, catalog.ARCHIVE_URL + '?page=3')
        self.assertEqual([y['year'] for y in years], [2026, 1990, 1974, 2025])
        self.assertTrue(all(y['page_url'].endswith('?page=3') for y in years))
        self.assertEqual(len(issues['2026']), 1)
        self.assertEqual([row['volume'] for row in issues['1990']], [63, 60])
        self.assertIn('January 1974', issues['1974'][0]['detail'])
        self.assertNotIn('2025', issues)  # Collapsed is not mistaken for empty/loaded.
        self.assertEqual(unsupported, [])

    def test_combined_issues_supplements_and_reject_foreign_links(self):
        urls = [catalog.ARCHIVE_URL.replace('/issues', suffix) for suffix in (
            '/vol/10/issue/1-2', '/vol/10/suppl/C', '/vol/11/suppl/C')]
        self.assertEqual(web.validate_urls(urls, 'cell'), urls)
        self.assertEqual([_sheet_name_from_url(u) for u in urls], ['v10-i1-2', 'v10-suppl-C', 'v11-suppl-C'])
        for url in ('https://evil.test/journal/cell/vol/1/issue/1', 'javascript:alert(1)', '/journal/nature/vol/1/issue/1'):
            self.assertEqual(catalog.canonical_issue_url(url), '')
        html = '<button aria-controls="y">2020 — Volume 180</button><div id="y"><a class="js-issue-item-link" href="/unexpected">Special issue</a></div>'
        self.assertEqual(catalog.parse_archive(html)[2], ['2020: Special issue'])

    def test_worker_error_and_cancel_preserve_completed_years(self):
        session = catalog.CatalogSession()
        def fail(mode, years, known, profile, stop, update, source='cell'):
            update(issue_year='1990', issue_rows=[{'url': 'kept'}])
            raise RuntimeError('network failed')
        with patch.object(catalog, 'read_catalog', side_effect=fail):
            session.start('issues', [1990, 1974], Path('.'))
            worker = session.thread
            if worker: worker.join(5)
        self.assertFalse(session.snapshot()['running'])
        self.assertEqual(session.snapshot()['status'], 'error')
        self.assertIn('1990', session.snapshot()['issues'])
        def cancel(mode, years, known, profile, stop, update, source='cell'):
            stop.wait(5)
            raise catalog.CatalogCancelled()
        with patch.object(catalog, 'read_catalog', side_effect=cancel):
            session.start('years', [], Path('.'))
            worker = session.thread
            session.cancel(); worker.join(5)
        self.assertEqual(session.snapshot()['status'], 'cancelled')
        self.assertIn('1990', session.snapshot()['issues'])

    def test_wait_pumps_events_for_manual_challenge(self):
        from types import SimpleNamespace
        pumps, updates = [], []
        page = SimpleNamespace(wait_for_timeout=lambda ms: pumps.append(ms))
        with patch.object(catalog, 'is_cloudflare', side_effect=[True, True, False]):
            result = catalog._wait(page, lambda:'ready', threading.Event(), lambda **v:updates.append(v))
        self.assertEqual(result, 'ready')
        self.assertEqual(len(pumps), 2)
        self.assertEqual([u['status'] for u in updates], ['waiting', 'reading'])
        stop = threading.Event(); stop.set()
        with self.assertRaises(catalog.CatalogCancelled):
            catalog._wait(page, lambda:False, stop, lambda **v:None)

    def test_hydration_dropped_click_is_retried_only_while_collapsed(self):
        class Button:
            clicks = 0
            def click(self): self.clicks += 1
            def get_attribute(self, name): return 'true' if self.clicks >= 2 else 'false'
        button = Button()
        class Page:
            ticks = 0
            def wait_for_timeout(self, ms): self.ticks += 1
            def content(self):
                anchor = '<a class="js-issue-item-link" href="/journal/cell/vol/180/issue/1">Volume 180, Issue 1</a>' if button.clicks >= 2 and self.ticks >= 5 else ''
                return '<button aria-controls="y">2020 — Volumes 180-183</button><div id="y">' + anchor + '</div>'
        page = Page()
        with patch.object(catalog, 'is_cloudflare', return_value=False), patch.object(catalog.time, 'monotonic', side_effect=lambda: float(page.ticks)):
            rows = catalog._read_year(page, button, 2020, catalog.ARCHIVE_URL, threading.Event(), lambda **v:None)
        self.assertEqual(button.clicks, 2)  # Further waits do not collapse it again.
        self.assertEqual(rows[0]['volume'], 180)


class CatalogWebTests(unittest.TestCase):
    def setUp(self):
        for obj, attr, value in ((web, 'CATALOGS', {key: catalog.CatalogSession(key) for key in web.SOURCES}),
                                 (web, 'SESSIONS', {k: web.Session(k) for k in web.SOURCES}),
                                 (web, 'bus', web.EventBus())):
            patcher = patch.object(obj, attr, value)
            patcher.start(); self.addCleanup(patcher.stop)

    def test_invalid_requests_do_not_launch_browser(self):
        for payload in ({'mode': 'no'}, {'mode': 'issues', 'years': [1900]},
                        {'mode': 'issues', 'years': [True]}, {'years': '2026'}):
            with self.assertRaises(web.HTTPException): web.catalog_start(payload, 'cell')
        with self.assertRaises(web.HTTPException): web.catalog_status('invalid')
        self.assertIsNone(web.CATALOGS['cell'].thread)

    def test_source_catalogs_run_independently_and_cancel_only_one(self):
        entered = {key: threading.Event() for key in ('nature','science')}
        release = threading.Event()
        def read(mode, years, known, profile, stop, update, source):
            entered[source].set()
            update(years=[dict(year=2026, label=source, page_url=catalog.ARCHIVE_URLS[source])])
            release.wait(5)
        with patch.object(catalog, 'read_catalog', side_effect=read):
            try:
                web.catalog_start({}, 'nature'); web.catalog_start({}, 'science')
                self.assertTrue(all(event.wait(2) for event in entered.values()))
                self.assertTrue(web.catalog_status('nature')['running'])
                self.assertTrue(web.catalog_status('science')['running'])
                self.assertFalse(web.catalog_status('cell')['running'])
                web.catalog_cancel('nature')
                self.assertTrue(web.CATALOGS['nature'].stop_event.is_set())
                self.assertFalse(web.CATALOGS['science'].stop_event.is_set())
                self.assertEqual(web.catalog_status('science')['years'][0]['label'], 'science')
            finally:
                workers = [c.thread for c in web.CATALOGS.values() if c.thread]
                release.set()
                for worker in workers: worker.join(5)

    def test_scraper_and_catalog_cannot_share_browser_profile(self):
        urls = dict(cell='https://www.sciencedirect.com/journal/cell/vol/189/issue/1',
                    nature='https://www.nature.com/nature/volumes/657/issues/8132',
                    science='https://www.science.org/toc/science/393/6817')
        for source, url in urls.items():
            with self.subTest(source=source):
                web.SESSIONS[source].thread = object()
                with self.assertRaises(web.HTTPException) as exc: web.catalog_start({}, source)
                self.assertEqual(exc.exception.status_code, 409)
                web.SESSIONS[source].thread = None
                web.CATALOGS[source].thread = object()
                with self.assertRaises(web.HTTPException) as exc:
                    web.start({'urls': [url]}, source)
                self.assertEqual(exc.exception.status_code, 409)
                web.CATALOGS[source].thread = None

    def test_concurrent_catalog_requests_launch_once(self):
        barrier, release = threading.Barrier(2), threading.Event()
        calls = []
        def read(*args, **kwargs): calls.append(1); release.wait(5)
        def start():
            barrier.wait()
            try: web.catalog_start({}, 'cell'); return 200
            except web.HTTPException as exc: return exc.status_code
        with patch.object(catalog, 'read_catalog', side_effect=read):
            try:
                with ThreadPoolExecutor(max_workers=2) as executor:
                    futures = [executor.submit(start) for _ in range(2)]
                    self.assertEqual(sorted(f.result() for f in futures), [200, 409])
                self.assertEqual(len(calls), 1)
            finally:
                worker = web.CATALOGS['cell'].thread
                release.set()
                if worker: worker.join(5)


if __name__ == '__main__':
    unittest.main()
