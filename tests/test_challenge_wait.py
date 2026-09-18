"""Challenge pages must wait for proof of recovery, never exhaust into a row."""
import sys
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
import scraper
import unified_web as web

URL = 'https://www.sciencedirect.com/science/article/pii/S0092867426003946'
CHALLENGE = '<title>ScienceDirect</title>' + ' ' * 35000 + '<meta content="Are you a robot?" name="citation_title">'
ARTICLE = '<meta name="citation_title" content="Recovered paper">'


class Page:
    url = URL
    html = CHALLENGE
    def content(self): return self.html
    def evaluate(self, *args): return False
    def wait_for_timeout(self, *args): pass
    def wait_for_load_state(self, *args, **kwargs): pass


class QuietCallbacks(scraper.ScraperCallbacks):
    def log(self, msg): pass


class ChallengeWaitTests(unittest.TestCase):
    def test_late_metadata_and_visible_title_challenges_are_detected(self):
        self.assertTrue(scraper.detect_cf_in_html(CHALLENGE)[0])
        self.assertTrue(scraper.detect_cf_in_html(' ' * 35000 + '<h1>Are you a <span>robot?</span></h1>')[0])
        self.assertTrue(scraper.detect_cf_in_html('<meta property="og:title" content="Just a moment...">')[0])
        self.assertFalse(scraper.detect_cf_in_html(ARTICLE)[0])

    def test_premature_confirmations_do_not_exhaust_wait(self):
        page = Page()
        class Callback(QuietCallbacks):
            calls = 0
            def cf_wait(self, target, current, attempt, maximum, check_clear):
                self.calls += 1
                if self.calls > 8: raise AssertionError('Unexpected loop')
                self.assert_not_clear = not check_clear()
                if self.calls == 6: page.html = ARTICLE
                return 'user_resumed'
        cb = Callback()
        self.assertTrue(scraper.wait_until_cf_clear(page, URL, max_attempts=1, cb=cb))
        self.assertEqual(cb.calls, 6)
        self.assertTrue(cb.assert_not_clear)

    def test_missed_detector_still_requires_real_extraction_after_manual_wait(self):
        page = Page()
        def extract(page, url):
            return dict(url=url, title='Are you a robot?' if page.html == CHALLENGE else 'Recovered paper')
        class Callback(QuietCallbacks):
            def cf_wait(self, target, current, attempt, maximum, check_clear):
                for _ in range(20):
                    if check_clear(): raise AssertionError('Challenge was incorrectly auto-cleared')
                page.html = ARTICLE
                return 'user_resumed'
        # Reproduce the exact mismatch observed in the user's logs.
        with patch.object(scraper, 'is_cloudflare', return_value=False):
            fields = scraper.extract_with_cf_retry(page, URL, Callback(), 'Preview', extract,
                                                   lambda *args:None, max_retries=1)
        self.assertEqual(fields['title'], 'Recovered paper')

    def test_explicit_skip_and_cancellation_do_not_extract(self):
        for answer in ('skip', 'cancelled'):
            with self.subTest(answer=answer):
                cb = QuietCallbacks()
                cb.cf_wait = lambda *args:answer
                self.assertFalse(scraper.wait_until_cf_clear(Page(), URL, cb=cb))

    def test_wrong_page_never_auto_clears(self):
        page = Page(); page.html = ARTICLE; page.url = 'https://www.sciencedirect.com/'
        class Callback(QuietCallbacks):
            def cf_wait(self, target, current, attempt, maximum, check_clear):
                if check_clear(): raise AssertionError('Wrong page accepted')
                return 'skip'
        self.assertFalse(scraper.wait_until_cf_clear(page, URL, cb=Callback(), force=True))

    def test_web_wait_stays_blocked_and_emits_no_article_until_recovery(self):
        sessions = {key:web.Session(key) for key in web.SOURCES}
        session = sessions['cell']
        page = Page()
        polled, recovered = threading.Event(), threading.Event()
        result = []
        def extract(page, url):
            if not recovered.is_set():
                return dict(title='Are you a robot?', url=url)
            return dict(title='Recovered paper', url=url)
        def detector(page):
            polled.set()
            return False  # Detector false negative must not cause auto-resume.
        with patch.object(web, 'SESSIONS', sessions), patch.object(web, 'bus', web.EventBus()), patch.object(scraper, 'is_cloudflare', side_effect=detector):
            cb = web.WebCallbacks('cell')
            worker = threading.Thread(target=lambda: result.append(scraper.extract_with_cf_retry(
                page, URL, cb, 'Preview', extract, lambda *args:None, max_retries=1)))
            worker.start()
            try:
                self.assertTrue(polled.wait(2))
                worker.join(0.8)
                self.assertTrue(worker.is_alive())
                self.assertEqual(session.status, 'cf_blocked')
                self.assertEqual(session.articles, {})
                self.assertEqual(result, [])
                recovered.set()
                worker.join(3)
                self.assertFalse(worker.is_alive())
                self.assertEqual(result[0]['title'], 'Recovered paper')
                self.assertEqual(session.status, 'running')
            finally:
                session.stop_event.set()
                worker.join(3)


if __name__ == '__main__':
    unittest.main()
