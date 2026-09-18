"""Local UI smoke-test server with synthetic data and isolated temporary files.

Run: uv run python tests/preview_server.py
No publisher websites are visited and production URL/cache files are not used.
"""
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
import uvicorn
import unified_web as web
import issue_catalog
from article_metadata import enrich_fields
from countries import parse_country
from excel_writer import write_excel
from fastapi.responses import HTMLResponse


def main():
    with tempfile.TemporaryDirectory(prefix='cns-ui-check-') as directory:
        urls = {
            'cell': 'https://www.sciencedirect.com/journal/cell/vol/189/issue/10',
            'nature': 'https://www.nature.com/nature/volumes/654/issues/8119',
            'science': 'https://www.science.org/toc/science/392/6804',
        }
        fixtures = Path(__file__).parent / 'fixtures'

        def fake_catalog(mode, selected, known, profile, stop, update):
            # Exercise the real background API, without opening a publisher browser.
            update(status='waiting', message='离线模拟：等待目录加载…')
            if stop.wait(1):
                raise issue_catalog.CatalogCancelled()
            years, issues, _ = issue_catalog.parse_archive((fixtures / 'cell_archive.html').read_text(encoding='utf-8'))
            if mode == 'years':
                update(years=years, years_complete=True, message='离线测试：年份目录已读取')
            else:
                for year in selected:
                    if str(year) not in issues:
                        raise RuntimeError('离线测试：该年份模拟加载失败')
                    update(issue_year=str(year), issue_rows=issues[str(year)])
                update(message='离线测试：期号已读取')

        issue_catalog.read_catalog = fake_catalog

        def make_run(source):
            def run(urls, out_path, cb, **kwargs):
                cb.on_state(dict(phase='issue_progress', issue_idx=1, issue_total=1, issue_url=urls[0]))
                cb.on_state(dict(phase='count_check', issue_url=urls[0], actual=1, expected=1, matched=True))
                # Long enough to exercise Stop, without contacting any external service.
                for _ in range(20):
                    if cb.is_cancelled():
                        cb.on_state(dict(phase='cancelled', out_path=''))
                        return
                    time.sleep(0.1)
                html = (fixtures / f'{source}_article.html').read_text(encoding='utf-8')
                if source == 'cell':
                    html += (fixtures / 'cell_history.html').read_text(encoding='utf-8')
                fields = dict(title='离线测试样例（不是实际论文）', doi='10.test/fixture', first_author='Alice Li', authors=['Alice Li'], first_aff='', url='https://example.org/fixture', issue_url=urls[0], section='Articles')
                fields = enrich_fields(fields, html, source, parse_country)
                cb.on_state(dict(phase='article_done', url=fields['url'], fields=fields))
                actual = write_excel(out_path, [(urls[0], [('Articles', fields['url'], fields)])], columns=web.SOURCES[source].columns)
                cb.on_state(dict(phase='all_done', out_path=actual))
            return run

        for source, cfg in web.SOURCES.items():
            cfg.urls_file = Path(directory) / f'{source}.txt'
            cfg.urls_file.write_text(urls[source], encoding='utf-8')
            cfg.cache_dir = Path(directory) / f'cache_{source}'
            cfg.default_out = lambda source=source: str(Path(directory) / f'{source}.xlsx')
            cfg.run_scraper = make_run(source)
        for route in web.app.routes:
            if getattr(route, 'path', '') == '/':
                web.app.routes.remove(route)
                break

        @web.app.get('/')
        def index():
            html = (web.STATIC_DIR / 'unified_index.html').read_text(encoding='utf-8')
            return HTMLResponse(html.replace('CNS 期刊爬虫 · 统一控制台', '离线测试 · 模拟数据'))

        uvicorn.run(web.app, host='127.0.0.1', port=8765, log_level='warning')


if __name__ == '__main__':
    main()
