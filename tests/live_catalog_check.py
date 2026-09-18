"""Read real archive years/issues in a visible browser; never save URL configs.

Usage: python tests/live_catalog_check.py nature --years 1869 2026
Complete any challenge in the browser; the event loop keeps pumping while waiting.
"""
import argparse
import json
import sys
import threading
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from issue_catalog import read_catalog
from unified_web import SOURCE_PROFILES


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('source', choices=SOURCE_PROFILES)
    parser.add_argument('--years', type=int, nargs='+')
    args = parser.parse_args()
    state = dict(issues={})

    def update(**values):
        if 'issue_year' in values:
            year, rows = values.pop('issue_year'), values.pop('issue_rows')
            state['issues'][year] = rows
            print(f'{args.source} {year}: {len(rows)} issues', flush=True)
        state.update(values)
        if values.get('message'):
            print(values['message'], flush=True)

    stop = threading.Event()
    read_catalog('years', [], [], SOURCE_PROFILES[args.source], stop, update, source=args.source)
    years = state['years']
    print(f'{len(years)} years: {years[-1]["year"]} - {years[0]["year"]}', flush=True)
    selected = args.years or [years[0]['year']]
    if not set(selected).issubset({y['year'] for y in years}):
        raise ValueError('Selected years are absent from the archive')
    read_catalog('issues', selected, years, SOURCE_PROFILES[args.source], stop, update, source=args.source)
    out = Path(__file__).resolve().parents[1] / 'data' / 'live_checks' / f'{args.source}_archive' / 'catalog.json'
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'PASS: {out}', flush=True)


if __name__ == '__main__':
    main()
