"""Publication dates and author/affiliation evidence, independent of browsers."""
from __future__ import annotations

import json
import re
from datetime import date
from html import unescape
from html.parser import HTMLParser


class Node:
    def __init__(self, tag='', attrs=(), parent=None):
        self.tag, self.attrs, self.parent = tag, dict(attrs), parent
        self.children = []

    def walk(self):
        yield self
        for child in self.children:
            if isinstance(child, Node):
                yield from child.walk()

    def text(self):
        if self.tag in ('script', 'style'):
            return ''
        return ' '.join(c.text() if isinstance(c, Node) else c for c in self.children)

    def has_class(self, name):
        return name in self.attrs.get('class', '').split()


class Document(HTMLParser):
    def __init__(self, html):
        super().__init__(convert_charrefs=True)
        self.root = Node()
        self.current = self.root
        self.feed(html)

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs, self.current)
        self.current.children.append(node)
        if tag not in ('meta', 'link', 'img', 'br', 'hr', 'input', 'source', 'wbr', 'area', 'base', 'embed', 'param', 'track', 'col'):
            self.current = node

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)

    def handle_endtag(self, tag):
        node = self.current
        while node.parent:
            if node.tag == tag:
                self.current = node.parent
                return
            node = node.parent

    def handle_data(self, data):
        self.current.children.append(data)


def clean(value):
    return re.sub(r'\s+', ' ', unescape(value or '')).strip()


MONTHS = {m.lower(): i for i, m in enumerate(
    ('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'), 1)}
MONTHS.update({m[:3]: i for m, i in list(MONTHS.items())})
MONTHS['sept'] = 9
DATE_PATTERN = r'(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}\s+[A-Za-z]+\.?\s+\d{4}|[A-Za-z]+\.?\s+\d{1,2},?\s+\d{4})'


def normalize_date(raw):
    raw = clean(raw)
    match = re.search(DATE_PATTERN, raw)
    if not match:
        return ''
    value = match.group()
    try:
        if value[:4].isdigit() and value[4] in '-/':
            year, month, day = map(int, re.split('[-/]', value))
        else:
            parts = value.replace(',', '').replace('.', '').split()
            if parts[0].isdigit():
                day, month, year = int(parts[0]), MONTHS[parts[1].lower()], int(parts[2])
            else:
                month, day, year = MONTHS[parts[0].lower()], int(parts[1]), int(parts[2])
        return date(year, month, day).isoformat()
    except (ValueError, KeyError):
        return ''


def meta_values(root):
    result = {}
    for n in root.walk():
        if n.tag == 'meta':
            key = (n.attrs.get('name') or n.attrs.get('property') or n.attrs.get('itemprop') or '').lower()
            result.setdefault(key, []).append(clean(n.attrs.get('content')))
    return result


def publication_dates(html, source):
    root = Document(html).root
    # Do not let publication labels in references or related-article cards win.
    for node in list(root.walk()):
        marker = ' '.join(str(node.attrs.get(k, '')) for k in ('id', 'class', 'data-title'))
        if node.parent and (node.tag in ('footer', 'nav') or re.search(r'\b(?:references|bibliography|card-related)\b', marker, re.I)):
            node.parent.children.remove(node)
    text = clean(root.text())
    result = dict(available_online='', version_of_record='', published_date='', date_evidence={})

    def record(key, raw, evidence):
        value = normalize_date(raw)
        if value and not result[key]:
            result[key] = value
            result['date_evidence'][key] = clean(evidence)[:200]

    if source == 'cell':
        # Explicit labels only. Never substitute issue/cover/accepted dates.
        for key, label in [('version_of_record', r'Version\s+of\s+Record'),
                           ('available_online', r'Available\s+online')]:
            pattern = rf'\b{label}\s*[:：,;–—-]?\s*({DATE_PATTERN})'
            for m in re.finditer(pattern, text, re.I):
                # "Version of Record available online" belongs to the VoR field.
                if key == 'available_online' and re.search(r'Version of Record\s*$', text[max(0, m.start()-25):m.start()], re.I):
                    continue
                record(key, m.group(1), m.group())
        for m in re.finditer(rf'Version\s+of\s+Record\s+available\s+online\s*:?\s*({DATE_PATTERN})', text, re.I):
            record('version_of_record', m.group(1), m.group())
        # Current ScienceDirect renders its article information from this dates
        # object. Read its exact labels even before the information panel opens.
        for match in re.finditer(r'"dates"\s*:\s*', html):
            try:
                dates, _ = json.JSONDecoder().raw_decode(html[match.end():])
            except ValueError:
                continue
            if not isinstance(dates, dict):
                continue
            for key, label in [('available_online', 'Available online'),
                               ('version_of_record', 'Version of Record')]:
                raw = dates.get(label)
                if isinstance(raw, str):
                    record(key, raw, f'dates.{label}: {raw}')
        # Explicit Elsevier date keys, including dates embedded before expansion.
        for key, names in [('available_online', ('available-online-date', 'availableOnlineDate')),
                           ('version_of_record', ('vor-available-online-date', 'versionOfRecordDate'))]:
            for name in names:
                for m in re.finditer(r'"' + re.escape(name) + r'"\s*:\s*"([^"]+)"', html):
                    record(key, m.group(1), name + ': ' + m.group(1))
    else:
        metas = meta_values(root)
        # Online/publication metadata describes this article, not its issue.
        for name in ('citation_online_date', 'dc.date', 'dc.date.issued', 'citation_publication_date', 'prism.publicationdate', 'datepublished'):
            for raw in metas.get(name, []):
                record('published_date', raw, f'{name}: {raw}')
        for n in root.walk():
            if n.tag == 'time' and ('datePublished' in n.attrs.get('itemprop', '') or n.parent and re.search(r'publish|pub-date', n.parent.attrs.get('class', ''), re.I)):
                raw = n.attrs.get('datetime') or n.text()
                record('published_date', raw, f'time: {raw}')
        if not result['published_date']:
            for m in re.finditer(rf'\bPublished(?:\s+online|\s+in\s+print)?\s*:?\s*({DATE_PATTERN})', text, re.I):
                record('published_date', m.group(1), m.group())
    return result


def _name_key(name):
    return re.sub(r'[^\w]', '', clean(name).casefold())


def _same_name(left, right):
    return _name_key(left) == _name_key(right)


def _json_nodes(value):
    if isinstance(value, dict):
        yield value
        for item in value.values():
            yield from _json_nodes(item)
    elif isinstance(value, list):
        for item in value:
            yield from _json_nodes(item)


def _json_text(value):
    if isinstance(value, dict):
        return clean(' '.join([str(value.get('_', ''))] + [_json_text(v) for k, v in value.items() if k == '$$']))
    if isinstance(value, list):
        return clean(' '.join(_json_text(v) for v in value))
    return ''


def cell_author_groups(html):
    decoder = json.JSONDecoder()
    # Decode complete author-group trees instead of fixed 800-character windows.
    for match in re.finditer(r'\{\s*"#name"\s*:\s*"author-group"', html):
        try:
            group, _ = decoder.raw_decode(html[match.start():])
        except ValueError:
            continue
        yield group


def cell_authors(html):
    for group in cell_author_groups(html):
        names = []
        seen = set()
        for obj in _json_nodes(group):
            if obj.get('#name') != 'author':
                continue
            aid = obj.get('$', {}).get('author-id')
            name = clean(' '.join(_json_text(n) for n in _json_nodes(obj) if n.get('#name') in ('given-name', 'surname')))
            if name and (aid or name) not in seen:
                seen.add(aid or name)
                names.append(name)
        if names:
            return names
    return []


def cell_author_affiliations(html, first_author):
    for group in cell_author_groups(html):
        objects = list(_json_nodes(group))
        affs = {}
        for obj in objects:
            if obj.get('#name') == 'affiliation':
                aid = obj.get('$', {}).get('id', '')
                texts = [_json_text(n) for n in _json_nodes(obj) if n.get('#name') == 'textfn']
                affs[aid] = clean(' '.join(texts) or _json_text(obj))
        for obj in objects:
            if obj.get('#name') != 'author':
                continue
            children = list(_json_nodes(obj))
            name = clean(' '.join(_json_text(n) for n in children if n.get('#name') in ('given-name', 'surname')))
            if not _same_name(name, first_author):
                continue
            refs = []
            for n in children:
                if n.get('#name') in ('cross-ref', 'xref'):
                    refs.extend(re.split(r'[\s,]+', n.get('$', {}).get('refid', '')))
            refs = list(dict.fromkeys(r for r in refs if r.startswith('aff')))
            if refs:
                return [affs[r] for r in refs if affs.get(r)], all(affs.get(r) for r in refs)
    return [], False


def author_affiliations(html, source, first_author):
    """Return explicitly linked affiliations and whether the mapping is complete."""
    if not first_author:
        return [], False
    if source == 'cell':
        return cell_author_affiliations(html, first_author)
    root = Document(html).root
    nodes = list(root.walk())
    if source == 'nature':
        affs = {}
        matches = []
        for n in nodes:
            aid = n.attrs.get('id', '')
            if not re.fullmatch(r'Aff\d+', aid, re.I):
                continue
            address = next((clean(c.text()) for c in n.walk() if c.has_class('c-article-author-affiliation__address')), '')
            affs[aid.lower()] = address
            for c in n.walk():
                if c.has_class('c-article-author-affiliation__authors-list'):
                    names = re.split(r'\s*(?:,|&|\band\b)\s*', clean(c.text()))
                    if any(_same_name(name, first_author) for name in names) and address:
                        matches.append(address)
        # Author links are preferable and let us detect unresolved affiliation IDs.
        for n in nodes:
            if n.tag == 'li' and n.has_class('c-article-author-list__item'):
                name = next((clean(c.attrs.get('content') or c.text()) for c in n.walk() if c.attrs.get('itemprop') == 'name' or c.attrs.get('data-test') == 'author-name'), '')
                if _same_name(name, first_author):
                    refs = [c.attrs.get('href', '')[1:].lower() for c in n.walk() if re.fullmatch(r'#Aff\d+', c.attrs.get('href', ''), re.I)]
                    if refs:
                        return list(dict.fromkeys(affs[r] for r in refs if affs.get(r))), all(affs.get(r) for r in refs)
        return list(dict.fromkeys(matches)), bool(matches)
    # Science: stay inside the first contributor, never read into con2.
    content = next((n for n in nodes if n.attrs.get('id') == 'con1_content'), None)
    if content:
        affs = []
        for n in content.walk():
            if n.attrs.get('property') == 'affiliation':
                value = next((clean(c.text()).rstrip('.') for c in n.walk() if c.attrs.get('property') == 'name'), '')
                if value:
                    affs.append(value)
        return list(dict.fromkeys(affs)), bool(affs)
    return [], False


def enrich_fields(fields, html, source, parse_country):
    metas = meta_values(Document(html).root)
    # Attribute order and quoting are handled by HTMLParser, not regex.
    for key, names in [('title', ('citation_title', 'dc.title')),
                       ('doi', ('citation_doi', 'prism.doi', 'publication_doi', 'doi'))]:
        for name in names:
            if not fields.get(key) and metas.get(name):
                fields[key] = metas[name][0]
    if source == 'cell':
        structured_authors = cell_authors(html)
        if structured_authors:
            fields['authors'] = structured_authors
            fields['first_author'] = structured_authors[0]
    if not fields.get('authors'):
        authors = metas.get('dc.creator') or metas.get('citation_author') or []
        if source == 'nature':
            authors = [' '.join(part.strip() for part in reversed(a.split(',', 1))) if a.count(',') == 1 else a for a in authors]
        fields['authors'] = authors
    if not fields.get('first_author') and fields.get('authors'):
        fields['first_author'] = fields['authors'][0]
    fields.update(publication_dates(html, source))
    affs, complete = author_affiliations(html, source, fields.get('first_author', ''))
    fields['first_author_affiliations'] = affs
    fields['affiliation_verified'] = complete
    if affs:
        fields['first_aff'] = '; '.join(affs)
    countries = list(dict.fromkeys(parse_country(aff) for aff in affs))
    fields['first_author_countries'] = countries
    fields['first_author_country'] = '; '.join(countries)
    fields['is_china'] = (True if 'China' in countries else
                          False if complete and countries and all(c and not c.startswith('?') for c in countries) else None)
    fields['source'] = source
    fields['statistics_policy'] = 'first-listed-author; any verified affiliation in China'
    return finalize_fields(fields, source)


FAILED_TITLES = ('[CF BLOCKED]', '[GOTO FAILED]')


def finalize_fields(fields, source):
    required = ['title', 'doi', 'first_author']
    required += ['available_online', 'version_of_record'] if source == 'cell' else ['published_date']
    missing = [key for key in required if not fields.get(key)]
    if not fields.get('affiliation_verified'):
        missing.append('author_affiliation_mapping')
    if fields.get('is_china') is None:
        missing.append('country')
    if fields.get('title') in FAILED_TITLES or not fields.get('title'):
        fields['extraction_status'] = 'failed'
        fields['is_china'] = None
    else:
        fields['extraction_status'] = 'partial' if missing else 'complete'
    fields['missing_fields'] = missing
    fields['source'] = source
    return fields


def china_label(value):
    return '是' if value is True else '否' if value is False else '待核实'


def excluded_title(title):
    return bool(re.match(r'^(Author|Publisher)\s+Correction\s*:', clean(title), re.I))


ID_PATTERNS = {
    'cell': r'/science/article/pii/(S\d{16})(?=[/?#]|$)',
    'nature': r'/articles/(s41586-\d{3}-\d{4,7}-[a-z0-9]+)(?=[/?#]|$)',
    'science': r'/doi/(?:abs/|full/|pdf/|epdf/)?(10\.1126/science\.[a-z0-9]+)(?=[/?#]|$)',
}


def article_id(url, source):
    m = re.search(ID_PATTERNS[source], url)
    return m.group(1) if m else ''


def expected_article_ids(html, source):
    """Independent HTMLParser traversal; same scope policy, separate DOM path."""
    root = Document(html).root
    found, excluded = set(), set()
    for node in root.walk():
        if node.tag != 'a':
            continue
        aid = article_id(node.attrs.get('href', ''), source)
        if not aid:
            continue
        parent = node
        outside = False
        while parent:
            if parent.tag in ('header', 'footer', 'nav', 'aside') or parent.has_class('card-related'):
                outside = True
                break
            parent = parent.parent
        if outside:
            continue
        if source in ('cell', 'nature') and excluded_title(node.text()):
            excluded.add(aid)
        found.add(aid)
    return found - excluded


def count_check(html, targets, source, issue_url):
    expected = expected_article_ids(html, source)
    actual = {article_id(url, source) for _, url, _ in targets} - {''}
    return dict(phase='count_check', issue_url=issue_url, expected=len(expected), actual=len(actual),
                missing_ids=sorted(expected-actual), extra_ids=sorted(actual-expected),
                matched=expected == actual and bool(expected))
