"""Lossless presentation text blocks reconstructed from the original data rules.

Only explicit leading service metadata is removed. No model, network, archive
bytecode or generated factual claims are involved. Text is never length-cut;
the presentation renderer owns wrapping and continuation pages.
"""
from __future__ import annotations

import html
import re
from html.parser import HTMLParser


_BLOCK_TAGS = frozenset({'p', 'div', 'section', 'article', 'header', 'footer', 'h1',
    'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'ul', 'ol', 'blockquote', 'pre', 'tr'})
_INLINE_TAGS = frozenset({'a', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'span',
    'font', 'small', 'big', 'sup', 'sub', 'code', 'table', 'tbody', 'thead', 'tfoot'})
_HIDDEN_TAGS = frozenset({'script', 'style'})
_PLACEHOLDERS = frozenset({'', '-', '—', 'None'})
_REGION = re.compile(r'^(?:Россия,[ \t]*)?Московская область,[ \t]*', re.I)
_ADDRESS = re.compile(r'^Адрес[ \t]*:[ \t]*([^;]+)(?:;[ \t]*(.*))?$', re.I)
# These are whole leading metadata fields, not arbitrary matches in prose.
_SERVICE_PREFIX = re.compile(
    r'^(?:Номер Добродела:[ \t]*\d+|ППМО:[ \t]*\d+(?:-\d+)*|'
    r'Авторизация прошла через ЕСИА)(?:[ \t]*;[ \t]*|[ \t]*$)', re.I)


class _TextReader(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.hidden = []

    def handle_starttag(self, tag, attrs):
        if self.hidden:
            if tag in _HIDDEN_TAGS:
                self.hidden.append(tag)
            return
        if tag in _HIDDEN_TAGS:
            self.hidden.append(tag)
        elif tag in _BLOCK_TAGS or tag in {'br', 'hr'}:
            self.parts.append('\n')
        elif tag in {'td', 'th'}:
            self.parts.append(' ')
        elif tag not in _INLINE_TAGS:
            # Unknown angle-bracket text may be meaningful notation.
            self.parts.append(self.get_starttag_text())

    def handle_endtag(self, tag):
        if self.hidden:
            if tag == self.hidden[-1]:
                self.hidden.pop()
            return
        if tag in _BLOCK_TAGS:
            self.parts.append('\n')
        elif tag in {'td', 'th'}:
            self.parts.append(' ')
        elif tag not in _INLINE_TAGS and tag not in {'br', 'hr'}:
            self.parts.append('</' + tag + '>')

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag in _HIDDEN_TAGS:
            self.handle_endtag(tag)

    def handle_data(self, data):
        if not self.hidden:
            self.parts.append(data)

    def handle_comment(self, data):
        pass


def plain_text(value) -> str:
    """Strip supported HTML markup, preserving line breaks and comparisons."""
    if value is None:
        return ''
    text = str(value).replace('\r\n', '\n').replace('\r', '\n')
    for _ in range(2):
        text = html.unescape(text)
    reader = _TextReader()
    reader.feed(text)
    reader.close()
    # Keep every nonempty source paragraph, without multiplying blank lines.
    lines = [re.sub(r'[^\S\n]+', ' ', line).strip()
             for line in ''.join(reader.parts).replace('\xa0', ' ').split('\n')]
    result = '\n'.join(line for line in lines if line)
    return '' if result in _PLACEHOLDERS else result


def clean(value) -> str:
    """Single-line form, for organisation names and compact labels only."""
    return re.sub(r'\s+', ' ', plain_text(value)).strip()


def short_org(value) -> str:
    """Known display abbreviations; never cut unfamiliar organisation names."""
    org = clean(value)
    if re.match(r'^Министерство жилищно-коммунального хозяйства\b', org, re.I):
        return 'Министерство ЖКХ'
    if re.match(r'^Администрация\b', org, re.I):
        return 'ОМСУ'
    return org


def source_line(row) -> str:
    source = clean(row.get('Источник обращения ЕЦУР')) or clean(row.get('Источник'))
    organisation = short_org(row.get('Ведомство исполнителя'))
    return ' / '.join(part for part in (source, organisation) if part)


def _leading_metadata(text: str) -> str:
    """Remove only whole service fields before the first substantive text."""
    lines = text.splitlines()
    while lines:
        line = lines[0]
        matched = False
        while True:
            match = _SERVICE_PREFIX.match(line)
            if not match:
                break
            matched = True
            line = line[match.end():].lstrip()
        if line:
            lines[0] = line
            break
        if not matched:
            break
        lines.pop(0)
    return '\n'.join(lines)


def split_annotation(value) -> tuple[str, str]:
    """Separate an explicit leading address from the full complaint body."""
    text = _leading_metadata(plain_text(value))
    if not text:
        return '', ''
    lines = text.splitlines()
    address = ''
    match = _ADDRESS.match(lines[0])
    if match:
        address = match.group(1).strip()
        remainder = match.group(2) or ''
        lines = ([remainder] if remainder else []) + lines[1:]
    elif len(lines) > 1 and _REGION.match(lines[0]):
        # A separate source line beginning with the region is the original
        # dashboard's address form. Keep the full remaining line without cuts.
        address = _REGION.sub('', lines[0]).strip()
        lines = lines[1:]
    body = _leading_metadata('\n'.join(lines))
    # Original source compared only the first 40 characters. Restrict dedup
    # to two complete, exactly equal entries so distinct promises survive.
    parts = [part.strip() for part in body.split(';') if part.strip()]
    if len(parts) == 2 and parts[0] == parts[1]:
        body = parts[0]
    return address, body


def content_blocks(row) -> list[tuple[str, bool]]:
    annotation = row.get('Аннотация/краткое содержание')
    if not plain_text(annotation):
        annotation = row.get('Факт')
    address, body = split_annotation(annotation)
    if not address:
        # Some exports place a labelled address in the separate fact field.
        address, _ = split_annotation(row.get('Факт'))
    blocks = []
    if address:
        blocks.append(('Адрес: ' + address, True))
    source = source_line(row)
    if source:
        blocks.append((source, False))
    if body:
        blocks.append((body, False))
    elif not address:
        blocks.append(('Содержание не указано', False))
    return blocks


def status_blocks(row) -> list[tuple[str, bool]]:
    status = plain_text(row.get('Статус')) or 'Статус не указан'
    answer = plain_text(row.get('Ответ'))
    result = [(status, True)]
    if answer:
        result.append((answer, False))
    return result
