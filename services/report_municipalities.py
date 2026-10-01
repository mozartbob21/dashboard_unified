"""Canonical municipality identity shared by reporting adapters."""
import json
import re
from functools import lru_cache
from pathlib import Path

REGISTRY = Path(__file__).resolve().parents[1] / 'data/municipality_registry.json'


def _plain(value):
    value = re.sub(r'\s+', ' ', str(value or '').casefold().replace('ё', 'е')).strip()
    value = re.sub(r'^(?:(?:городской округ|муниципальный округ)\s+|[гм]\s*\.?\s*о(?:\.\s*|\s+))', '', value)
    value = re.sub(r'\s*\(\s*зато\s*\)\s*$', '', value)
    value = re.sub(r'\s+[гм]\s*\.?\s*о\.?\s*$', '', value)
    return value.strip(' .')


def _title(value):
    raw = re.sub(r'^(?:(?:городской округ|муниципальный округ)\s+|[гм]\s*\.?\s*о(?:\.\s*|\s+))', '', str(value or '').strip(), flags=re.I)
    raw = re.sub(r'\s*\(\s*зато\s*\)\s*$', '', raw, flags=re.I).strip()
    raw = re.sub(r'\s+[гм]\s*\.?\s*о\.?\s*$', '', raw, flags=re.I).strip()
    return raw.title() if raw.isupper() else raw


@lru_cache(maxsize=2)
def _registry(path, modified):
    try:
        if Path(path).stat().st_size > 2_000_000:
            return {}, {}
        data = json.loads(Path(path).read_text(encoding='utf-8')).get('municipalities', {})
        aliases, titles = {}, {}
        for name, item in data.items():
            canonical = _plain(name)
            titles[canonical] = _title(name)
            aliases[canonical] = canonical
            for alias in item.get('aliases', []) if isinstance(item, dict) else []:
                if isinstance(alias, str):
                    aliases[_plain(alias)] = canonical
        return aliases, titles
    except (OSError, ValueError, TypeError, AttributeError):
        return {}, {}


def _maps():
    try:
        return _registry(str(REGISTRY), REGISTRY.stat().st_mtime_ns)
    except OSError:
        return {}, {}


def key(value):
    raw = _plain(value)
    return _maps()[0].get(raw, raw)


def display(value):
    canonical = key(value)
    title = _maps()[1].get(canonical)
    if title:
        return title
    return _title(value)


def is_municipality(value):
    """Exclude missing/aggregate/combined labels from the city catalogue."""
    raw = key(value)
    return bool(raw and re.search(r'[а-яa-z]', raw) and '/' not in raw
                and raw not in {'null', 'none', 'nan', 'итого', 'всего', 'все',
                                'москва', 'московская область', 'область', 'нет данных',
                                'не определен', 'не определено', 'не указан', 'не указано',
                                '(пусто)', 'прочие', 'неизвестно'})
