"""Source-specific, explainable drill-downs from verified local snapshots.

Never infer percentages from status labels or substitute missing counts with
zero. Identical repeated tables are harmless; conflicting municipal records
are excluded, rather than silently selecting a value.
"""
from collections import defaultdict

from services.water_dashboard.metrics import _count, _number, _normal, _percent

MUNICIPALITY_HEADERS = ('омсу', 'муниципалитет', 'муниципальный округ', 'городской округ')
TASK_HEADERS = ('кол-во задач', 'количество задач', 'просроченные задачи',
                'количество просроченных задач', 'кол-во просроченных задач')
SYSTEM_HEADERS = ('системных', 'системные адреса', 'количество системных адресов', 'кол-во системных адресов')
RESONANT_HEADERS = ('резонансных', 'резонансные адреса', 'количество резонансных адресов', 'кол-во резонансных адресов')
ATTENDANCE_HEADERS = ('присутствовали на перекличках, %', 'явка, %', 'явка (%)')

EDO_CURRENT_HEADER = 'Доля документов, обработанных в электронном виде, % от общего объема (текущая неделя)'
EDO_PREVIOUS_HEADER = 'Доля документов, обработанных в электронном виде, % от общего объема (прошлая неделя)'


def _name(value):
    text = ' '.join(str(value or '').split())
    if not text or _normal(text).startswith(('итого', 'всего', 'общий итог')):
        return ''
    if text.isupper():
        return text.title()
    return text


def _rows(data, columns):
    """Yield only tables with every requested field, by exact header aliases."""
    for table in data.get('tables') or []:
        headers = [_normal(header) for header in table.get('headers') or []]
        indices = {}
        for key, aliases in dict(name=MUNICIPALITY_HEADERS, **columns).items():
            indices[key] = next((headers.index(_normal(alias)) for alias in aliases if _normal(alias) in headers), None)
        if any(value is None for value in indices.values()):
            continue
        for row in table.get('rows') or []:
            record = {key: row[index] if index < len(row) else None for key, index in indices.items()}
            record['name'] = _name(record['name'])
            if record['name']:
                yield record


def _unique(data, columns, convert):
    """Keep one unambiguous record per municipality, preserving missingness."""
    collected = defaultdict(list)
    names = {}
    for record in _rows(data, columns):
        key = _normal(record['name'])
        names.setdefault(key, record['name'])
        item = convert(record)
        if item is not None:
            item['name'] = names[key]
        if item not in collected[key]:
            collected[key].append(item)
    accepted = [items[0] for items in collected.values() if len(items) == 1 and items[0] is not None]
    return accepted, len(collected) - len(accepted)


def _secondary(label, value, unit=''):
    return {'label': label, 'value': value, 'unit': unit}


def _result(basis, items, excluded, *, higher_is_better=False, best=False, note='', kind='rankings', metrics=None):
    # The private score is used for categorical ordering only and never exposed
    # as a made-up percentage (in particular for ZULUGIS statuses).
    score = lambda item: item.get('_score', item['value'])
    ascending = sorted(items, key=lambda item: (score(item), _normal(item['name']), _normal(item.get('municipality'))))
    descending = sorted(items, key=lambda item: (-score(item), _normal(item['name']), _normal(item.get('municipality'))))
    clean = lambda item: {key: value for key, value in item.items() if not key.startswith('_')}
    groups = []
    if kind == 'rankings':
        worst = ascending if higher_is_better else descending
        groups.append({'id': 'worst', 'title': 'Топ-5 отстающих', 'items': [clean(item) for item in worst[:5]]})
        if best:
            leaders = descending if higher_is_better else ascending
            groups.append({'id': 'best', 'title': 'Топ-5 лучших', 'items': [clean(item) for item in leaders[:5]]})
    return {
        'schema_version': 1, 'kind': kind, 'basis': basis,
        'coverage': {'entities': len(items), 'excluded': excluded, 'note': note},
        'groups': groups,
        'entities': [clean(item) for item in sorted(items, key=lambda item: _normal(item['name']))],
        'metrics': metrics or [],
    }


def is_edo_ranking_table(table):
    headers = {_normal(header) for header in table.get('headers') or []}
    return 'рсо' in headers and bool(headers & set(MUNICIPALITY_HEADERS)) and _normal(EDO_CURRENT_HEADER) in headers


def _edo_rankings(data):
    records = defaultdict(list)
    for table in data.get('ranking_tables') or []:
        if not is_edo_ranking_table(table):
            continue
        headers = [_normal(header) for header in table.get('headers') or []]
        rso_index = headers.index('рсо')
        city_index = next(headers.index(key) for key in MUNICIPALITY_HEADERS if key in headers)
        current_index = headers.index(_normal(EDO_CURRENT_HEADER))
        previous_index = headers.index(_normal(EDO_PREVIOUS_HEADER)) if _normal(EDO_PREVIOUS_HEADER) in headers else None
        for row in table.get('rows') or []:
            cell = lambda index: row[index] if index is not None and index < len(row) else None
            name = ' '.join(str(cell(rso_index) or '').split())
            municipality = _name(cell(city_index))
            if not name or _normal(name).startswith(('итого', 'всего')) or not municipality:
                continue
            value = _percent(cell(current_index), bounded=True)
            previous = _percent(cell(previous_index), bounded=True)
            record = None if value is None else {
                'name': name, 'municipality': municipality, 'value': value, 'unit': '%',
                'secondary': [_secondary('ОМСУ', municipality), _secondary('Прошлая неделя', previous, '%')],
            }
            key = (_normal(name), _normal(municipality))
            if record not in records[key]:
                records[key].append(record)
    items = [values[0] for values in records.values() if len(values) == 1 and values[0] is not None]
    excluded = len(records) - len(items)
    reason = data.get('ranking_error') or 'Полная таблица РСО ещё не получена. Обновите этот источник.'
    result = _result(
        'Доля документов, обработанных в электронном виде за текущую неделю, по РСО. '
        'Выше доля — лучше. При равенстве — по алфавиту. Показатель ЭЦП в общей карточке оценивается отдельно.',
        items, excluded, higher_is_better=True, best=True,
        note=(f'Рейтинг среди {len(items)} РСО с распознаваемыми данными в полной таблице источника.' if items else reason),
    )
    result['entity_type'] = 'rso'
    result['source_url'] = data.get('ranking_url', '')
    return result


def source_details(sid, data):
    """No network or writes: details belong to the source's existing data date."""
    data = data or {}
    if sid == 'valves':
        statuses = {'не внесено': (0, 'Не внесено'), 'частично': (1, 'Частично'), 'полностью': (2, 'Полностью')}

        def convert(row):
            status = statuses.get(_normal(row['status']))
            if status is None:
                return None
            return {'name': row['name'], 'value': status[1], 'unit': '', '_score': status[0],
                    'secondary': [_secondary('План замены', _count(row['plan']), 'шт.')]}

        items, excluded = _unique(data, {
            'status': ('статус занесения в zulugis',),
            'plan': ('кол-во задвижек планируемых к замене, шт',),
        }, convert)
        return _result('Статус внесения в ZULUGIS: «Не внесено» → «Частично» → «Полностью». '
                       'При одинаковом статусе — по алфавиту.', items, excluded,
                       higher_is_better=True, best=True,
                       note='В источнике нет процента выполнения по ОМСУ. План замены приведён для справки и не влияет на место.')
    if sid == 'tasks':
        items, excluded = _unique(data, {'count': TASK_HEADERS}, lambda row:
            {'name': row['name'], 'value': _count(row['count']), 'unit': 'задач', 'secondary': []}
            if _count(row['count']) is not None else None)
        return _result('Больше просроченных задач — ниже результат. При равенстве — по алфавиту.', items, excluded)
    if sid in ('sys_vs', 'sys_kr'):
        def convert(row):
            systems, resonant = _count(row['systems']), _count(row['resonant'])
            if systems is None:
                return None
            return {'name': row['name'], 'value': systems, 'unit': 'адр.',
                    'secondary': [_secondary('Резонансные адреса', resonant, 'адр.')]}
        items, excluded = _unique(data, {'systems': SYSTEM_HEADERS, 'resonant': RESONANT_HEADERS}, convert)
        return _result('Больше системных адресов — ниже результат. Резонансные адреса показаны отдельно: '
                       'категории не суммируются. При равенстве — по алфавиту.', items, excluded)
    if sid == 'edo_rso':
        return _edo_rankings(data)
    if sid == 'meetings':
        items, excluded = _unique(data, {'attendance': ATTENDANCE_HEADERS}, lambda row:
            {'name': row['name'], 'value': _percent(row['attendance'], bounded=True), 'unit': '%', 'secondary': []}
            if _percent(row['attendance'], bounded=True) is not None else None)
        return _result('Доля присутствия на перекличках: ниже явка — ниже результат. '
                       'При равенстве — по алфавиту.', items, excluded, higher_is_better=True)
    if sid == 'nvos':
        # Keep all typed measurements for compact source detail and AI reports.
        # Municipal data remain in the verified source tables, not guessed from
        # overall collection percentages or K/M abbreviated money values.
        return _result('Подтверждённые итоговые показатели источника; год и неделя разделены.', [], 0,
                       kind='metrics', metrics=list(data.get('metrics') or []))
    return _result('', [], 0)
