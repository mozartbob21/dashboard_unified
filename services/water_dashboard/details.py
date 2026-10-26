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
KONTUR_COMPLETION_HEADER = 'Процент выполнения дорожной карты'
KONTUR_TASK_HEADER = 'Ссылка на задачи дорожной карты'


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


def _controller_workload(data):
    records = defaultdict(list)
    required = ('контролер', '#', 'контроль сроков', 'категория')
    for table in data.get('tables') or []:
        headers = [_normal(header) for header in table.get('headers') or []]
        if not all(key in headers for key in required):
            continue
        indices = [headers.index(key) for key in required]
        for row in table.get('rows') or []:
            cells = [row[index] if index < len(row) else None for index in indices]
            name = _name(cells[0])
            if not name or _normal(name) in ('null', 'none', '—', '-'):
                continue
            count = _count(cells[1])
            record = None if count is None else {
                'name': name, 'value': count, 'unit': 'задач',
                'secondary': [_secondary('Контроль сроков', cells[2]),
                              _secondary('Категории', cells[3])],
            }
            if record not in records[_normal(name)]:
                records[_normal(name)].append(record)
    items = [rows[0] for rows in records.values() if len(rows) == 1 and rows[0] is not None]
    result = _result(
        'Количество задач, ожидающих проверки у каждого контролёра. '
        'Это объём задач на проверке, а не число просроченных задач. При равенстве — по алфавиту.',
        items, len(records) - len(items),
        note='Статус контроля сроков и категории приведены из таблицы источника.',
    )
    result['groups'][0].update(id='workload', title='Топ-5 по числу задач на проверке')
    result['entity_type'] = 'controller'
    return result


def is_kontur_ranking_table(table):
    headers = {_normal(header) for header in table.get('headers') or []}
    return (
        'рсо' in headers
        and bool(headers & set(MUNICIPALITY_HEADERS))
        and _normal(KONTUR_COMPLETION_HEADER) in headers
        and _normal(KONTUR_TASK_HEADER) in headers
    )


def _kontur_rankings(data):
    records = defaultdict(list)
    # Partial pagination is not a valid basis for an overall top-five.
    tables = [] if data.get('ranking_error') else data.get('ranking_tables') or []
    for table in tables:
        if not is_kontur_ranking_table(table):
            continue
        headers = [_normal(header) for header in table.get('headers') or []]
        rso_index = headers.index('рсо')
        city_index = next(headers.index(key) for key in MUNICIPALITY_HEADERS if key in headers)
        value_index = headers.index(_normal(KONTUR_COMPLETION_HEADER))
        task_index = headers.index(_normal(KONTUR_TASK_HEADER))
        for row in table.get('rows') or []:
            cell = lambda index: row[index] if index < len(row) else None
            # Preserve legal abbreviations (АО, МУП, РСО) and the source's
            # organisation spelling; title-casing is only for municipalities.
            name = ' '.join(str(cell(rso_index) or '').split())
            municipality = _name(cell(city_index))
            task = ' '.join(str(cell(task_index) or '').split())
            if (not name or _normal(name).startswith(('итого', 'всего'))
                    or not municipality or not task.isdecimal()):
                continue
            value = _percent(cell(value_index), bounded=True)
            record = None if value is None else {
                'name': name, 'municipality': municipality, 'task_id': task,
                'value': value, 'unit': '%',
                'secondary': [_secondary('ОМСУ', municipality), _secondary('Задача ДК', task)],
            }
            # The task distinguishes branches of one RSO in the same municipality.
            key = (_normal(name), _normal(municipality), task)
            if record not in records[key]:
                records[key].append(record)
    items = [rows[0] for rows in records.values() if len(rows) == 1 and rows[0] is not None]
    reason = data.get('ranking_error') or 'Полная таблица дорожных карт ещё не получена. Обновите этот источник.'
    result = _result(
        'Процент выполнения дорожной карты по РСО. Выше процент — лучше. '
        'При равенстве — по алфавиту. Организации и их подразделения различаются по ОМСУ и задаче дорожной карты.',
        items, len(records) - len(items), higher_is_better=True, best=True,
        note=(f'Рейтинг среди {len(items)} дорожных карт с распознаваемыми данными в полной таблице источника.'
              if items else reason),
    )
    result['entity_type'] = 'rso'
    result['source_url'] = data.get('ranking_url', '')
    return result


KONTUR_STAGE_LABELS = (
    'Заключение соглашения', 'Регистрация пользователя',
    'Реестр ВЗУ и датчики', 'Объекты и координаты',
    'Тестирование API', 'Передача данных датчиков',
)
KONTUR_STAGE_HEADERS = (
    '1. Заключение Соглашения', '2. Регистрация пользователя',
    '3. Актуализация реестра ВЗУ и добавления датчиков',
    '4. Актуализация объектов и определение координат в системе ЖКХ Контур',
    '5. Тестирование API по получению значений датчиков',
    '6. Транслировать данные с датчиков каждого объекта из реестра в систему ЖКХ Контур',
)


def kontur_stages(data):
    """Count RSO/branches independently at each named stage, never task totals."""
    if data.get('ranking_error'):
        return []
    records = defaultdict(set)
    for table in data.get('ranking_tables') or []:
        if not is_kontur_ranking_table(table):
            continue
        headers = [_normal(h) for h in table.get('headers') or []]
        if any(_normal(h) not in headers for h in KONTUR_STAGE_HEADERS):
            return []  # An old/partial schema cannot imply zero unfinished RSO.
        name_i = headers.index('рсо')
        city_i = next(headers.index(h) for h in MUNICIPALITY_HEADERS if h in headers)
        task_i = headers.index(_normal(KONTUR_TASK_HEADER))
        indices = [headers.index(_normal(h)) for h in KONTUR_STAGE_HEADERS]
        for row in table.get('rows') or []:
            cell = lambda i: _normal(row[i]) if i < len(row) else ''
            if not cell(name_i) or not _name(cell(city_i)) or not cell(task_i).isdecimal():
                continue
            records[(cell(name_i), cell(city_i), cell(task_i))].add(tuple(cell(i) for i in indices))
    if not records:
        return []
    accepted = [next(iter(v)) for v in records.values() if len(v) == 1]
    statuses = ('Закрыта', 'Новая', 'В работе', 'На проверке', 'Возвращена на доработку', 'Просрочена', 'Получен отказ')
    result = []
    for i, label in enumerate(KONTUR_STAGE_LABELS):
        counts = {status: sum(row[i] == _normal(status) for row in accepted) for status in statuses}
        unknown = len(records) - sum(counts.values())
        result.append({'id': f'stage_{i+1}', 'label': label, 'total': len(records),
                       'closed': counts['Закрыта'],
                       'pending': None if unknown else len(records) - counts['Закрыта'],
                       'unknown': unknown, 'statuses': counts})
    return result


def is_meeting_rso_table(table):
    headers = {_normal(h) for h in table.get('headers') or []}
    return 'рсо' in headers and bool(headers & set(MUNICIPALITY_HEADERS)) and bool(headers & set(ATTENDANCE_HEADERS))


def meeting_rso_records(data):
    if data.get('error'):
        return [], 0
    records = defaultdict(list)
    for table in data.get('tables') or []:
        if not is_meeting_rso_table(table):
            continue
        headers = [_normal(h) for h in table['headers']]
        name_i = headers.index('рсо')
        city_i = next(headers.index(h) for h in MUNICIPALITY_HEADERS if h in headers)
        att_i = next(headers.index(h) for h in ATTENDANCE_HEADERS if h in headers)
        for row in table.get('rows') or []:
            cell = lambda i: row[i] if i < len(row) else None
            name, city = ' '.join(str(cell(name_i) or '').split()), _name(cell(city_i))
            if not _name(name) or not city:
                continue
            value = _percent(cell(att_i), bounded=True)
            item = None if value is None else {'name': name, 'municipality': city, 'value': value,
                       'unit': '%', 'secondary': [_secondary('ОМСУ', city)]}
            key = (_normal(name), _normal(city))
            if item not in records[key]:
                records[key].append(item)
    items = [v[0] for v in records.values() if len(v) == 1 and v[0] is not None]
    return items, len(records) - len(items)


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
    if sid == 'controller_tasks':
        return _controller_workload(data)
    if sid == 'energy_audit':
        return _result(
            'Показатели энергоаудита из подписанных индикаторов источника. '
            'Все приборы и приборы с потреблением свыше 10 000 кВт·ч в месяц показаны отдельно; '
            'значения и единицы измерения приведены как в первоисточнике.',
            [], 0, kind='metrics', metrics=list(data.get('metrics') or []),
        )
    if sid == 'gkh_kontur':
        result = _kontur_rankings(data)
        result['stages'] = kontur_stages(data)
        return result
    if sid == 'meetings':
        items, excluded = _unique(data, {'attendance': ATTENDANCE_HEADERS}, lambda row:
            {'name': row['name'], 'value': _percent(row['attendance'], bounded=True), 'unit': '%', 'secondary': []}
            if _percent(row['attendance'], bounded=True) is not None else None)
        result = _result('Доля присутствия на перекличках: ниже явка — ниже результат. '
                       'ОМСУ и РСО оцениваются отдельно. При равенстве — по алфавиту.', items, excluded, higher_is_better=True)
        result['groups'][0]['title'] = 'Топ-5 отстающих ОМСУ'
        rso, rso_excluded = meeting_rso_records(data.get('rso') or {})
        result['groups'].append({'id': 'rso_worst', 'title': 'Топ-5 отстающих РСО',
            'items': sorted(rso, key=lambda r: (r['value'], _normal(r['name']), _normal(r['municipality'])))[:5]})
        result['rso_entities'] = rso
        result['rso_coverage'] = {'entities': len(rso), 'excluded': rso_excluded}
        result['rso_error'] = (data.get('rso') or {}).get('error', '')
        return result
    if sid == 'nvos':
        # Keep all typed measurements for compact source detail and AI reports.
        # Municipal data remain in the verified source tables, not guessed from
        # overall collection percentages or K/M abbreviated money values.
        return _result('Подтверждённые итоговые показатели источника; год и неделя разделены.', [], 0,
                       kind='metrics', metrics=list(data.get('metrics') or []))
    return _result('', [], 0)
