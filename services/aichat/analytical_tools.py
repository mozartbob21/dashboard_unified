"""Allowlisted, read-only calculations over the platform's existing water stores.

The model never selects a path, SQL, Python expression or source field. Only
dates, a canonical municipality and the three existing water kinds are accepted.
"""
from datetime import date, timedelta
import hashlib
import json
import re

from fastapi import HTTPException
from core.roles import check_module_access
from services import report_municipalities as cities
from services.aichat import complaint_causes, remaining_sources, report_periods
from services.mingkh import water_map


SPECS = {
    'analytics:mingkh-map': ('mingkh', 'Карта жалоб МинЖКХ — вода', '/mingkh/water-map'),
    'analytics:edds-water': ('edds', 'ЕДДС — свод жалоб по воде', '/edds'),
}
SERVICES = {'all': 'ХВС, водоотведение и ГВС вместе', 'hvs': 'Холодное водоснабжение',
            'vo': 'Водоотведение', 'gvs': 'Горячее водоснабжение'}
KINDS = {'hvs': 0, 'vo': 1, 'gvs': 2}
MAX_DAYS = 366
MAX_DAILY_DAYS = 10000
MAX_MUNICIPALITIES = 1024


def _access(user, tool_id):
    if not isinstance(tool_id, str) or tool_id not in SPECS:
        raise ValueError('Этот аналитический инструмент не поддерживается.')
    if not user or not user.get('is_active', True) or not check_module_access(user, SPECS[tool_id][0]):
        raise HTTPException(403, 'Нет доступа к выбранному аналитическому блоку.')


def _day(value):
    if not isinstance(value, str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', value):
        raise ValueError('Укажите дату в формате ГГГГ-ММ-ДД.')
    return date.fromisoformat(value)


def _known():
    # Only the established municipality registry and bundled taxonomy, never
    # arbitrary strings from imported records, become displayed choices.
    return complaint_causes._known_cities()


def validate(user, tool_id, params):
    _access(user, tool_id)
    if not isinstance(params, dict) or set(params) - {'start', 'end', 'municipality', 'service'}:
        raise ValueError('Допустимы только период, ОМСУ и вид водной услуги.')
    if not {'start', 'end'} <= set(params):
        raise ValueError('Укажите начало и конец периода анализа.')
    first, last = _day(params['start']), _day(params['end'])
    today = report_periods.today()
    if first > last or (last-first).days >= MAX_DAYS or first > today or last.year > today.year:
        raise ValueError('Укажите период до 366 дней. Начало не может быть в будущем; доступны годы до текущего включительно.')
    service = params.get('service', 'all')
    if not isinstance(service, str) or service not in SERVICES:
        raise ValueError('Доступные услуги — ХВС, водоотведение, ГВС или все три вместе.')
    city = params.get('municipality', '')
    if not isinstance(city, str) or len(city) > 160:
        raise ValueError('Укажите одно название ОМСУ из справочника.')
    city = city.strip()
    if city:
        city = _known().get(cities.key(city))
        if not city:
            raise ValueError('ОМСУ не найден в справочнике. Уточните официальное название муниципалитета.')
    return {'start': first.isoformat(), 'end': last.isoformat(), 'municipality': city, 'service': service}


def _snapshot(tool_id):
    if tool_id == 'analytics:mingkh-map':
        path = water_map.DATA_FILE if water_map.DATA_FILE.is_file() else water_map.SEED_FILE
    else:
        path = remaining_sources.DATA_DIR / 'edds/water_daily.json'
    return complaint_causes._read(path)


def _bounds(data, tool_id):
    if tool_id == 'analytics:mingkh-map':
        meta = data.get('meta', {})
        first, last = _day(meta['from']), _day(meta['to'])
    else:
        days = data.get('days')
        if not isinstance(days, dict) or not days or len(days) > MAX_DAILY_DAYS:
            raise ValueError('No saved days')
        parsed = [_day(key) for key in days]
        first, last = min(parsed), max(parsed)
    if first > last:
        raise ValueError('Invalid source bounds')
    return {'from': first.isoformat(), 'to': last.isoformat()}


def capabilities(user, tool_id):
    _access(user, tool_id)
    data = _snapshot(tool_id)
    try:
        bounds = _bounds(data, tool_id)
    except (ValueError, TypeError, KeyError, AttributeError):
        bounds = None
    return {'municipalities': sorted(set(_known().values())), 'services': dict(SERVICES),
            'available_period': bounds,
            'coverage_basis': 'saved_archive_bounds' if tool_id == 'analytics:mingkh-map' else 'explicit_saved_days'}


def _result(tool_id, params):
    module, title, url = SPECS[tool_id]
    return {'tool_id': tool_id, 'module': module, 'source': {'id': tool_id, 'title': title, 'url': url},
            'status': 'missing', 'coverage': 'missing',
            'requested_period': {'from': params['start'], 'to': params['end']},
            'available_period': None, 'effective_period': None,
            'filters': {'municipality': params['municipality'], 'service': params['service'],
                        'service_label': SERVICES[params['service']]},
            'coverage_basis': 'saved_archive_bounds' if tool_id == 'analytics:mingkh-map' else 'explicit_saved_days',
            'completeness_verified': False, 'metrics': [], 'warnings': [], 'as_of': None}


def _window(result):
    available, requested = result['available_period'], result['requested_period']
    first = max(available['from'], requested['from'])
    last = min(available['to'], requested['to'], report_periods.today().isoformat())
    if first > last:
        result['warnings'].append('За запрошенный период в сохранённом источнике нет подтверждённых данных.')
        return None
    result['effective_period'] = {'from': first, 'to': last}
    if result['effective_period'] != requested:
        result['warnings'].append('Период неполный. Показана только доступная часть запрошенного диапазона.')
    return result['effective_period']


def _metric(result, value):
    service = result['filters']['service']
    result['metrics'] = [{'id': 'complaints', 'label': 'Жалобы по выбранным видам водной услуги',
                          'value': value, 'unit': 'жалоб',
                          'definition': result['tool_id'] + ':water_counts_v1:' + service}]


def _map(data, result):
    meta = data.get('meta', {})
    result['as_of'] = complaint_causes._stamp(meta.get('live_updated')) or complaint_causes._stamp(meta.get('updated'))
    result['warnings'].append('Подсчитаны записи сохранённого архива карты. Полнота поступлений вне архива не подтверждена.')
    columns, rows, names = data.get('cols'), data.get('rows'), data.get('omsu')
    if (not isinstance(columns, list) or len(columns) != len(set(columns)) or not isinstance(rows, list)
            or len(rows) > water_map.MAX_RECORDS or meta.get('count') != len(rows)
            or not isinstance(names, list) or len(names) > MAX_MUNICIPALITIES):
        raise ValueError('Invalid map structure')
    positions = {key: columns.index(key) for key in ('id', 'created', 'omsu', 'kind')}
    known = _known()
    places = [known.get(cities.key(name)) if isinstance(name, str) else None for name in names]
    city, service = result['filters']['municipality'], result['filters']['service']
    if city and city not in places:
        result['warnings'].append('Для выбранного ОМСУ в источнике нет подтверждённой записи. Это не означает отсутствие жалоб.')
        return result
    span = _window(result)
    if not span:
        return result
    kinds = {0, 1, 2} if service == 'all' else {KINDS[service]}
    seen, total, unresolved = set(), 0, 0
    for row in rows:
        if not isinstance(row, list) or len(row) != len(columns):
            raise ValueError('Invalid map row')
        ident, raw_day, place, kind = (row[positions[key]] for key in ('id', 'created', 'omsu', 'kind'))
        if type(ident) is not int or ident <= 0 or ident in seen or type(kind) is not int or kind not in range(5):
            raise ValueError('Invalid id or kind')
        seen.add(ident)
        if type(place) is not int or not 0 <= place < len(places) or not isinstance(raw_day, str):
            raise ValueError('Invalid place')
        day = _day(raw_day[:10]).isoformat()
        if not result['available_period']['from'] <= day <= result['available_period']['to']:
            raise ValueError('Row outside declared source period')
        if kind not in kinds or not span['from'] <= day <= span['to']:
            continue
        unresolved += int(places[place] is None)
        if not city or places[place] == city:
            total += 1
    complete = span == result['requested_period'] and not (city and unresolved)
    result.update(status='ok' if complete else 'partial', coverage='full' if complete else 'partial')
    if city and unresolved:
        result['warnings'].append('Часть водных жалоб не имеет подтверждённого ОМСУ. Итог по округу может быть неполным.')
    if total or complete:
        _metric(result, total)
    elif city and unresolved:
        result.update(status='missing', coverage='missing')
    else:
        # A zero in a truncated interval is still a confirmed archive zero for
        # that explicitly displayed interval, never for the requested full year.
        _metric(result, 0)
    omitted = meta.get('omitted')
    if type(omitted) is int and omitted > 0:
        result['warnings'].append('Архив применяет правила исключения карты. Всего исключённых записей — ' + str(omitted) + '.')
    return result


def _daily(data, result):
    days = data['days']
    result['as_of'] = complaint_causes._stamp(data.get('updated'))
    if data.get('kinds', ['hvs', 'vo', 'gvs']) != ['hvs', 'vo', 'gvs']:
        raise ValueError('Different daily metric ordering')
    span = _window(result)
    if not span:
        return result
    city, service = result['filters']['municipality'], result['filters']['service']
    selected = (0, 1, 2) if service == 'all' else (KINDS[service],)
    city_key = cities.key(city) if city else ''
    total, represented, missing = 0, 0, 0
    first, last = _day(span['from']), _day(span['to'])
    for offset in range((last-first).days+1):
        value = days.get((first + timedelta(days=offset)).isoformat())
        if not isinstance(value, dict) or len(value) > MAX_MUNICIPALITIES:
            missing += 1
            continue
        found = False
        daily = 0
        used = set()
        for name, counts in value.items():
            if not isinstance(name, str):
                raise ValueError('Invalid daily city')
            key = cities.key(name)
            if key in used:
                raise ValueError('Duplicate municipality alias')
            used.add(key)
            if city and key != city_key:
                continue
            if (not isinstance(counts, list) or len(counts) != 3 or
                    any(type(item) is not int or not 0 <= item <= 2**53 for item in counts)):
                raise ValueError('Invalid daily count')
            daily += sum(counts[index] for index in selected)
            if daily > 2**53:
                raise ValueError('Daily total outside safe integer range')
            found = True
        if city and not found:
            missing += 1
        else:
            total += daily
            if total > 2**53:
                raise ValueError('Total outside safe integer range')
            represented += 1
    result['days_present'] = represented
    result['days_missing'] = missing
    if not represented:
        result['warnings'].append('Нет подтверждённых дней для выбранного ОМСУ и периода. Нулевая сумма не подставлена.')
        return result
    complete = span == result['requested_period'] and missing == 0
    result.update(status='ok' if complete else 'partial', coverage='full' if complete else 'partial',
                  completeness_verified=complete)
    if missing:
        result['warnings'].append('Внутри доступного периода не подтверждено дней — ' + str(missing) +
                                  '. Сумма включает только явно сохранённые дни и записи выбранного ОМСУ.')
    result['warnings'].append('Свод ЕДДС учитывает обращения Добродела по воде, а не технологические инциденты АРМ ЕДДС.')
    _metric(result, total)
    return result


def execute(user, tool_id, params, on_progress=None):
    params = validate(user, tool_id, params)
    result = _result(tool_id, params)
    if on_progress:
        on_progress('Проверяем период и считаем записи сохранённого источника…')
    data = _snapshot(tool_id)
    try:
        result['snapshot_id'] = hashlib.sha256(json.dumps(data, ensure_ascii=False, sort_keys=True,
                                                        separators=(',', ':')).encode('utf-8')).hexdigest()
        result['available_period'] = _bounds(data, tool_id)
        return _map(data, result) if tool_id == 'analytics:mingkh-map' else _daily(data, result)
    except (ValueError, TypeError, KeyError, IndexError, AttributeError, OverflowError):
        # Never leak source payloads, paths, addresses or imported instructions.
        result.update(status='missing', coverage='missing', metrics=[], effective_period=None,
                      error_code='ANALYTICS_SOURCE_UNAVAILABLE',
                      warnings=['Сохранённый источник отсутствует или его структура не подтверждена. Откройте блок и обновите данные.'])
        return result
