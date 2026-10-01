"""Safe saved reporting projections for EDDS and the MinZhKH dashboard."""
import gzip
import json
import os
import tempfile
from collections import Counter
from datetime import date, datetime, timezone
from functools import lru_cache
from pathlib import Path

from services.water_ai_context import normalized, text, finite_number
from services.report_municipalities import key as municipality_key, display as municipality_display

BASE = Path(__file__).resolve().parents[2]
MAX_BYTES = 60 * 1024 * 1024


@lru_cache(maxsize=8)
def _read_version(path, modified, size):
    try:
        if size > MAX_BYTES:
            return {}
        opener = gzip.open if str(path).endswith('.gz') else open
        with opener(path, 'rb') as stream:
            content = stream.read(MAX_BYTES + 1)
        if len(content) > MAX_BYTES:
            return {}
        result = json.loads(content.decode('utf-8-sig'))
        return result if isinstance(result, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


def _load(path):
    try:
        stat = path.stat()
        return _read_version(str(path), stat.st_mtime_ns, stat.st_size)
    except OSError:
        return {}


def _integer(v):
    return v if type(v) is int and v >= 0 else None


def _base(sid, module, title, stamp='', warning=''):
    return {'id': sid, 'module': module, 'title': title, 'source_url': '/' + module,
            'collected_at': text(stamp, 80), 'status': 'saved' if stamp else 'missing',
            'warning': warning, 'scope': 'module', 'metrics': [], 'rows': []}


@lru_cache(maxsize=8192)
def _same(a, b):
    return municipality_key(a) == municipality_key(b)


def reset_comparisons():
    # Resolve aliases again for every report, then reuse identical city pairs.
    _same.cache_clear()


def _daily(data, municipality):
    result = _base('edds-complaints', 'edds', 'ЕДДС: жалобы Добродела по воде', data.get('updated'),
                   'Это свод жалоб, а не число технологических инцидентов АРМ ЕДДС.')
    days = data.get('days')
    if not isinstance(days, dict):
        result.update(status='missing', warning='Свод жалоб Добродела ещё не сохранён. Откройте ЕДДС и обновите жалобы.')
        return result
    sums = [0, 0, 0]
    known_days, represented_days, invalid = [], [], 0
    for day, cities in days.items():
        try:
            date.fromisoformat(day)
        except (TypeError, ValueError):
            continue
        if not isinstance(cities, dict):
            continue
        known_days.append(day)
        for name, values in cities.items():
            if municipality and not _same(name, municipality):
                continue
            if not isinstance(values, list) or len(values) != 3 or any(_integer(v) is None for v in values):
                invalid += 1
                continue
            sums = [a + b for a, b in zip(sums, values)]
            if day not in represented_days:
                represented_days.append(day)
    result['scope'] = 'municipality' if municipality else 'module'
    result['period'] = {'from': min(known_days) if known_days else None, 'to': max(known_days) if known_days else None,
                        'days_present': len(represented_days) if municipality else len(known_days), 'all_saved_days': len(known_days)}
    result['metrics'] = [{'label': label, 'value': value if known_days and (not municipality or represented_days) and not invalid else None, 'unit': ''}
                         for label, value in zip(['Жалобы ХВС за сохранённые дни', 'Жалобы водоотведения за сохранённые дни', 'Жалобы ГВС за сохранённые дни'], sums)]
    result['warning'] += ' Сумма только по сохранённым дням; отсутствие дня не означает ноль жалоб.'
    if municipality and not represented_days:
        result['warning'] += ' Для выбранного муниципалитета нет отдельных записей; число жалоб не определено.'
    if invalid:
        result['warning'] += ' Есть некорректные строки; суммы не определены.'
    return result


def project_mingkh_dataset(data):
    """Aggregate before persistence; no complaint text or account identity."""
    dims, periods = data.get('dims'), data.get('rows')
    if not isinstance(dims, dict) or not isinstance(periods, dict) or not isinstance(dims.get('omsu'), list):
        raise ValueError('Unrecognized MinZhKH dataset')
    out = {'schema_version': 1, 'collected_at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
           'source_updated': text(data.get('updated'), 100), 'periods': {}, 'excluded': {}, 'unknown_dates': {}}
    for key in ('curr', 'prev', 'appg'):
        raw_rows = periods.get(key)
        if not isinstance(raw_rows, list):
            continue
        by_municipality = Counter()
        bad = 0
        for row in raw_rows:
            if not isinstance(row, list) or not row or type(row[0]) is not int or not 0 <= row[0] < len(dims['omsu']):
                bad += 1
                continue
            by_municipality[municipality_display(text(dims['omsu'][row[0]], 240))] += 1
        bounds = (data.get('bounds') or {}).get(key) or []
        out['periods'][key] = {'bounds': [text(v, 30) for v in bounds[:2]], 'total': len(raw_rows) if not bad else None,
                              'municipalities': dict(by_municipality), 'invalid_rows': bad}
        out['excluded'][key] = _integer((data.get('excluded') or {}).get(key))
        out['unknown_dates'][key] = _integer((data.get('unknown_dates') or {}).get(key))
    if 'curr' not in out['periods']:
        raise ValueError('No current period')
    return out


def persist_mingkh_dataset(data):
    """Called only after a successful portal dataset build; atomic, best effort."""
    temporary = None
    try:
        result = project_mingkh_dataset(data)
        target = BASE / 'data/mingkh/report-summary.json'
        target.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix='.report-summary-', dir=target.parent)
        with os.fdopen(fd, 'w', encoding='utf-8') as out:
            json.dump(result, out, ensure_ascii=False, allow_nan=False)
        os.replace(temporary, target)
        return True
    except (OSError, ValueError, TypeError):
        return False
    finally:
        if temporary and os.path.exists(temporary):
            os.unlink(temporary)


def _mingkh_summary(data, municipality):
    result = _base('mingkh-appeals', 'mingkh', 'МИНЖКХ: обращения по периодам', data.get('collected_at'))
    result['data_date'] = text(data.get('source_updated'), 100)
    result['scope'] = 'municipality' if municipality else 'module'
    result['periods'] = {}
    if data.get('schema_version') != 1:
        result.update(status='missing', warning='Свод дашборда ещё не сохранён. Откройте МИНЖКХ и обновите обращения; после этого свод доступен чату и после перезапуска.')
        return result
    for key, label in [('curr', 'Обращения: текущий период'), ('prev', 'Обращения: предыдущий период'), ('appg', 'Обращения: аналогичный период прошлого года')]:
        item = data.get('periods', {}).get(key)
        if not isinstance(item, dict):
            continue
        result['periods'][key] = list(item.get('bounds', []))
        if municipality:
            present = [v for name, v in item.get('municipalities', {}).items() if _same(name, municipality) and _integer(v) is not None]
            value = sum(present) if present else None
            if item.get('invalid_rows'):
                value = None
        else:
            value = _integer(item.get('total'))
        result['metrics'].append({'label': label, 'value': value, 'unit': ''})
    result['warning'] = 'Периоды соответствуют последней успешной загрузке дашборда; они могут отличаться от сегодняшней недели.'
    if municipality and any(m['value'] is None for m in result['metrics']):
        result['warning'] += ' В части периодов нет отдельных записей муниципалитета; отсутствие покрытия не считается нулём.'
    return result


def _map(data, municipality):
    meta = data.get('meta') or {}
    result = _base('mingkh-water-map', 'mingkh', 'МИНЖКХ: сохранённая карта жалоб', meta.get('live_updated') or meta.get('updated'))
    result['source_url'] = '/mingkh/water-map'
    result['period'] = {'from': text(meta.get('from'), 30), 'to': text(meta.get('to'), 30)}
    result['scope'] = 'municipality' if municipality else 'module'
    cols = data.get('cols', [])
    if not isinstance(cols, list) or 'omsu' not in cols or 'kind' not in cols or not isinstance(data.get('rows'), list):
        result.update(status='missing', warning='Сохранённая карта жалоб недоступна.')
        return result
    muni_idx, kind_idx = cols.index('omsu'), cols.index('kind')
    municipalities = data.get('omsu', [])
    counts = Counter()
    matched = 0
    bad = 0
    for row in data['rows']:
        if not isinstance(row, list) or len(row) <= max(muni_idx, kind_idx):
            bad += 1
            continue
        idx, kind = row[muni_idx], row[kind_idx]
        if type(idx) is not int or not 0 <= idx < len(municipalities) or kind not in (0, 1, 2, 3, 4):
            bad += 1
            continue
        if not municipality or _same(municipalities[idx], municipality):
            counts[kind] += 1
            matched += 1
    labels = ['ХВС', 'Водоотведение', 'ГВС', 'Капремонт', 'Прочее']
    result['metrics'] = [{'label': 'На карте: ' + label, 'value': counts[i] if not bad and (not municipality or matched) else None, 'unit': ''} for i, label in enumerate(labels)]
    result['warning'] = 'Данные только сохранённой карты, с её фильтрами и периодом; это не полный итог дашборда обращений.'
    if municipality and not matched:
        result['warning'] += ' Муниципалитет не представлен отдельными строками; показатели не определены.'
    if meta.get('archive'):
        result['status'] = 'stale'
        result['warning'] += ' Использован импортированный архив; обновление портала не подтверждено.'
    if bad:
        result['warning'] += ' Есть неполные строки, суммы не определены.'
    return result


def _zips(data_dir, municipality):
    data = _load(data_dir / 'zip_curator/published.json')
    report = _base('zips-published', 'zips', 'ЗиП: опубликованные остатки РСО', data.get('published_at'))
    grid = data.get('rows') or []
    required = ['РСО', 'Округ', 'Наименование', 'Количество', 'Ед.']
    if not grid or not isinstance(grid[0], list) or not all(h in grid[0] for h in required):
        report.update(status='missing', warning='Опубликованный свод остатков ещё не сохранён.')
        return [report]
    positions = {key: grid[0].index(key) for key in required}
    rows = [r for r in grid[1:] if isinstance(r, list) and len(r) > max(positions.values())
            and (not municipality or _same(r[positions['Округ']], municipality))]
    report['scope'] = 'municipality' if municipality else 'module'
    report['metrics'] = [
        {'label': 'Позиций в опубликованном своде', 'value': len(rows), 'unit': ''},
        {'label': 'РСО в опубликованном своде', 'value': len({text(r[positions['РСО']]) for r in rows}), 'unit': ''},
        {'label': 'Позиций без единицы измерения', 'value': sum(not text(r[positions['Ед.']]) for r in rows), 'unit': ''}]
    report['rows'] = [{key: text(row[i], 180) for key, i in positions.items()} for row in rows[:10]]
    report['omitted_rows'] = max(0, len(rows) - 10)
    report['warning'] = 'Опубликованные остатки; неподтверждённые загрузки не включены. Количества с разными единицами измерения не суммируются.'
    if municipality and not rows:
        report['warning'] += ' Отдельных строк муниципалитета нет; это не подтверждает отсутствие запасов.'
    return [report]


def _arm_value(key, value):
    if key == 'people_known_sum':
        number = finite_number(value)
        return number if number is not None and number >= 0 else None
    return _integer(value)


def _arm(data, municipality):
    result = _base('edds-arm', 'edds', 'АРМ ЕДДС: технологические инциденты ВС/ВО', data.get('collected_at'))
    result['scope'] = 'municipality' if municipality else 'module'
    result['period'] = {k: text((data.get('period') or {}).get(k), 30) for k in ('from', 'to')}
    if data.get('schema_version') != 1 or not isinstance(data.get('totals'), dict):
        result.update(status='missing', warning='Серверный свод инцидентов АРМ ещё не сохранён. Откройте ЕДДС и обновите период; жалобы Добродела не подставляются вместо инцидентов.')
        return result
    counts = data['totals']
    if municipality:
        matches = [v for name, v in data.get('municipalities', {}).items() if _same(name, municipality) and isinstance(v, dict)]
        keys = ('incidents', 'closed', 'active', 'people_known_sum', 'people_missing')
        counts = {k: sum(v[k] for v in matches) if matches and all(_arm_value(k, v.get(k)) is not None for v in matches) else None for k in keys}
        if not matches:
            result['warning'] = 'В своде нет отдельной строки муниципалитета. Показатели неизвестны, не нулевые.'
    labels = [('incidents', 'Инциденты ВС/ВО'), ('closed', 'Закрытые инциденты'), ('active', 'Незакрытые инциденты'),
              ('people_known_sum', 'Жители по заполненным заявкам (не уникальные)'), ('people_missing', 'Заявки без числа жителей')]
    result['metrics'] = [{'label': label, 'value': _arm_value(k, counts.get(k)), 'unit': ''} for k, label in labels]
    broken = _integer(data.get('broken_rows')) or 0
    unknown = _integer(data.get('unknown_municipality')) or 0
    if broken or unknown:
        result['warning'] += f' Пропущено повреждённых строк: {broken}; инцидентов без муниципалитета: {unknown}. Покрытие может быть неполным.'
    result['warning'] += ' Период — последняя серверная выгрузка АРМ, только объекты водоснабжения/водоотведения.'
    return result


def reports_for_module(module, data_dir, municipality=''):
    data_dir = Path(data_dir)
    if module == 'zips':
        return _zips(data_dir, municipality)
    if module == 'edds':
        return [_daily(_load(data_dir / 'edds/water_daily.json'), municipality),
                _arm(_load(data_dir / 'edds/arm-summary.json'), municipality)]
    return [_mingkh_summary(_load(data_dir / 'mingkh/report-summary.json'), municipality),
            _map(_load(data_dir / 'mingkh/water-map.json.gz'), municipality)]


def names_for_module(module, data_dir):
    data_dir = Path(data_dir)
    if module == 'zips':
        rows = _load(data_dir / 'zip_curator/published.json').get('rows') or []
        if rows and isinstance(rows[0], list) and 'Округ' in rows[0]:
            pos = rows[0].index('Округ')
            return {text(row[pos]) for row in rows[1:] if isinstance(row, list) and len(row) > pos}
        return set()
    if module == 'edds':
        data = _load(data_dir / 'edds/water_daily.json')
        return {text(name) for cities in (data.get('days') or {}).values() if isinstance(cities, dict) for name in cities} | set(_load(data_dir / 'edds/arm-summary.json').get('municipalities') or {})
    data = _load(data_dir / 'mingkh/report-summary.json')
    names = {text(name) for p in (data.get('periods') or {}).values() for name in p.get('municipalities', {})}
    names.update(text(name) for name in _load(data_dir / 'mingkh/water-map.json.gz').get('omsu', []))
    return names
