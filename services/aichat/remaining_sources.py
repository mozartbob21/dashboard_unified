"""Local, allowlisted reporting snapshots for ECUR and water-task checks.

No refreshes are performed by readers. Credentials, complaint text, addresses,
task descriptions and arbitrary client metadata are never persisted here.
"""
import json
import os
import tempfile
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from functools import lru_cache
from pathlib import Path

from services.report_municipalities import key as muni_key, display as muni_display

DATA_DIR = Path(__file__).resolve().parents[2] / 'data'
MAX_BYTES = 8 * 1024 * 1024
MAX_ROWS = 50_000
WATER_GROUPS = (('rez', 'close'), ('rez', 'rw'), ('sys', 'close'), ('sys', 'ext'), ('sys', 'rw'))
DEADLINE_LABELS = {'total': 'Жалоб в сохранённой выборке', 'overdue': 'Срок истёк на дату сбора',
                   'today': 'Срок в день сбора', 'week': 'Срок после дня сбора до конца недели',
                   'month': 'Срок после текущей недели до конца месяца', 'later': 'Срок позднее конца месяца',
                   'unknown_deadline': 'Не распознан срок'}


def _text(value, limit=160):
    return str(value).strip()[:limit] if isinstance(value, (str, int, float)) else ''


def _stamp():
    return datetime.now(timezone.utc).isoformat(timespec='seconds')


@lru_cache(maxsize=4)
def _read_version(path, modified, size):
    try:
        if size > MAX_BYTES:
            return {}
        result = json.loads(Path(path).read_text(encoding='utf-8-sig'))
        return result if isinstance(result, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


def _read(path):
    try:
        stat = path.stat()
        return _read_version(str(path), stat.st_mtime_ns, stat.st_size)
    except OSError:
        return {}


def _number(value):
    return value if type(value) is int and value >= 0 else None


def _atomic_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix='.report-', suffix='.json', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as stream:
            json.dump(value, stream, ensure_ascii=False, allow_nan=False)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def _date(value):
    raw = _text(value, 80).split('T')[0].split(' ')[0]
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d.%m.%y'):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def _ecur_counts():
    return {**dict.fromkeys(DEADLINE_LABELS, 0), 'status_counts': {}, 'category_counts': {}}


def _count_ecur(target, row, indices, today):
    target['total'] += 1
    deadline = _date(row[indices['Срок']])
    week_end = today + timedelta(days=6 - today.weekday())
    month_end = (today.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    key = ('unknown_deadline' if not deadline else 'overdue' if deadline < today else
           'today' if deadline == today else 'week' if deadline <= week_end else
           'month' if deadline <= month_end else 'later')
    target[key] += 1
    for field, source in [('status_counts', 'Статус'), ('category_counts', 'Категория ЕЦУР')]:
        label = _text(row[indices[source]]) or 'Не указано'
        counts = target[field]
        if label not in counts and len(counts) >= 100:
            label = 'Прочие категории' if field == 'category_counts' else 'Прочие статусы'
        counts[label] = counts.get(label, 0) + 1


def project_ecur(rows, meta=None, *, today=None):
    """Project a successful grid response. A header-only grid is a valid zero."""
    required = ('Район', 'Категория ЕЦУР', 'Срок', 'Статус')
    if not isinstance(rows, list) or not rows or not isinstance(rows[0], list) or len(rows) > MAX_ROWS + 1:
        raise ValueError('Неподдерживаемый результат ЕЦУР')
    if any(rows[0].count(field) != 1 for field in required):
        raise ValueError('В результате ЕЦУР отсутствуют обязательные колонки')
    indices = {field: rows[0].index(field) for field in required}
    current = today or date.today()
    result = {'schema_version': 1, 'collected_at': _stamp(), 'as_of': current.isoformat(),
              'selection': 'Активные жалобы по настроенной выборке ЕЦУР; срок не ранее даты сбора.',
              'totals': _ecur_counts(), 'municipalities': {}, 'unknown_municipality_rows': 0}
    for row in rows[1:]:
        if not isinstance(row, list) or len(row) <= max(indices.values()):
            raise ValueError('Неполная строка ЕЦУР; прежний снимок сохранён')
        _count_ecur(result['totals'], row, indices, current)
        raw_name = _text(row[indices['Район']], 240)
        if not raw_name or raw_name in {'—', '-'}:
            result['unknown_municipality_rows'] += 1
            continue
        name = muni_display(raw_name)
        key = muni_key(name)
        if key not in result['municipalities']:
            if len(result['municipalities']) >= 500:
                raise ValueError('Слишком много муниципалитетов в результате ЕЦУР')
            result['municipalities'][key] = {'name': name, **_ecur_counts()}
        _count_ecur(result['municipalities'][key], row, indices, current)
    return result


def persist_ecur_snapshot(rows, meta=None, *, data_dir=None):
    """Best effort: inability to write a report must not break a portal login."""
    try:
        _atomic_json(Path(data_dir or DATA_DIR) / 'ecur/report-summary.json', project_ecur(rows, meta))
        return True
    except (OSError, ValueError, TypeError):
        return False


def _municipality_registry(data_dir):
    data = _read(Path(data_dir) / 'municipality_registry.json').get('municipalities', {})
    known = {}
    if not isinstance(data, dict):
        return known
    for name, info in data.items():
        title = muni_display(name)
        known[muni_key(name)] = title
        if isinstance(info, dict) and isinstance(info.get('aliases'), list):
            for alias in info['aliases']:
                if isinstance(alias, str):
                    known[muni_key(alias)] = title
    return known


def project_water_check(payload, *, data_dir=None):
    """Accept only five classified task buckets. No free text survives."""
    if not isinstance(payload, dict):
        raise ValueError('Ожидается результат проверки задач')
    source = payload.get('result', payload)
    if not isinstance(source, dict):
        raise ValueError('Некорректный результат проверки')
    result = {'schema_version': 1, 'collected_at': _stamp(), 'kind': 'recommendations',
              'rez': {'close': [], 'rw': []}, 'sys': {'close': [], 'ext': [], 'rw': []}}
    known = _municipality_registry(data_dir or DATA_DIR)
    seen = set()
    for family, action in WATER_GROUPS:
        group = source.get(family)
        rows = group.get(action) if isinstance(group, dict) else None
        if not isinstance(rows, list) or len(rows) > MAX_ROWS:
            raise ValueError('Отсутствует раздел результата: ' + family + '.' + action)
        for row in rows:
            if not isinstance(row, dict) or type(row.get('id')) is not int or not 0 < row['id'] < 10**12:
                raise ValueError('Некорректный номер задачи')
            if row['id'] in seen:
                raise ValueError('Задача повторяется в результатах')
            seen.add(row['id'])
            if len(seen) > MAX_ROWS:
                raise ValueError('Результат содержит слишком много задач')
            # Project names are not geography; only exact registry identities
            # (including known aliases) may become a municipality.
            raw = row.get('municipality') or row.get('project')
            name = known.get(muni_key(_text(raw, 240)), '')
            result[family][action].append({'id': row['id'], 'municipality': name})
    total = payload.get('total_tasks', len(seen))
    if type(total) is not int or not len(seen) <= total <= MAX_ROWS:
        raise ValueError('Некорректное число загруженных задач')
    result['total_tasks'] = total
    result['unclassified_tasks'] = total - len(seen)
    result['unknown_municipality_rows'] = sum(not row['municipality'] for f, a in WATER_GROUPS for row in result[f][a])
    return result


def save_water_check(payload, *, data_dir=None):
    directory = Path(data_dir or DATA_DIR)
    result = project_water_check(payload, data_dir=directory)
    _atomic_json(directory / 'water_rm/last_check.json', result)
    return result


def _report(module, title, data, municipality):
    return {'id': module + '-saved', 'module': module, 'title': title,
            'source_url': '/water-rm' if module == 'water_rm' else '/ecur',
            'collected_at': _text(data.get('collected_at'), 80), 'status': 'saved',
            'scope': 'municipality' if municipality else 'module', 'metrics': [], 'rows': [], 'warning': ''}


def reports_for_module(module, data_dir, municipality=''):
    if module not in {'ecur', 'water_rm'}:
        raise ValueError('Неподдерживаемый модуль')
    path = Path(data_dir) / ('ecur/report-summary.json' if module == 'ecur' else 'water_rm/last_check.json')
    data = _read(path)
    title = 'ЕЦУР: сохранённая выборка жалоб' if module == 'ecur' else 'РМ Водоснабжение: проверка задач'
    report = _report(module, title, data, municipality)
    valid = data.get('schema_version') == 1 and isinstance(data.get('collected_at'), str) and bool(data['collected_at'])
    if module == 'ecur':
        valid = valid and isinstance(data.get('totals'), dict) and _number(data['totals'].get('total')) is not None and isinstance(data.get('municipalities'), dict)
    else:
        valid = valid and _number(data.get('total_tasks')) is not None and all(isinstance(data.get(f), dict) and isinstance(data[f].get(a), list) for f, a in WATER_GROUPS)
    if not valid:
        report.update(status='missing', warning='Результат ещё не сохранён. Откройте блок и выполните успешное обновление; отсутствие снимка не означает ноль.')
        return [report]
    if module == 'ecur':
        selected = next((value for key, value in data.get('municipalities', {}).items() if muni_key(key) == muni_key(municipality)), None) if municipality else data.get('totals')
        selected = selected if isinstance(selected, dict) else {}
        report['metrics'] = [{'label': label, 'value': _number(selected.get(key)), 'unit': ''} for key, label in DEADLINE_LABELS.items()]
        report['data_date'] = _text(data.get('as_of'), 30)
        for field in ('status_counts', 'category_counts'):
            counts = selected.get(field, {})
            report[field] = dict(Counter({_text(k): v for k, v in counts.items() if _number(v) is not None}).most_common(10)) if isinstance(counts, dict) else {}
        report['warning'] = 'Только последняя успешная выборка ЕЦУР: активные жалобы с настроенными фильтрами, срок не ранее даты сбора. Сроковые группы рассчитаны на указанную дату; это не все обращения области.'
        if municipality and not selected:
            report['warning'] += ' Отдельных строк муниципалитета нет; областной итог не подставлен, значения не определены.'
    elif module == 'water_rm':
        rows = [(family, action, row) for family, action in WATER_GROUPS for row in data.get(family, {}).get(action, []) if isinstance(row, dict)]
        selected = [(f, a, r) for f, a, r in rows if not municipality or muni_key(r.get('municipality')) == muni_key(municipality)]
        exists = bool(selected) or not municipality
        counts = Counter(action for _, action, _ in selected)
        report['metrics'] = [{'label': label, 'value': value if exists else None, 'unit': ''} for label, value in (
            ('Классифицировано задач', len(selected)), ('Рекомендовано закрыть', counts['close']),
            ('Рекомендовано продлить', counts['ext']), ('Рекомендовано доработать', counts['rw']))]
        if not municipality:
            report['metrics'].extend([{'label': 'Задач в выгрузке', 'value': _number(data.get('total_tasks')), 'unit': ''},
                                      {'label': 'Не включено в классификацию', 'value': _number(data.get('unclassified_tasks')), 'unit': ''},
                                      {'label': 'Классифицированных задач без определённого муниципалитета', 'value': _number(data.get('unknown_municipality_rows')), 'unit': ''}])
        report['warning'] = 'Результат анализа: рекомендации по задачам, а не подтверждение их применения в Redmine. Муниципалитет определяется только по точному справочному совпадению; по теме задачи география не угадывается.'
        if municipality and not exists:
            report['warning'] += ' Отдельных строк муниципалитета нет; значения не определены.'
    else:
        raise ValueError('Неподдерживаемый модуль')
    return [report]


def names_for_module(module, data_dir):
    if module == 'ecur':
        data = _read(Path(data_dir) / 'ecur/report-summary.json')
        municipalities = data.get('municipalities', {})
        return {muni_display(row.get('name', key)) for key, row in municipalities.items() if isinstance(row, dict)} if isinstance(municipalities, dict) else set()
    if module == 'water_rm':
        data = _read(Path(data_dir) / 'water_rm/last_check.json')
        names = set()
        for family, action in WATER_GROUPS:
            rows = data.get(family, {}).get(action, []) if isinstance(data.get(family), dict) else []
            if isinstance(rows, list):
                names.update(muni_display(row['municipality']) for row in rows if isinstance(row, dict) and row.get('municipality'))
        return names
    return set()
