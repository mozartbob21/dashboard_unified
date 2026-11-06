"""Fresh current overdue counts using the existing DataLens collector in isolation.

This source has no implemented historical/date filter. It cannot establish how
many tasks were overdue in a past month. Per-request work files and browser
profile are temporary; the dashboard's saved result and settings stay untouched.
"""
from datetime import datetime
import hashlib
import json
import re
from tempfile import TemporaryDirectory
from zoneinfo import ZoneInfo

from core.roles import check_module_access
from . import block_execution, complaint_causes
from services import report_municipalities as cities

ID = 'query:overdue'
MAX_WRAPPERS = 256
MAX_BYTES = 64 * 1024 * 1024
MAX_NODES = 100000
MAX_ROWS = 2048


def validate(user, tool_id, params):
    if tool_id != ID:
        raise block_execution.InvalidParameters('Этот источник не поддерживается.')
    if not user or not user.get('is_active', True) or not check_module_access(user, 'overdue'):
        raise block_execution.AccessDenied('Нет доступа к просроченным задачам.')
    if not isinstance(params, dict) or set(params) - {'municipality'}:
        raise block_execution.InvalidParameters('Источник просроченных задач поддерживает только текущее состояние и фильтр ОМСУ. Исторический период в этом источнике не подключён.')
    name = params.get('municipality', '')
    if not isinstance(name, str) or len(name) > 160:
        raise block_execution.InvalidParameters('Укажите одно название ОМСУ из справочника.')
    name = name.strip()
    if name:
        name = complaint_causes._known_cities().get(cities.key(name))
        if not name:
            raise block_execution.InvalidParameters('ОМСУ не найден в справочнике. Уточните официальное название.')
    return {'municipality': name}


def _objects(value):
    pending = [(value, 0)]
    visited = 0
    while pending:
        current, depth = pending.pop()
        visited += 1
        if visited > MAX_NODES or depth > 30:
            raise ValueError('Source structure limit')
        if isinstance(current, dict):
            yield current
            pending.extend((item, depth + 1) for item in current.values() if isinstance(item, (list, dict)))
        elif isinstance(current, list):
            pending.extend((item, depth + 1) for item in current if isinstance(item, (list, dict)))


def _count(value):
    if type(value) is int and 0 <= value <= 10**9:
        return value
    if isinstance(value, str) and re.fullmatch(r'\d[\d \u00a0\u202f]*', value.strip()):
        value = int(re.sub(r'\s', '', value))
        if value <= 10**9:
            return value
    raise ValueError('Unconfirmed count')


def _rows(payload):
    """Parse the same reviewed OMSU table/bar-chart shapes without null-to-zero coercion."""
    candidates = []
    for obj in _objects(payload):
        columns, data = obj.get('columns'), obj.get('data')
        if isinstance(columns, list) and isinstance(data, list) and columns and data:
            names = [str(column.get('name') or column.get('title') or '').lower() if isinstance(column, dict)
                     else str(column).lower() for column in columns]
            municipality = [i for i, name in enumerate(names) if 'омсу' in name or 'муниципал' in name]
            count = [i for i, name in enumerate(names) if 'кол-во' in name or 'количество' in name or 'задач' in name]
            if len(municipality) != 1 or len(count) != 1:
                continue
            if len(data) > MAX_ROWS:
                raise ValueError('Source row limit')
            rows = []
            for row in data:
                if not isinstance(row, (list, tuple)) or len(row) != len(columns):
                    raise ValueError('Incomplete source row')
                name = row[municipality[0]]
                if not isinstance(name, str) or not name.strip() or len(name) > 160:
                    raise ValueError('Invalid municipality')
                if name.strip().casefold().startswith('итого'):
                    continue
                rows.append((name.strip(), _count(row[count[0]])))
            if rows:
                candidates.append(rows)
        categories, graphs = obj.get('categories'), obj.get('graphs')
        if isinstance(categories, list) and isinstance(graphs, list) and categories and graphs:
            if len(graphs) != 1 or not isinstance(graphs[0], dict):
                continue
            values = graphs[0].get('data')
            if not isinstance(values, list) or len(values) != len(categories) or len(values) > MAX_ROWS:
                raise ValueError('Incomplete source chart')
            rows = []
            for name, value in zip(categories, values):
                if not isinstance(name, str) or not name.strip() or len(name) > 160:
                    raise ValueError('Invalid municipality')
                if isinstance(value, dict):
                    value = next((value[key] for key in ('y', 'value', 'count') if key in value), None)
                rows.append((name.strip(), _count(value)))
            candidates.append(rows)
    if not candidates:
        raise ValueError('No confirmed source table')
    # Several differently aggregated tables are not a safe basis for a total.
    unique = {tuple(sorted(rows)) for rows in candidates}
    if len(unique) != 1:
        raise ValueError('Ambiguous source table')
    rows = list(unique.pop())
    if len({cities.key(name) for name, _ in rows}) != len(rows):
        raise ValueError('Duplicate municipality')
    return rows


def _fresh_rows(paths, fetched):
    from services.overdue import service
    if not isinstance(fetched, dict) or not fetched.get('public_entry_seen') or not fetched.get('chart_run_count') or fetched.get('capture_limit_hit'):
        raise ValueError('Source loading not confirmed')
    files = list(paths.responses.glob('*.json'))
    if not files or len(files) > MAX_WRAPPERS or sum(path.stat().st_size for path in files) > MAX_BYTES:
        raise ValueError('Source response limit')
    wrappers = service.load_wrappers(paths)
    chart_id = service.find_omsu_chart_id(wrappers)
    if not chart_id:
        raise ValueError('OMSU widget not confirmed')
    results = []
    for wrapper in wrappers:
        if '/charts/api/run' not in str(wrapper.get('url', '')) or wrapper.get('status') != 200:
            continue
        payload = wrapper.get('data')
        if not isinstance(payload, dict):
            continue
        conf = payload.get('_confStorageConfig') or {}
        ident = payload.get('id') or (conf.get('entryId') if isinstance(conf, dict) else None) or payload.get('entryId') or payload.get('chartId')
        if ident == chart_id:
            results.append(_rows(payload))
    if not results or any(value != results[-1] for value in results):
        raise ValueError('Source widget not confirmed consistently')
    return results[-1]


def execute(user, tool_id, params, on_progress=None):
    params = validate(user, tool_id, params)
    from services.overdue import service
    result = {'tool_id': ID, 'module': 'overdue',
              'source': {'id': ID, 'title': 'Просроченные задачи — текущий источник', 'url': '/overdue'},
              'status': 'missing', 'coverage': 'snapshot', 'filters': params,
              'requested_period': None, 'effective_period': None, 'available_period': None,
              'metrics': [], 'warnings': ['Источник показывает текущее состояние. Исторические итоги за месяц или год этим запросом не подтверждаются.'],
              'completeness_verified': False, 'as_of': None}
    try:
        if callable(on_progress):
            on_progress('Получение текущих данных источника просроченных задач')
        with TemporaryDirectory(prefix='neurona-overdue-') as directory:
            paths = service.RunPaths.isolated(directory)
            fetched = service.fetch_dashboard_data(paths, headless=True, quiet=True,
                                                    max_capture_bytes=MAX_BYTES, max_wait_seconds=180)
            rows = _fresh_rows(paths, fetched)
        # Do not accept an absent row as proof of zero overdue tasks.
        matches = [(name, value) for name, value in rows if not params['municipality'] or cities.key(name) == cities.key(params['municipality'])]
        if not matches:
            result['warnings'].append('В свежем ответе источника нет строки выбранного ОМСУ. Нулевое значение не подтверждено.')
            return result
        result['status'] = 'ok'
        result['as_of'] = datetime.now(ZoneInfo('Europe/Moscow')).isoformat(timespec='seconds')
        result['snapshot_id'] = hashlib.sha256(json.dumps(rows, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
        result['metrics'] = [{'id': 'overdue_count', 'label': 'Текущие просроченные задачи', 'value': sum(value for _, value in matches),
                              'unit': 'задач', 'definition': 'query:overdue:current_omsu_count_v1'}]
        result['warnings'].append('Время указано для получения ответа, а не для исторического периода задач. Общие настройки и сохранённый результат блока не изменены.')
        if callable(on_progress):
            on_progress('Свежий ответ источника проверен; ОМСУ отфильтрован')
    except Exception:
        result['warnings'].append('Не удалось подтвердить свежую таблицу источника. Проверьте доступ к дашборду и повторите запрос; прежние файлы не использовались.')
    return result
