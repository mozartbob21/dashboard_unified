"""Fresh fixed-source dashboard reads with disposable browser/storage state."""
from datetime import datetime, timezone
import math
from pathlib import Path
import re
from tempfile import TemporaryDirectory
from uuid import uuid4

from fastapi import HTTPException
from core.roles import check_module_access
from services import report_municipalities as cities
from services.aichat.operational_lookup import _plain
from services.mingkh.dashboard import normalize_date
from services.water_dashboard import builder, config, diagnostics
from services import water_ai_context

ID = 'query:water-dashboard'
SOURCE_IDS = tuple(item['id'] for item in config.SOURCES)
_CITY = re.compile(r'[А-ЯЁа-яёA-Za-z][А-ЯЁа-яёA-Za-z0-9 .()—–-]{0,159}')


def validate(user, params):
    if not user or not user.get('is_active', True) or not check_module_access(user, 'water-dashboard'):
        raise HTTPException(403, 'Нет доступа к сводному дашборду.')
    if not isinstance(params, dict) or set(params) - {'source', 'municipality'}:
        raise ValueError('Для сводного дашборда доступны источник и ОМСУ. Исторические даты сбора не поддерживаются.')
    if not isinstance(params.get('source'), str) or params['source'] not in SOURCE_IDS:
        raise ValueError('Выберите один из подключённых источников сводного дашборда.')
    city = params.get('municipality', '')
    if not isinstance(city, str) or (city.strip() and not _CITY.fullmatch(city.strip())):
        raise ValueError('Укажите одно название ОМСУ без команд или разметки.')
    return {'source': params['source'], 'municipality': city.strip()}


def _numeric(value):
    return type(value) in (int, float) and math.isfinite(value) and abs(value) <= 2**53-1


def _metric(sid, item, ident=None):
    value = item.get('value')
    if not _numeric(value):
        return None
    return {'id': ident or item['id'], 'label': item['label'], 'value': value, 'unit': item.get('unit') or '',
            'definition': ID + ':' + sid + ':typed_metric_v1:' + (ident or item['id'])}


def _result(params, spec):
    return {'tool_id': ID, 'module': 'water-dashboard',
            'source': {'id': ID + ':' + params['source'], 'title': spec['name'], 'url': '/water-dashboard'},
            'status': 'missing', 'coverage': 'missing', 'requested_period': None,
            'effective_period': None, 'available_period': None,
            'filters': dict(params), 'coverage_basis': 'fresh_dashboard_observation',
            'completeness_verified': False, 'comparison_consistency': 'independent_live_reads',
            'snapshot_id': uuid4().hex, 'metrics': [], 'warnings': [], 'as_of': None, 'retrieved_at': None}


def _fail(result, code, message):
    result.update(status='missing', coverage='missing', metrics=[], error_code=code, message=message)
    return result


def _partial(result, message):
    result['coverage'] = 'partial'
    result['warnings'].append(message)


def _municipal_metrics(snapshot, source, params, result):
    sid, city = params['source'], params['municipality']
    key = cities.key(city)
    details = source.get('details') or {}
    excluded = (details.get('coverage') or {}).get('excluded')
    if type(excluded) is int and excluded:
        _partial(result, 'Часть строк источника не прошла проверку. Муниципальная выборка может быть неполной.')
    if source.get('ranking_error'):
        _partial(result, 'Полная таблица организаций не получена. Показаны только подтверждённые строки текущей выгрузки.')
    if sid in water_ai_context.TABLE_FIELDS:
        fields = water_ai_context.TABLE_FIELDS[sid]
        records = [row for row in snapshot['table'] if cities.key(row.get('name')) == key]
        if len(records) != 1:
            return []
        result['filters']['municipality'] = _plain(records[0]['name'], 160)
        unit = '%' if sid == 'meetings' else 'адр.' if sid in ('sys_vs', 'sys_kr') else 'шт.'
        metrics = [m for field, label in fields.items()
                   if (m := _metric(sid, {'id': field, 'label': label, 'value': records[0].get(field), 'unit': unit}))]
        if len(metrics) != len(fields):
            _partial(result, 'Часть показателей выбранного ОМСУ не подтверждена. Отсутствующие значения не заменены нулём.')
        return metrics
    records = [item for item in details.get('entities', []) if isinstance(item, dict) and
               cities.key(item.get('municipality') or (item.get('name') if sid not in ('edo_rso', 'gkh_kontur', 'controller_tasks') else '')) == key]
    if not records:
        return []
    metrics = []
    if len(records) == 1 and sid == 'valves':
        record = records[0]
        if record.get('value') in ('Не внесено', 'Частично', 'Полностью'):
            result['warnings'].append('Статус внесения задвижек — ' + record['value'] + '.')
        metrics = [m for index, item in enumerate(record.get('secondary', []))
                   if (m := _metric(sid, item, 'municipal_' + str(index)))]
        if len(metrics) != len(record.get('secondary', [])):
            _partial(result, 'Часть числовых показателей задвижек не подтверждена. Показаны только проверенные значения.')
    elif sid in ('edo_rso', 'gkh_kontur'):
        # Source-specific organization values cannot be averaged into an invented
        # municipal percentage. Show the confirmed observations individually.
        for record in records[:10]:
            if _numeric(record.get('value')):
                result['warnings'].append('Организация «' + _plain(record.get('name'), 120) + '» — ' +
                                          str(record['value']) + ' ' + _plain(record.get('unit'), 20) + '.')
        if len(records) > 10:
            result['warnings'].append('В источнике есть ещё организаций — ' + str(len(records)-10) + '. Полный перечень доступен в блоке.')
        result['warnings'].append('Значения отдельных организаций не объединены в муниципальный процент.')
        metrics = [_metric(sid, {'id': 'organizations', 'label': 'Организаций в подтверждённой выборке ОМСУ',
                                 'value': len(records), 'unit': 'организаций'})]
    return metrics


def execute(user, params, on_progress=None):
    params = validate(user, params)
    sid = params['source']
    spec = next(item for item in config.SOURCES if item['id'] == sid)
    result = _result(params, spec)
    stage = on_progress or (lambda label: None)
    try:
        from services.water_dashboard import scraper
        stage('Запускаем сбор выбранного источника сводного дашборда')
        with TemporaryDirectory(prefix='neurona-water-query-', ignore_cleanup_errors=True) as temporary:
            raw = scraper.scrape_all(source_ids=[sid], profile_dir=Path(temporary) / 'profile',
                                     debug_dir=None, headless=True, on_progress=stage)
            stage('Проверяем показатели источника и выбранный ОМСУ')
            snapshot = builder.build_snapshot(raw, source_ids=[sid], persist=False)
        source = snapshot['sources'][sid]
        if not source.get('ok') or sid not in snapshot.get('last_updated_sources', []):
            return _fail(result, 'WATER_SOURCE_UNCONFIRMED', 'Источник не подтвердил новый итоговый показатель. Предыдущий снимок не использован. Проверьте доступность источника в сводном дашборде.')
        result['retrieved_at'] = datetime.now(timezone.utc).isoformat(timespec='seconds')
        source_date = normalize_date(source.get('data_date'))
        result['as_of'] = source_date or None
        if source_date:
            result['effective_period'] = {'from': source_date, 'to': source_date}
        else:
            result['warnings'].append('Источник не подтвердил отчётную дату. Время получения не является периодом показателя.')
        result['warnings'].append('Получен текущий срез выбранного источника с его штатными фильтрами. Произвольный исторический период не запрашивался.')
        if params['municipality']:
            metrics = _municipal_metrics(snapshot, source, params, result)
        else:
            metrics = [value for item in source['metrics'] if (value := _metric(sid, item))]
            unavailable = []
            for item in source['metrics']:
                if isinstance(item.get('value'), str) and re.fullmatch(r'[\d\s.,%/—–+−-]{1,80}', item['value']) and re.search(r'\d', item['value']):
                    result['warnings'].append(item['label'] + ' — ' + item['value'] + '.')
                elif not _numeric(item.get('value')):
                    unavailable.append(item['label'])
            if unavailable:
                _partial(result, 'Не подтверждены показатели: ' + ', '.join(unavailable) + '. Отсутствующие значения не заменены нулём.')
        if not metrics:
            names = [name for name in water_ai_context.municipality_names(snapshot) if _CITY.fullmatch(name)]
            result['available_municipalities'] = names[:30]
            return _fail(result, 'WATER_MUNICIPALITY_UNAVAILABLE', 'В новой выгрузке нет отдельного подтверждённого показателя для выбранного ОМСУ. Областной итог вместо него не подставлен.')
        complete = result['coverage'] != 'partial'
        result.update(status='ok' if complete else 'partial', coverage='full' if complete else 'partial',
                      completeness_verified=complete, metrics=metrics)
        return result
    except HTTPException:
        raise
    except Exception as error:
        # Browser errors may contain page contents, profile paths or credentials.
        code = diagnostics.error_code(error)
        return _fail(result, 'WATER_QUERY_' + code.upper(), diagnostics.MESSAGES[code])
