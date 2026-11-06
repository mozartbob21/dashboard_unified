"""Read one live rating source, preserving the original 2026 methodology.

Only the existing fixed GET-only source adapters and pure calculations are used.
Neither the shared rating snapshot nor source settings are changed by this tool.
"""
from datetime import date, datetime, timezone
import hashlib
import json
import math

from fastapi import HTTPException

from core.roles import check_module_access
from services import report_municipalities
from services.omsu_rating import calculate, source
from services.omsu_rating.methodology import ALL_OMSU
from . import report_periods


ID = 'query:omsu-rating'
METHOD_YEAR = 2026
PERIOD_ENDS = ('03-31', '06-30', '09-30', '11-30')
UNITS = {4: '%', 5: 'баллов', 6: '%', 7: '%'}


def _access(user, tool_id):
    if tool_id != ID:
        raise ValueError('Этот инструмент рейтинга ОМСУ не подключён.')
    if not user or not user.get('is_active', True) or not check_module_access(user, 'omsu-rating'):
        raise HTTPException(403, 'Нет доступа к рейтингу ОМСУ.')


def validate(user, tool_id, params):
    _access(user, tool_id)
    if not isinstance(params, dict) or set(params) - {'indicator', 'period', 'municipality'}:
        raise ValueError('Для рейтинга доступны только показатель, отчётный период и ОМСУ.')
    if type(params.get('indicator')) is not int or params['indicator'] not in (4, 5, 6, 7):
        raise ValueError('Выберите показатель рейтинга 4, 5, 6 или 7.')
    if type(params.get('period')) is not int or params['period'] not in range(4):
        raise ValueError('Выберите период методики 2026 года — квартал, полугодие, 9 или 11 месяцев.')
    if params['indicator'] == 7 and params['period'] != 2:
        raise ValueError('Для показателя 7 методика предусматривает только 9 месяцев 2026 года.')
    city = params.get('municipality', '')
    if not isinstance(city, str) or len(city) > 160:
        raise ValueError('Укажите одно название ОМСУ из справочника рейтинга.')
    city = city.strip()
    if city:
        cities = {report_municipalities.key(value): value for value in ALL_OMSU}
        city = cities.get(report_municipalities.key(city))
        if city is None:
            raise ValueError('ОМСУ отсутствует в справочнике рейтинга. Уточните название муниципалитета.')
    return {'indicator': params['indicator'], 'period': params['period'], 'municipality': city}


def _finite(value):
    return type(value) in (int, float) and math.isfinite(value)


def execute(user, tool_id, params, on_progress=None):
    params = validate(user, tool_id, params)
    indicator, period = params['indicator'], params['period']
    config = calculate.CONFIG[indicator]
    label = calculate.PERIODS[period] + ' 2026 года'
    result = {'tool_id': ID, 'module': 'omsu-rating',
              'source': {'id': ID, 'title': 'Рейтинг ОМСУ — ' + config['name'], 'url': '/omsu-rating'},
              'status': 'missing', 'coverage': 'missing', 'origin': 'live',
              'requested_period': {'label': label}, 'effective_period': None,
              'period_label': label, 'methodology_year': METHOD_YEAR,
              'filters': dict(params), 'metrics': [], 'rating_rows': [], 'warnings': [],
              'completeness_verified': False, 'as_of': None}

    def stage(message):
        if callable(on_progress):
            on_progress(message)

    try:
        stage('Получаем текущие данные источника рейтинга ОМСУ')
        if indicator < 7:
            from services.auth.source_api_keys import get_source_api_key
            key = get_source_api_key('omsu_rating')
            if not key:
                result.update(error_code='RATING_ACCESS_NOT_CONFIGURED',
                              message='API-ключ источника рейтинга ОМСУ не настроен. Администратору нужно сохранить его в настройках доступа.')
                return result
            data = source.fetch_redmine(config['query'], key)
        else:
            data = source.fetch_sheet()
        stamp = datetime.now(timezone.utc).isoformat()
        stage('Рассчитываем рейтинг по исходной методике 2026 года')
        # Full groups must be ranked and colored before municipality filtering.
        # The original calculation also supplies period-to-period place deltas.
        built = calculate.build({str(indicator): {'data': data, 'updated_at': stamp}})
        item = built['indicators'][str(indicator)]
        rows = item['periods'][period]
        result['snapshot_id'] = hashlib.sha256(json.dumps(data, ensure_ascii=False, sort_keys=True,
                                                         allow_nan=False, separators=(',', ':')).encode()).hexdigest()
    except Exception:
        # SourceErrors are already reviewed, but a new/changed upstream adapter
        # must not expose raw bodies, API keys or network URLs through chat.
        result.update(error_code='RATING_SOURCE_UNAVAILABLE',
                      message='Источник рейтинга ОМСУ не подтвердил данные. Проверьте доступ и состояние источника на странице блока. Сохранённые данные не подставлялись.')
        return result

    result.update(as_of=stamp, effective_period={'label': label})
    result['warnings'].append(config['note'])
    known = {report_municipalities.key(value): value for value in ALL_OMSU}
    selected, unknown = [], 0
    for row in rows:
        name = known.get(report_municipalities.key(row.get('omsu', '')))
        if name is None:
            unknown += 1
            continue
        if params['municipality'] and name != params['municipality']:
            continue
        if not _finite(row.get('value')) or type(row.get('place')) is not int:
            continue
        safe = {'municipality': name, 'value': row['value'], 'unit': UNITS[indicator],
                'place': row['place'], 'place_label': str(row['place_label']),
                'group': row['group'], 'group_size': row['group_size'], 'zone': row['zone'],
                'previous_place': row.get('previous_place'), 'place_delta': row.get('delta'),
                'zeroed': bool(row.get('no_docs')), 'provisional': bool(item['provisional'])}
        if indicator == 7:
            details = row['details']
            safe.update(source_present=any(detail.get('found') for detail in details),
                        missing_rso_count=sum(not detail.get('found') for detail in details))
        selected.append(safe)
    if unknown:
        result['warnings'].append('В источнике есть строки с ОМСУ вне справочника рейтинга. Они не показаны в ответе; исходные места и группы не пересчитывались.')
    if not selected:
        result['message'] = 'Для выбранного показателя, периода и ОМСУ источник не подтвердил значение. Отсутствие значения не считается нулём.'
        return result
    if params['municipality'] and len(selected) != 1:
        result['message'] = 'Источник содержит несколько вариантов записи выбранного ОМСУ. Однозначное значение не подтверждено; проверьте данные в блоке.'
        return result
    result['rating_rows'] = selected
    partial = bool(unknown)
    end = date.fromisoformat(str(METHOD_YEAR) + '-' + PERIOD_ENDS[period])
    if end > report_periods.today():
        partial = True
        result['warnings'].append('Отчётный период ещё не завершён. Показаны текущие значения полей источника, а не окончательные итоги периода.')
    if indicator == 7 and any(row.get('missing_rso_count') for row in selected):
        partial = True
        result['warnings'].append('Часть РСО не найдена в таблице. По исходной методике отсутствие подтверждения учитывается нулём; это не подтверждённое отсутствие выполненных работ.')
    if params['municipality']:
        row = selected[0]
        result['metrics'] = [
            {'id': 'rating_value', 'label': config['metric'], 'value': row['value'], 'unit': row['unit'],
             'definition': f'omsu-rating:methodology2026:k{indicator}:value'},
            {'id': 'rating_place', 'label': 'Место в группе ЗАТО' if row['group'] == 'zato' else 'Место в основной группе',
             'value': row['place'], 'unit': 'место', 'definition': f'omsu-rating:methodology2026:k{indicator}:place:' + row['group']},
        ]
    else:
        result['metrics'] = [{'id': 'rating_municipalities', 'label': 'ОМСУ с рассчитанным значением',
                              'value': len(selected), 'unit': 'ОМСУ',
                              'definition': f'omsu-rating:methodology2026:k{indicator}:municipality-count'}]
    result.update(status='partial' if partial else 'ok', coverage='partial' if partial else 'full',
                  completeness_verified=not partial and not item['provisional'])
    result['warnings'].append('Данные запрошены из источника для этого расчёта. Общий сохранённый рейтинг и настройки блока не изменялись.')
    return result
