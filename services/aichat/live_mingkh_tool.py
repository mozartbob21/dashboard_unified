"""A parameterized read of the real MinZhKH portal using its existing client.

This adapter never loads a saved report as a substitute, never calls the
dashboard's persistence/cache entry point and never changes shared filters.
"""
from collections import Counter
from datetime import date
import hashlib
import re
from uuid import uuid4

from fastapi import HTTPException
from core.roles import check_module_access
from services import report_municipalities as cities
from services.aichat import report_periods
from services.mingkh import dashboard
from services.mingkh.client import PortalError
from services.mingkh.sms.client import Pentaho
from services.mingkh.sms import filters as dashboard_filters
from services.mingkh.sms.svod_obrashcheniya import _omsu_rows

ID = 'query:mingkh-complaints'
TITLE = 'МинЖКХ — запрос жалоб из действующего источника'
URL = '/mingkh'
# Labels and filter IDs are the existing dashboard's reviewed taxonomy. Keep
# the catalog, request dispatch and response verification on this single map.
_FILTER_LABELS = {'direction': dict(dashboard_filters.DIRECTIONS),
                  'synt_group': dict(dashboard_filters.THEMES),
                  'subtopic': dict(dashboard_filters.SUBTOPICS)}
TOPICS = {'all': ('Все направления МинЖКХ', {})}
for _topic, _filter, _ident in (
    ('water', 'direction', '6'), ('sewerage', 'synt_group', '39'),
    ('water_quality', 'synt_group', '64'), ('hot_water', 'subtopic', '92'),
    ('heating', 'direction', '15'), ('capital_repair', 'subtopic', '548'),
    ('garbage', 'subtopic', '80'), ('rusty_water', 'subtopic', '419'),
    ('no_water', 'subtopic', '308'), ('pressure', 'subtopic', '233'),
    ('wells', 'subtopic', '174'), ('sewage_discharge', 'subtopic', '426'),
    ('boilers', 'subtopic', '183'), ('heat_networks', 'subtopic', '515'),
    ('network_modernization', 'subtopic', '215'), ('ecology', 'synt_group', '173'),
    ('yards', 'direction', '7'), ('public_areas', 'direction', '12'),
):
    TOPICS[_topic] = (_FILTER_LABELS[_filter][_ident], {_filter: _ident})
_CITY = re.compile(r'[А-ЯЁа-яёA-Za-z][А-ЯЁа-яёA-Za-z0-9 .()—–-]{0,159}')
_OTHER_TOPIC = 'Прочие / неопределённые'
_BREAKDOWN_LABELS = {
    dimension: {label.casefold().replace('ё', 'е'): label for ident, label in values if ident != '-1'}
    for dimension, values in (('theme', dashboard_filters.THEMES), ('subtopic', dashboard_filters.SUBTOPICS))
}


def _controlled_label(dimension, value):
    """Project to source-controlled labels; never return unknown source prose."""
    if not isinstance(value, str):
        return _OTHER_TOPIC
    return _BREAKDOWN_LABELS[dimension].get(value.strip().casefold().replace('ё', 'е'), _OTHER_TOPIC)


def _distribution(counts):
    return [{'label': label, 'count': count}
            for label, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))]


def _access(user):
    if not user or not user.get('is_active', True) or not check_module_access(user, 'mingkh'):
        raise HTTPException(403, 'Нет доступа к жалобам МинЖКХ.')


def validate(user, params):
    _access(user)
    if not isinstance(params, dict) or set(params) - {'start', 'end', 'municipality', 'topic'}:
        raise ValueError('Для запроса МинЖКХ допустимы период, ОМСУ и поддерживаемая тема.')
    if not {'start', 'end'} <= set(params):
        raise ValueError('Укажите начало и конец периода запроса МинЖКХ.')
    if any(not isinstance(params[key], str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', params[key])
           for key in ('start', 'end')):
        raise ValueError('Даты должны быть в формате ГГГГ-ММ-ДД.')
    first, last = (date.fromisoformat(params[key]) for key in ('start', 'end'))
    today = report_periods.today()
    if first > last or (last-first).days > 365 or first > today or (today-first).days > 1095 or last.year > today.year:
        raise ValueError('Как в дашборде МинЖКХ, доступен период до 366 дней, начало не старше трёх лет и не в будущем.')
    topic = params.get('topic', 'all')
    if not isinstance(topic, str) or topic not in TOPICS:
        raise ValueError('Тема не поддерживается этим запросом. Выберите тему из каталога инструмента МинЖКХ.')
    city = params.get('municipality', '')
    if not isinstance(city, str) or (city.strip() and not _CITY.fullmatch(city.strip())):
        raise ValueError('Укажите одно название ОМСУ. Его наличие будет проверено в справочнике источника.')
    return {'start': first.isoformat(), 'end': last.isoformat(), 'municipality': city.strip(), 'topic': topic}


def _result(params):
    return {'tool_id': ID, 'module': 'mingkh', 'source': {'id': ID, 'title': TITLE, 'url': URL},
            'status': 'missing', 'coverage': 'missing', 'requested_period': {'from': params['start'], 'to': params['end']},
            'effective_period': None, 'available_period': None,
            'filters': {'municipality': params['municipality'], 'topic': params['topic'], 'topic_label': TOPICS[params['topic']][0]},
            'coverage_basis': 'completed_live_queries', 'completeness_verified': False,
            'metrics': [], 'topic_breakdown': [], 'theme_breakdown': [],
            'warnings': [], 'as_of': None, 'query_count': 0}


def _source_revision(portal):
    value, _ = portal.last_updated()
    # Only a validated date/time, never source prose, becomes evidence or a key.
    day = dashboard.normalize_date(value)
    if not day or not isinstance(value, str) or len(value) > 40:
        return None, None
    return day, value.strip()


def _fail(result, code, message):
    result.update(status='missing', coverage='missing', metrics=[], topic_breakdown=[], theme_breakdown=[], error_code=code, message=message)
    return result


def _municipality(portal, wanted):
    _, raw = portal.query('q_omsu', portal.periods(6, 0, 13, 7), timeout=30)
    if not isinstance(raw, list) or len(raw) > 1024:
        raise PortalError('Некорректный справочник ОМСУ источника.')
    rows = _omsu_rows(raw)
    choices = [(ident, name) for ident, name in rows if _CITY.fullmatch(name) and cities.is_municipality(name)]
    if not wanted:
        return '-1', '', [name for _, name in choices][:30]
    matches = [(ident, name) for ident, name in choices if cities.key(name) == cities.key(wanted)]
    if len(matches) != 1:
        return None, '', [name for _, name in choices][:30]
    return matches[0][0], matches[0][1], []


def _count(columns, rows, city, start, end, topic):
    required = [dashboard.COLS[dim] for dim in ('omsu', 'direction', 'theme', 'subtopic', 'date')]
    if any(key not in columns for key in required) or len(columns) != len(set(columns)):
        raise PortalError('Структура детализации МинЖКХ изменилась.')
    positions = {key: columns.index(dashboard.COLS[key]) for key in ('omsu', 'direction', 'theme', 'subtopic', 'date')}
    # Reuse exactly the dashboard's exclusions. Parameter filtering is applied
    # at the portal; the city/date consistency check detects ignored filters.
    count, excluded, undated = 0, 0, 0
    topics, themes = Counter(), Counter()
    filter_name = next(iter(TOPICS[topic][1]), None)
    dimension = {'direction': 'direction', 'synt_group': 'theme', 'subtopic': 'subtopic'}.get(filter_name)
    expected = TOPICS[topic][0].casefold().replace('ё', 'е')
    for row in rows:
        if not isinstance(row, list) or len(row) != len(columns):
            raise PortalError('Неполная строка детализации МинЖКХ.')
        if city and cities.key(row[positions['omsu']]) != cities.key(city):
            raise PortalError('Источник не подтвердил фильтр ОМСУ.')
        if dimension and str(row[positions[dimension]]).strip().casefold().replace('ё', 'е') != expected:
            raise PortalError('Источник не подтвердил тематический фильтр.')
        if any(row[positions[key]] in names for key, names in dashboard.EXCLUDE.items()):
            excluded += 1
            continue
        received = dashboard.normalize_date(row[positions['date']])
        if not received:
            undated += 1
        elif not start <= received <= end:
            raise PortalError('Источник вернул строки за пределами запрошенного периода.')
        count += 1
        topics[_controlled_label('subtopic', row[positions['subtopic']])] += 1
        themes[_controlled_label('theme', row[positions['theme']])] += 1
    return count, excluded, undated, _distribution(topics), _distribution(themes)


def execute(user, params, on_progress=None):
    params = validate(user, params)
    result = _result(params)
    stage = on_progress or (lambda label: None)
    try:
        stage('Подключаемся к действующему источнику МинЖКХ')
        portal = Pentaho()
        source_day, before = _source_revision(portal)
        ident, city, choices = _municipality(portal, params['municipality'])
        if ident is None:
            result['available_municipalities'] = choices
            return _fail(result, 'MINGKH_UNKNOWN_MUNICIPALITY', 'ОМСУ не найден однозначно в действующем справочнике МинЖКХ. Уточните его название.')
        result['filters']['municipality'] = city
        today = report_periods.today()
        end = min(params['end'], today.isoformat(), source_day or today.isoformat())
        result.update(as_of=before, available_period={'from': None, 'to': source_day})
        if end < params['start']:
            return _fail(result, 'MINGKH_PERIOD_UNAVAILABLE', 'Источник ещё не обновлён до начала запрошенного периода. Архивные значения не подставлены.')
        effective = {'from': params['start'], 'to': end}
        result['effective_period'] = effective
        cs, ce = ((today-date.fromisoformat(effective[key])).days for key in ('from', 'to'))
        query_params = portal.periods(cs, ce, cs, ce, omsu=ident, **TOPICS[params['topic']][1])
        stage('Получаем детализацию МинЖКХ за ' + params['start'] + ' — ' + end)
        raw, errors = dashboard.fetch_periods(portal, query_params, {'curr': (cs, ce)})
        result['query_count'] = len(list(dashboard.chunks(cs, ce)))
        if 'curr' not in raw or errors:
            return _fail(result, 'MINGKH_QUERY_FAILED', 'Источник не вернул все части периода. Неполная выгрузка не показана как итог.')
        stage('Проверяем даты, ОМСУ и полноту полученных частей')
        count, excluded, undated, topics, themes = _count(*raw['curr'], city, effective['from'], effective['to'], params['topic'])
        _, after = _source_revision(portal)
        stable = bool(before and after and before == after)
        result['snapshot_id'] = hashlib.sha256(('mingkh-live-v1\0' + before).encode()).hexdigest() if stable else uuid4().hex
        if not stable:
            result['warnings'].append('Версия источника не подтверждена или изменилась во время запроса. Сравнение с другим запуском нельзя считать окончательным.')
        if effective != result['requested_period']:
            result['warnings'].append('Запрошенный период доступен частично. Будущие даты и даты после обновления источника не включены.')
        if end == today.isoformat():
            result['warnings'].append('Последний день ещё не завершён. Данные за него могут поступать позднее.')
        if undated:
            result['warnings'].append('Не подтверждена дата у ' + str(undated) + ' обращений. Они учтены по ответу источника; период не считается полностью проверенным.')
        complete = stable and source_day is not None and effective == result['requested_period'] and end < today.isoformat() and not undated
        result.update(status='ok' if complete else 'partial', coverage='full' if complete else 'partial',
                      completeness_verified=complete, excluded=excluded, unknown_dates=undated,
                      topic_breakdown=topics, theme_breakdown=themes)
        result['metrics'] = [{'id': 'complaints', 'label': 'Обращения МинЖКХ из действующего источника',
                              'value': count, 'unit': 'жалоб', 'definition': ID + ':first_work_date_exclusions_v1:' + params['topic']}]
        result['warnings'].append('Период применяется к дате первого взятия в работу, как в дашборде МинЖКХ. Использованы его штатные исключения.')
        result['warnings'].append('Распределение показывает тематику обращений, а не установленные причины их возникновения. Нераспознанные темы объединены в «Прочие / неопределённые».')
        return result
    except HTTPException as error:
        if error.status_code in (401, 403):
            raise
        return _fail(result, 'MINGKH_CREDENTIALS_UNAVAILABLE', 'Не удалось прочитать доступ к МинЖКХ. Администратору нужно проверить его в разделе «Логины и пароли».')
    except PortalError:
        return _fail(result, 'MINGKH_SOURCE_UNAVAILABLE', 'Запрос к МинЖКХ не завершён. Проверьте доступ к источнику и сертификаты на сервере. Сохранённые данные вместо ответа источника не использованы.')
    except (OSError, ValueError, TypeError, KeyError, IndexError):
        return _fail(result, 'MINGKH_INVALID_RESPONSE', 'Источник МинЖКХ вернул неподтверждённую структуру или фильтры. Итог не опубликован.')
