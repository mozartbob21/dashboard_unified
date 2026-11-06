"""Read the live collective-complaint source without importing workflow state."""
from datetime import date, datetime, timezone
import hashlib
import json
import re
from uuid import uuid4

from fastapi import HTTPException
from core.roles import check_module_access
from services import report_municipalities as cities
from services.auth.integrations import credentials
from services.aichat import report_periods
from services.aichat.collective_source import _signatures
from services.collective.client import CollectivePortal, CURATOR, HEADERS, PortalError
from services.mingkh.dashboard import normalize_date

ID = 'query:collective'
TITLE = 'Коллективные жалобы — запрос из действующего источника'
URL = '/mingkh/collective'
_LABEL = re.compile(r'[А-ЯЁа-яёA-Za-z][А-ЯЁа-яёA-Za-z0-9 .(),—–/\-]{0,159}')
MAX_ROWS = 300000


def _access(user):
    if not user or not user.get('is_active', True) or not check_module_access(user, 'collective'):
        raise HTTPException(403, 'Нет доступа к коллективным жалобам.')


def validate(user, params):
    _access(user)
    if not isinstance(params, dict) or set(params) - {'start', 'end', 'municipality', 'status'}:
        raise ValueError('Для коллективных жалоб допустимы период, ОМСУ и точный статус источника.')
    if not {'start', 'end'} <= set(params):
        raise ValueError('Укажите начало и конец периода коллективных жалоб.')
    if any(not isinstance(params[key], str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', params[key]) for key in ('start', 'end')):
        raise ValueError('Даты должны быть в формате ГГГГ-ММ-ДД.')
    first, last = (date.fromisoformat(params[key]) for key in ('start', 'end'))
    today = report_periods.today()
    if first > last or (last-first).days > 365 or first > today or last.year > today.year:
        raise ValueError('Для одного запроса выберите до 366 дней, начало не в будущем и окончание не позднее текущего года.')
    result = {'start': first.isoformat(), 'end': last.isoformat()}
    for key in ('municipality', 'status'):
        value = params.get(key, '')
        if not isinstance(value, str) or (value.strip() and not _LABEL.fullmatch(value.strip())):
            raise ValueError('Укажите название ОМСУ и статус без специальных команд или разметки.')
        result[key] = value.strip()
    return result


def _result(params):
    return {'tool_id': ID, 'module': 'collective', 'source': {'id': ID, 'title': TITLE, 'url': URL},
            'status': 'missing', 'coverage': 'missing',
            'requested_period': {'from': params['start'], 'to': params['end']},
            'effective_period': None, 'available_period': None,
            'filters': {'municipality': params['municipality'], 'status': params['status']},
            'coverage_basis': 'completed_live_query', 'completeness_verified': False,
            'comparison_consistency': 'independent_live_reads', 'snapshot_id': uuid4().hex,
            'as_of': None, 'retrieved_at': None, 'metrics': [], 'warnings': []}


def _fail(result, code, message):
    result.update(status='missing', coverage='missing', metrics=[], error_code=code, message=message)
    return result


def _identity(row):
    # Matching rules mirror the workspace's primary/secondary/fallback order,
    # without opening its identity tables or modifying saved manual annotations.
    norm = lambda value: ' '.join(str(value or '').split()).casefold()
    if any(len(str(row[index] or '')) > 300 for index in (0, 5, 6)):
        raise ValueError('Invalid source identity')
    if norm(row[5]):
        return 'primary:' + norm(row[5])
    if norm(row[6]):
        return 'secondary:' + json.dumps([norm(row[0]), norm(row[6])], ensure_ascii=False)
    stable = json.dumps([norm(row[index]) for index in (0, 1, 2, 8, 9)], ensure_ascii=False)
    return 'fallback:' + hashlib.sha256(stable.encode()).hexdigest()


def _selection(result, rows, params, effective):
    if not isinstance(rows, list) or len(rows) > MAX_ROWS:
        raise ValueError('Invalid row count')
    seen, unique = {}, []
    names, statuses = {}, {}
    for row in rows:
        if not isinstance(row, list) or len(row) != len(HEADERS):
            raise ValueError('Invalid row width')
        if any(isinstance(value, (dict, list, tuple)) for value in row):
            raise ValueError('Invalid row value')
        # A changed/ignored curator filter must not silently widen this module.
        if CURATOR not in str(row[3] or '').casefold():
            raise ValueError('Unconfirmed curator')
        received = normalize_date(row[1])
        if received and not effective['from'] <= received <= effective['to']:
            raise ValueError('Response outside requested dates')
        ident = _identity(row)
        fields = (received, cities.key(row[2]), str(row[14] or '').strip(), _signatures(row[10]))
        if ident in seen:
            if seen[ident] != fields:
                raise ValueError('Contradictory duplicate source identity')
            continue
        seen[ident] = fields
        unique.append((row, received))
        name, status = str(row[2] or '').strip(), str(row[14] or '').strip()
        if _LABEL.fullmatch(name) and cities.is_municipality(name):
            names.setdefault(cities.key(name), cities.display(name))
        if _LABEL.fullmatch(status):
            statuses.setdefault(status.casefold(), status)
    if params['municipality'] and cities.key(params['municipality']) not in names:
        result['available_municipalities'] = sorted(names.values())[:30]
        return _fail(result, 'COLLECTIVE_MUNICIPALITY_UNAVAILABLE', 'В текущем ответе источника нет выбранного ОМСУ. Это не доказывает нулевое число обращений. Уточните ОМСУ или период.')
    if params['status'] and params['status'].casefold() not in statuses:
        result['available_statuses'] = sorted(statuses.values())[:30]
        return _fail(result, 'COLLECTIVE_STATUS_UNAVAILABLE', 'Точный статус отсутствует в ответе источника за этот период. Выберите один из доступных статусов.')
    city_key = cities.key(params['municipality']) if params['municipality'] else ''
    status_key = params['status'].casefold()
    result['filters'] = {'municipality': names[city_key] if city_key else '',
                         'status': statuses[status_key] if status_key else ''}
    selected = [(row, received) for row, received in unique
                if (not city_key or cities.key(row[2]) == city_key)
                and (not status_key or str(row[14] or '').strip().casefold() == status_key)]
    unknown_dates = sum(not received for row, received in selected)
    unknown_places = sum(cities.key(row[2]) not in names for row, _ in unique) if city_key else 0
    complete = effective == result['requested_period'] and effective['to'] < report_periods.today().isoformat() and not unknown_dates and not unknown_places
    result.update(status='ok' if complete else 'partial', coverage='full' if complete else 'partial',
                  completeness_verified=complete, unknown_dates=unknown_dates, duplicated_rows=len(rows)-len(unique))
    definitions = ID + ':received_date_curator_v1'
    result['metrics'] = [{'id': 'complaints', 'label': 'Коллективные обращения в ответе источника',
                          'value': len(selected), 'unit': 'обращений', 'definition': definitions + ':complaints'}]
    signatures = [_signatures(row[10]) for row, _ in selected]
    missing = sum(value is None for value in signatures)
    result['signatures_coverage'] = {'known': len(signatures)-missing, 'unknown': missing}
    if missing:
        result['warnings'].append('Для ' + str(missing) + ' обращений число подписей не подтверждено. Общая сумма подписей не рассчитана.')
    else:
        total = sum(signatures)
        if total > 2**53-1:
            raise ValueError('Signature total outside safe range')
        result['metrics'].append({'id': 'signatures', 'label': 'Подписи в подтверждённых обращениях',
                                  'value': total, 'unit': 'подписей', 'definition': definitions + ':signatures'})
    if unknown_dates:
        result['warnings'].append('Не подтверждена дата у ' + str(unknown_dates) + ' обращений; полнота периода не установлена.')
    if unknown_places:
        result['warnings'].append('В ответе есть обращения без подтверждённого ОМСУ. Муниципальный итог может быть неполным.')
    return result


def execute(user, params, on_progress=None):
    params = validate(user, params)
    result = _result(params)
    stage = on_progress or (lambda label: None)
    try:
        access = credentials('mingkh')
        if not access or not all(isinstance(access.get(key), str) and access[key].strip() for key in ('username', 'password')):
            return _fail(result, 'COLLECTIVE_CREDENTIALS_MISSING', 'Администратору нужно сохранить доступ к МинЖКХ в разделе «Логины и пароли».')
        today = report_periods.today()
        effective = {'from': params['start'], 'to': min(params['end'], today.isoformat())}
        result['effective_period'] = effective
        stage('Запрашиваем коллективные жалобы из действующего источника')
        _, rows = CollectivePortal(**access).fetch(date.fromisoformat(effective['from']), date.fromisoformat(effective['to']), today=today)
        stage('Проверяем период и считаем обращения и подписи')
        result['retrieved_at'] = datetime.now(timezone.utc).isoformat(timespec='seconds')
        result['warnings'].append('Источник не сообщает дату актуальности и общую версию. Это отдельный ответ на запрос; время получения не заменяет дату обновления источника.')
        result['warnings'].append('Период применяется к дате поступления. ОМСУ и точный статус выбираются из полученной выгрузки, как на странице коллективных жалоб. Ручные доклады и назначения не меняются.')
        if effective != result['requested_period'] or effective['to'] == report_periods.today().isoformat():
            result['warnings'].append('Период неполный — текущий день ещё не завершён; будущие даты не запрошены.')
        return _selection(result, rows, params, effective)
    except HTTPException as error:
        if error.status_code in (401, 403):
            raise
        return _fail(result, 'COLLECTIVE_CREDENTIALS_UNAVAILABLE', 'Не удалось прочитать сохранённый доступ к МинЖКХ. Обратитесь к администратору.')
    except PortalError:
        return _fail(result, 'COLLECTIVE_SOURCE_UNAVAILABLE', 'Источник коллективных жалоб не ответил корректно. Проверьте доступ и защищённое соединение на сервере. Архивные значения не подставлены.')
    except (ValueError, TypeError, KeyError, IndexError, OSError):
        return _fail(result, 'COLLECTIVE_INVALID_RESPONSE', 'Ответ источника не подтвердил структуру, даты или уникальность обращений. Итог не опубликован.')
