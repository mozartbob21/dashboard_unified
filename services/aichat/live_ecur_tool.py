"""A private, bounded live ECUR session; shared UI state is never read or changed."""
from collections import Counter
from datetime import datetime, timezone
import hashlib
import re
from uuid import uuid4

import requests
from fastapi import HTTPException
from core.roles import check_module_access
from services.auth.integrations import credentials
from services import report_municipalities as cities
from services.aichat import report_periods
from services.ecur import client
from services.mingkh.dashboard import normalize_date

ID = 'query:ecur'
_CITY = re.compile(r'[А-ЯЁа-яёA-Za-z][А-ЯЁа-яёA-Za-z0-9 .()—–-]{0,159}')
_STATUS = re.compile(r'[А-ЯЁа-яёA-Za-z][А-ЯЁа-яёA-Za-z0-9 .(),—–/\-]{0,159}')


def validate(user, params):
    if not user or not user.get('is_active', True) or not check_module_access(user, 'ecur'):
        raise HTTPException(403, 'Нет доступа к жалобам ЕЦУР.')
    if not isinstance(params, dict) or set(params) - {'municipality'}:
        raise ValueError('Этот запрос ЕЦУР получает текущие жалобы. Доступен только фильтр ОМСУ, без исторических дат.')
    city = params.get('municipality', '')
    if not isinstance(city, str) or len(city) > 160 or (city.strip() and not _CITY.fullmatch(city.strip())):
        raise ValueError('Укажите одно название ОМСУ без команд или разметки.')
    return {'municipality': city.strip()}


def _fail(result, code, message):
    result.update(status='missing', coverage='missing', metrics=[], error_code=code, message=message)
    return result


def execute(user, params, on_progress=None):
    params = validate(user, params)
    today = report_periods.today()
    scope_key = hashlib.sha256((client.CURATOR + '\0' + client.STATUSES + '\0' + client.DISTRICTS).encode()).hexdigest()[:16]
    result = {'tool_id': ID, 'module': 'ecur', 'source': {'id': ID, 'title': 'ЕЦУР — текущие жалобы', 'url': '/ecur'},
              'status': 'missing', 'coverage': 'missing', 'requested_period': None, 'effective_period': None,
              'available_period': None, 'as_of': None, 'retrieved_at': None,
              'filters': {**params, 'source_scope': scope_key, 'deadline_from': today.isoformat()},
              'snapshot_id': uuid4().hex, 'comparison_consistency': 'independent_live_reads',
              'coverage_basis': 'completed_live_current_report', 'completeness_verified': False,
              'metrics': [], 'warnings': [
                  'Использован текущий оперативный отчёт ЕЦУР со штатными куратором и статусами блока, срок исполнения — с сегодняшнего дня.',
                  'Это текущий срез активных жалоб, а не все обращения за период. Источник не сообщает общую версию и отчётную дату.']}
    if client.DISTRICTS:
        result['warnings'].append('На сервере задан дополнительный фильтр районов источника; результат ограничен этим набором.')
    stage = on_progress or (lambda label: None)
    session = None
    try:
        access = credentials('edds')
        if not access or not all(isinstance(access.get(key), str) and access[key].strip() for key in ('username', 'password')):
            return _fail(result, 'ECUR_CREDENTIALS_MISSING', 'Администратору нужно сохранить доступ «Добродел» в разделе «Логины и пароли». Чужая открытая сессия ЕЦУР не используется.')
        stage('Входим в источник ЕЦУР с сохранённым доступом Добродела')
        session = requests.Session()
        session.trust_env = False
        session.verify = True
        active, error = client.dobrodel_login(access['username'], access['password'], session=session, strict=True)
        if error or active is not session:
            return _fail(result, 'ECUR_LOGIN_FAILED', 'Вход в ЕЦУР не подтверждён. Проверьте сохранённый доступ Добродела, сеть и сертификаты сервера.')
        stage('Получаем страницы текущих жалоб ЕЦУР')
        rows, error = client.fetch_report(session, max_pages=21, max_rows=10000, page_size=500, today=today, strict=True)
        if error or not rows or rows[0] != client.HEADER:
            return _fail(result, 'ECUR_FETCH_FAILED', 'Выгрузка ЕЦУР не подтверждена полностью. Возможны ошибка источника или превышение предела 10 000 записей; частичный ответ не показан как полный.')
        seen, names, records = set(), {}, []
        for row in rows[1:]:
            if not isinstance(row, list) or len(row) != 11:
                raise ValueError('Invalid row')
            ident = str(row[0] or '').strip()
            if not re.fullmatch(r'\d{1,20}', ident) or ident in seen:
                raise ValueError('Missing or duplicate source identity')
            seen.add(ident)
            city = str(row[2] or '').strip()
            if _CITY.fullmatch(city) and cities.is_municipality(city):
                names.setdefault(cities.key(city), cities.display(city))
            deadline = normalize_date(row[9])
            if deadline and deadline < today.isoformat():
                raise ValueError('Ignored deadline filter')
            records.append((row, deadline))
        city = cities.key(params['municipality']) if params['municipality'] else ''
        if city and city not in names:
            result['available_municipalities'] = sorted(names.values())[:30]
            return _fail(result, 'ECUR_MUNICIPALITY_UNAVAILABLE', 'В текущем ответе ЕЦУР нет выбранного ОМСУ. Отсутствие строки не подтверждает отсутствие жалоб.')
        selected = [(row, deadline) for row, deadline in records if not city or cities.key(row[2]) == city]
        unknown_dates = sum(not deadline for row, deadline in selected)
        unknown_places = sum(cities.key(row[2]) not in names for row, _ in records) if city else 0
        counts = Counter(str(row[10] or '').strip() for row, _ in selected)
        status_counts = {status: count for status, count in counts.items() if _STATUS.fullmatch(status)}
        unknown_statuses = sum(count for status, count in counts.items() if not _STATUS.fullmatch(status))
        stage('Проверяем ОМСУ, сроки и статусы полученных жалоб')
        result['filters']['municipality'] = names[city] if city else ''
        complete = not unknown_dates and not unknown_places
        result.update(status='ok' if complete else 'partial', coverage='full' if complete else 'partial',
                      retrieved_at=datetime.now(timezone.utc).isoformat(timespec='seconds'),
                      completeness_verified=complete,
                      metrics=[{'id': 'complaints', 'label': 'Текущие жалобы в ответе ЕЦУР', 'value': len(selected),
                                'unit': 'жалоб', 'definition': ID + ':current_deadline_v1:' + scope_key}],
                      status_counts=status_counts)
        if unknown_dates:
            result['warnings'].append('Для ' + str(unknown_dates) + ' жалоб не подтверждён срок. Полнота фильтра по сроку не установлена.')
        if unknown_places:
            result['warnings'].append('В ответе есть жалобы без подтверждённого района. Муниципальный итог может быть неполным.')
        if unknown_statuses:
            result['warnings'].append('Для ' + str(unknown_statuses) + ' жалоб статус не распознан; они учтены в общем количестве.')
        return result
    except HTTPException as error:
        if error.status_code in (401, 403):
            raise
        return _fail(result, 'ECUR_CREDENTIALS_UNAVAILABLE', 'Не удалось прочитать сохранённый доступ Добродела. Обратитесь к администратору.')
    except (ValueError, TypeError, KeyError, IndexError, OSError, requests.RequestException):
        return _fail(result, 'ECUR_INVALID_RESPONSE', 'Источник ЕЦУР не подтвердил данные. Проверьте подключение и структуру отчёта; предыдущие значения не подставлены.')
    finally:
        if session is not None:
            session.close()
