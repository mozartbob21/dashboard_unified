"""Fixed Pentaho CDA contract recovered statically from the supplied dashboard."""
import json
import os
import ssl
from datetime import date
from urllib.parse import urlencode

import httpx
from services.mingkh.client import BASE, PortalError

CDA_PATH = '/public/jalobi/jalobi.cda'
QUERIES = frozenset({'q_vedom_kurat', 'q_download'})
CURATOR = 'министерство жилищно-коммунального хозяйства московской области'
HEADERS = [
    'Источник', 'Дата поступления обращения', 'ОМСУ', 'Ведомство куратора',
    'Ведомство исполнителя', 'Номер обращения ЕЦУР, МСЭД',
    'Номер обращения исходной системы', 'Источник обращения ЕЦУР',
    'Аннотация/краткое содержание', 'Факт', 'Подписи (фактическое число подписей)',
    'Ответ', 'Почта заявителя', 'Повтор', 'Статус', 'Комментарии',
]
MAX_RESPONSE = 40 * 1024 * 1024


def period(start, end, today=None):
    today = today or date.today()
    try:
        start, end = date.fromisoformat(start), date.fromisoformat(end)
    except (ValueError, TypeError):
        raise ValueError('Укажите начало и конец периода.') from None
    if start > end or end > today:
        raise ValueError('Начало периода должно быть не позже конца; будущие даты недоступны.')
    return start, end


def query_params(start, end, kurators, today=None):
    today = today or date.today()
    return {
        'all_per': '1' if start == end == today else '0',
        'curr_period_start': f'{(today - start).days} day',
        'curr_period_end': f'{(today - end).days} day',
        'omsu': ['-1'], 'vedom_ispoln': ['-1'], 'vedom_kurat': [k[0] for k in kurators],
        'status': ['-1'], 'class': ['-1'],
    }


class CollectivePortal:
    def __init__(self, username, password):
        self.username, self.password = username, password

    def query(self, query, params=None):
        if query not in QUERIES:
            raise PortalError('Неизвестный запрос коллективных обращений.')
        data = [('path', CDA_PATH), ('dataAccessId', query)]
        for key, value in (params or {}).items():
            for item in value if isinstance(value, list) else [value]:
                data.append(('param' + key, str(item)))
        try:
            context = ssl.create_default_context(cafile=os.getenv('MINGKH_CA_BUNDLE') or None)
            with httpx.Client(auth=(self.username, self.password), verify=context, trust_env=False,
                              follow_redirects=False, timeout=httpx.Timeout(300, connect=15)) as client:
                with client.stream('POST', BASE, content=urlencode(data).encode(),
                                   headers={'Content-Type': 'application/x-www-form-urlencoded'}) as response:
                    if response.status_code in (401, 403) or response.is_redirect:
                        raise PortalError('Портал не принял логин и пароль. Администратору нужно проверить доступ МИНЖКХ.')
                    if response.status_code != 200:
                        raise PortalError(f'Портал временно недоступен (HTTP {response.status_code}).',
                                          retryable=response.status_code == 429 or response.status_code >= 500)
                    payload = bytearray()
                    for part in response.iter_bytes():
                        payload.extend(part)
                        if len(payload) > MAX_RESPONSE:
                            raise PortalError('Выгрузка слишком большая. Выберите меньший период.')
            rows = json.loads(payload)['resultset']
            if not isinstance(rows, list) or len(rows) > 300000:
                raise ValueError()
            width = len(HEADERS) if query == 'q_download' else 2
            if any(not isinstance(row, list) or
                   (len(row) != width if query == 'q_download' else len(row) < width) or
                   any(isinstance(cell, (list, dict)) for cell in row) for row in rows):
                raise ValueError()
            return rows
        except PortalError:
            raise
        except (httpx.TimeoutException, httpx.NetworkError):
            raise PortalError('Соединение с порталом прервалось. Повторите загрузку.', retryable=True) from None
        except (httpx.TransportError, OSError):
            raise PortalError('Нет защищённого соединения с cur.bi.mosreg.ru. Проверьте доступ и сертификаты.') from None
        except (ValueError, KeyError, TypeError):
            raise PortalError('Портал вернул неожиданный формат коллективных обращений.') from None

    def fetch(self, start, end):
        kurators = []
        for row in self.query('q_vedom_kurat'):
            if CURATOR not in str(row[1] or '').lower():
                continue
            cid = str(row[0]).strip() if row[0] is not None else ''
            if not cid or cid == '-1':
                raise PortalError('Портал вернул некорректный идентификатор куратора МИНЖКХ. Проверьте изменения справочника.')
            kurators.append((cid, str(row[1])))
        # Empty or invalid curator filters could expand the export to every department.
        # Fail the refresh so the browser retains its previous, correctly labelled snapshot.
        if not kurators:
            raise PortalError('Не удалось найти куратора МИНЖКХ в справочнике портала. Проверьте доступ или изменения справочника.')
        rows = self.query('q_download', query_params(start, end, kurators))
        return kurators, rows
