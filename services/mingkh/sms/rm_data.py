# -*- coding: utf-8 -*-
"""Блок 4 СМС: быстрые и системные решения по данным Redmine.

Мероприятие — дата в одном из полей задачи «Адреса по жалобам», попавшая в
период. Сверено 16.10.2026 с СМС за 01–07.10 по журналам RM на момент отправки:
быстрые 6 адресов / 6 мероприятий, системные 5 / 6 — совпало точно.
Данные всегда живые: что внесено в RM на момент запуска.
"""
import concurrent.futures
import json
import os
import re
import ssl
import httpx
from .context import date_range

TRACKER = 25
RM = 'https://mgkh.rm.mosreg.ru'
MAX_RESPONSE_BYTES = 10 * 1024 * 1024
MAX_ISSUES = 50000


class RedmineError(RuntimeError):
    pass

# поле RM -> (группа, полное название, короткое для СМС)
FIELDS = {
    43: ("quick", "Промывка МКД ХВС и ГВС", "промывка МКД"),
    53: ("quick", "Промывка уличного трубопровода через гидранты", "промывка уличного трубопровода"),
    55: ("quick", "Промывка уличного трубопровода от ЦТП (котельной) до МКД", "промывка от ЦТП до МКД"),
    228: ("quick", "Промывка сетей гидродинамическим методом", "гидродинамическая промывка"),
    229: ("system", "Установка датчиков мутности", "датчики мутности"),
    45: ("system", "Установка частотных преобразователей на ВЗУ", "частотные преобразователи"),
    47: ("system", "Установка спускных кранов на лежаках в МКД", "спускные краны"),
    49: ("system", "Установка грязевого фильтра на вводе в МКД", "грязевой фильтр"),
    # 51 «Регистрация проведения промывки в ЕДС» — подтверждение, а не мероприятие:
    # в СМС не входит (в СМС за 01–07.10 две такие даты не посчитаны).
}
TYPES = {"334": "Резонансные", "335": "Системный адрес", "336": "Системные отключения"}


def configured():
    return bool((os.getenv('MGKH_API_KEY') or '').strip())


def _get(path):
    if not re.fullmatch(r'/issues\.json\?status_id=\*&tracker_id=25&limit=100&offset=[0-9]{1,6}', path):
        raise ValueError('Недопустимый запрос к РМ МинЖКХ.')
    key = (os.getenv('MGKH_API_KEY') or '').strip()
    if not key:
        raise RedmineError('На сервере не настроен ключ доступа к РМ МинЖКХ (MGKH_API_KEY).')
    try:
        context = ssl.create_default_context(cafile=os.getenv('MGKH_CA_BUNDLE') or None)
        with httpx.Client(headers={'X-Redmine-API-Key': key}, verify=context, trust_env=False,
                          follow_redirects=False, timeout=httpx.Timeout(120, connect=15)) as client:
            with client.stream('GET', RM + path) as response:
                if response.status_code in (401, 403) or response.is_redirect:
                    raise RedmineError('РМ МинЖКХ не принял ключ доступа. Администратору нужно проверить настройки сервера.')
                if response.status_code != 200:
                    raise RedmineError('РМ МинЖКХ временно недоступен (HTTP %d). Повторите сбор позже.' % response.status_code)
                content = bytearray()
                for chunk in response.iter_bytes():
                    content.extend(chunk)
                    if len(content) > MAX_RESPONSE_BYTES:
                        raise RedmineError('РМ МинЖКХ вернул слишком большой ответ.')
        result = json.loads(content)
        if not isinstance(result, dict) or not isinstance(result.get('issues'), list) or len(result['issues']) > 100:
            raise ValueError
        if any(not isinstance(issue, dict) for issue in result['issues']):
            raise ValueError
        return result
    except RedmineError:
        raise
    except httpx.TimeoutException:
        raise RedmineError('РМ МинЖКХ не ответил вовремя. Повторите сбор позже.') from None
    except (httpx.TransportError, ssl.SSLError, OSError):
        raise RedmineError('Нет защищённого соединения с РМ МинЖКХ. Проверьте доступ с сервера и сертификаты.') from None
    except (ValueError, TypeError, KeyError):
        raise RedmineError('РМ МинЖКХ вернул неожиданный формат данных.') from None


def issues():
    """All issue statuses, with bounded pagination and no partial success."""
    query = '/issues.json?status_id=*&tracker_id=%d&limit=100&offset=%d'
    first = _get(query % (TRACKER, 0))
    count = first.get('total_count')
    if type(count) is not int or not 0 <= count <= MAX_ISSUES:
        raise RedmineError('РМ МинЖКХ вернул неверное или слишком большое количество задач.')
    result = list(first['issues'])
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        for page in executor.map(lambda offset: _get(query % (TRACKER, offset)), range(100, count, 100)):
            result.extend(page['issues'])
    identifiers = [issue.get('id') for issue in result]
    if (len(result) != count or any(type(ident) is not int or ident <= 0 for ident in identifiers)
            or len(set(identifiers)) != count):
        raise RedmineError('Список задач РМ МинЖКХ изменился во время загрузки. Повторите сбор.')
    return result


def measures(date_from, date_to):
    """Мероприятия с датой в [date_from, date_to] (ISO-строки)."""
    date_from, date_to = (value.isoformat() for value in date_range(date_from, date_to, max_days=366))
    out = []
    for i in issues():
        fields = i.get('custom_fields') or []
        if not isinstance(fields, list) or any(not isinstance(field, dict) for field in fields):
            raise RedmineError('РМ МинЖКХ вернул некорректные поля задач.')
        cf = {c.get('id'): c.get('value') for c in fields}
        for label in ('assigned_to', 'status'):
            if i.get(label) is not None and not isinstance(i[label], dict):
                raise RedmineError('РМ МинЖКХ вернул некорректные данные задачи.')
        for fid, (group, title, short) in FIELDS.items():
            d = str(cf.get(fid) or "")[:10]
            if date_from <= d <= date_to:
                out.append({
                    "id": i["id"], "url": "%s/issues/%d" % (RM, i["id"]),
                    "subject": i.get("subject") or "",
                    "omsu": (i.get("assigned_to") or {}).get("name") or "",
                    "status": (i.get('status') or {}).get('name') or '', 
                    "type": TYPES.get(str(cf.get(221)), str(cf.get(221) or "")),
                    "group": group, "field": title, "short": short, "date": d,
                })
    out.sort(key=lambda m: (m["group"], m["omsu"], m["date"], m["id"]))
    return {"measures": out, "rm": RM}
