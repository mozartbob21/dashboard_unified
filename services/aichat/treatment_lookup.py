"""Bounded, read-only lookup in the saved treatment snapshot for the chat.

The model gets numbers, dates, enumerated topics/statuses and public municipal
names. Free text, addresses, coordinates, contacts and incident data are not
part of this search projection.
"""
from __future__ import annotations

from datetime import date
from functools import lru_cache
import re

from fastapi import HTTPException

from core.roles import check_module_access
from services import report_municipalities as municipalities
from services.aichat import report_periods
from services.aichat.report_context import _variants
from services.mingkh import treatment


TITLE = 'Жалобы по очистным сооружениям и изливам'
URL = '/mingkh/treatment'
STATUS = {'work': 'В работе', 'closed': 'Закрыта'}
TOPIC_WORDS = {
    'izliv': r'\b(?:излив\w*|сброс\w*(?:\s+сток\w*)?|разлив\w*)\b',
    'smell': r'\b(?:запах\w*|зловони\w*|вон\w*)\b',
    'kos': r'\bработ[аеуы]\s+(?:кос|очистн\w*(?:\s+сооружен\w*)?)\b',
}
GROUP_WORDS = {
    'work': r'\b(?:в\s+работе|открыт\w*|активн\w*|не[\s-]*закрыт\w*|не[\s-]*решен\w*)\b',
    'closed': r'\b(?:закрыт\w*|закрыл\w*|решенн\w*)\b',
}
FACILITY = r'\b(?:объект\w*|сооружени\w*|кос|бос|гос|оск)\b'
APPEAL = r'\b(?:жалоб\w*|обращени\w*)\b'
NAMED_FACILITY = re.compile(FACILITY + r'\s*[«"]([а-яa-z0-9 .№-]{2,80})[»"]')
DATE = r'\b(?:\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})\b'
# Words that describe a search or the source, rather than a matching condition.
NEUTRAL = re.compile(
    r'\b(?:каки\w*|сколько|есть|ли|всего|все[хм]?|найди(?:те)?|найти|покаж\w*|'
    r'поищи(?:те)?|проверь(?:те)?|поиск\w*|информаци\w*|количеств\w*|список|данн\w*|'
    r'последни\w*|сним\w*|раздел\w*|блок\w*|дашборд\w*|нейрон\w*|минжкх|'
    r'очистн\w*|сооружени\w*|объект\w*|кос|жалоб\w*|обращени\w*|омсу|'
    r'муниципалитет\w*|округ\w*|городск\w*|город\w*|территори\w*|статус\w*|тем\w*|'
    r'номер\w*|пожалуйста|этом|этим|этого|котор\w*|сейчас|наход\w*|со|с|по|к|у|и|'
    r'в|во|на|за|от|про|для|из|о|об|период\w*)\b')
GUIDE = 'Уточните ОМСУ, номер жалобы, объект КОС (номер или название в кавычках), тему либо статус. Свободные тексты обращений в поиске не используются.'


def _result(note, *, as_of=None, items=None, total=None):
    return {'title': TITLE, 'url': URL, 'as_of': as_of, 'total': total,
            'items': items or [], 'note': note, 'scope': 'mingkh'}


@lru_cache(maxsize=2)
def _city_patterns(names):
    # Reuse the chat's municipality declensions, including "в Дубне", "в Химках".
    return tuple((name, re.compile(r'(?<!\w)(?:' + '|'.join(re.escape(v) for v in
                    sorted(_variants(name), key=len, reverse=True)) + r')(?!\w)')) for name in names)


def _take(q, rules):
    found = set()
    for key, pattern in rules.items():
        if re.search(pattern, q):
            found.add(key)
            q = re.sub(pattern, ' ', q)
    return q, found


def _ids(q, pattern):
    matches = re.findall(pattern, q)
    return re.sub(pattern, ' ', q), {int(v) for v in matches}


def _linked(row, by_number, radii):
    distance = row.get('distance')
    radius = max(radii[row['topic']], 10 if row.get('mk') else 0)
    return by_number.get(row.get('near')) if distance is not None and distance <= radius else None


def search(user, query, limit=5):
    if not check_module_access(user, 'mingkh'):
        raise HTTPException(403, 'Нет доступа к дашборду по жалобам МинЖКХ.')
    if not isinstance(query, str) or not query.strip() or len(query) > 2000:
        return _result(GUIDE)
    q = query.casefold().replace('ё', 'е').strip()
    if re.search(r'\b(?:еддс|авари\w*|инцидент\w*)\b', q):
        return _result('Этот поиск охватывает жалобы и объекты КОС. Для событий используйте поиск по ЕДДС.')
    if re.search(r'\b(?:кроме|исключая|исключи\w*|либо|или)\b', q):
        return _result('Укажите одно условие или задайте отдельный вопрос для каждого варианта.')
    try:
        snapshot = treatment.overview(False)
    except (OSError, ValueError, EOFError):
        return _result('Сохранённый снимок не удалось прочитать.')
    if snapshot is None:
        return _result('В разделе ещё нет сохранённого снимка.')
    as_of = snapshot['as_of']
    end = date.fromisoformat(as_of)
    start = date.fromisoformat(snapshot['from'])
    period, error = report_periods.extract(q)
    if error:
        return _result('Не удалось определить период запроса. Укажите даты в формате ДД.ММ.ГГГГ.', as_of=as_of)
    if period and (date.fromisoformat(period['from']) < start or date.fromisoformat(period['to']) > end):
        return _result(f'Снимок охватывает {start.isoformat()} — {as_of}. Запрошенный период выходит за эти границы.', as_of=as_of)
    if period:
        q = re.sub(DATE, ' ', q)
        q = re.sub(r'\b(?:сегодня|вчера|позавчера)\b', ' ', q)

    source_asked = bool(re.search(r'\bочистн\w*\b', q))
    facility_asked = re.search(FACILITY, q) is not None
    requested_names = [name.strip() for name in NAMED_FACILITY.findall(q)]
    q = NAMED_FACILITY.sub(' ', q)
    aliases, registry = municipalities._maps()
    cities = set()
    for name, pattern in _city_patterns(tuple(registry)):
        if pattern.search(q):
            cities.add(name)
            q = pattern.sub(' ', q)
    appeal_asked = re.search(APPEAL, q) is not None
    q, object_ids = _ids(q, FACILITY + r'\s*(?:№\s*|номер\w*\s*|#\s*)?(\d{1,12})(?![\d-])')
    q, complaint_ids = _ids(q, APPEAL + r'\s*(?:№\s*|номер\w*\s*|#\s*)?(\d{1,16})(?![\d-])')
    if not object_ids and not complaint_ids:
        bare = re.fullmatch(r'\s*(?:№|#)?\s*(\d{1,16})\s*', q)
        if bare:
            complaint_ids.add(int(bare[1]))
            q = ''
        else:
            q, complaint_ids = _ids(q, r'(?:№|#)\s*(\d{1,16})\b')
    q, groups = _take(q, GROUP_WORDS)
    q, topics = _take(q, TOPIC_WORDS)
    q, states = _take(q, {'inactive':r'\bне[\s-]*действующ\w*\b', 'active':r'\bдействующ\w*\b'})
    q, linkage = _take(q, {
        'unlinked': r'\b(?:не\s*привязан\w*|без\s+привязки)\b',
        'linked': r'\b(?:привязан\w*|рядом)\b',
    })
    q, facility_types = _take(q, {'бос':r'\bбос\b','гос':r'\bгос\b','оск':r'\bоск\b',
                                  'поля фильтрации':r'\bполя\s+фильтрации\b'})
    unknown = re.findall(r'[а-яa-z0-9]+', NEUTRAL.sub(' ', q))
    if (unknown or len(linkage) > 1 or not any((cities, facility_asked, source_asked, appeal_asked,
            object_ids, complaint_ids, topics, groups, states, linkage, facility_types))):
        return _result(GUIDE, as_of=as_of)
    facility_mode = (facility_asked or bool(facility_types or states)) and not (
        appeal_asked or complaint_ids or groups or topics or linkage)
    if period and facility_mode:
        return _result('Для объектов доступно состояние на дату снимка. Период применяется к дате поступления жалобы.', as_of=as_of)
    cap = min(10, max(1, limit)) if type(limit) is int else 5
    by_number = {k['n']: k for k in snapshot['kos']}

    def known_city(row):
        # Never repeat an arbitrary imported string as a municipality (or prompt).
        raw = municipalities.key(row['omsu'])
        return aliases.get(raw, raw) if raw in registry or raw in aliases else None

    def matches_object(k):
        if object_ids and (k is None or k['n'] not in object_ids):
            return False
        name = str((k or {}).get('name') or '').casefold().replace('ё', 'е')
        if requested_names and not any(wanted == name for wanted in requested_names):
            return False
        if facility_types and not any(re.search(r'\b' + re.escape(t) + r'\b', name) for t in facility_types):
            return False
        if states:
            source_state = str((k or {}).get('state') or '').casefold().strip()
            state = {'действующие':'active', 'не действующие':'inactive', 'не дейтсвующие':'inactive',
                     'недействующие':'inactive'}.get(source_state)
            if state not in states:
                return False
        return True

    if facility_mode:
        rows = [k for k in sorted(by_number.values(), key=lambda k:k['n']) if matches_object(k)
                and (not cities or known_city(k) in cities)]
        items = []
        for k in rows[:cap]:
            city = registry.get(known_city(k))
            text = f"Объект КОС №{k['n']}" + (f' · {city}' if city else '')
            items.append({'text':text, 'facility_id':k['n'], 'municipality':city})
        return _result('Объекты реестра на дату снимка.', as_of=as_of, items=items, total=len(rows))

    matches = []
    for r in snapshot['complaints']:
        if complaint_ids and r['id'] not in complaint_ids:
            continue
        if (topics and r['topic'] not in topics) or (groups and r['sg'] not in groups):
            continue
        day = date.fromisoformat(r['created'][:10]).isoformat()
        if period and not period['from'] <= day <= period['to']:
            continue
        k = _linked(r, by_number, snapshot['radii'])
        if (not matches_object(k) or ('linked' in linkage and not k) or ('unlinked' in linkage and k)):
            continue
        city = known_city(k or r)
        if cities and city not in cities:
            continue
        matches.append((day, r['id'], r, k, registry.get(city)))
    matches.sort(key=lambda item:(item[0],item[1]), reverse=True)
    items = []
    for day, ident, row, k, city in matches[:cap]:
        facts = [f'Жалоба №{ident}', day, treatment.TOPICS[row['topic']], STATUS[row['sg']]]
        if city:
            facts.append(city)
        facts.append(f"КОС №{k['n']}" if k else 'Без привязки к КОС')
        items.append({'text':' · '.join(facts), 'complaint_id':ident, 'date':day,
                      'municipality':city, 'facility_id':k['n'] if k else None})
    note = 'Совпадения в снимке по дате поступления. ОМСУ привязанной жалобы определяется по объекту, как в дашборде.'
    if not (object_ids or facility_types or requested_names or linkage or states):
        note += ' Учтены также жалобы без привязки.'
    return _result(note, as_of=as_of, items=items, total=len(matches))
