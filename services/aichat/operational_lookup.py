"""Read only the published inventory and safe columns from a saved complaint period."""
from __future__ import annotations

from contextlib import closing
from datetime import datetime
from functools import lru_cache
import json
import re
import sqlite3
import unicodedata
from zoneinfo import ZoneInfo

from fastapi import HTTPException

from core.roles import check_module_access
from services import report_municipalities as cities
from services.aichat import report_periods
from services.aichat.collective_source import _signatures
from services.aichat.report_context import _variants
from services.max_bot import stock
from utils import db

SOURCES = {
    'zips': ('Опубликованные остатки ЗиП РСО', '/zips'),
    'collective': ('Коллективные жалобы', '/mingkh/collective'),
}
# Intent words never expand the selected records. All remaining terms must
# match within a single row. Model numbers and dimensions retain their order.
_FILLER = re.compile(
    r'\b(?:покаж\w*|найд\w*|найти|поищи\w*|провер\w*|посмотр\w*|пожалуйста|'
    r'поиск\w*|сколько|каки[ех]|какой|какая|какое|каков\w*|можно|есть|ли|где|что|'
    r'данны\w*|сведени\w*|информаци\w*|опубликованн\w*|сохраненн\w*|список|'
    r'остат(?:ок|к\w*|ков)|зип\w*|рсо|материал\w*|позиц\w*|количеств\w*|'
    r'коллективн\w*|жалоб\w*|обращени\w*|карточк\w*|номеру?|'
    r'округ\w*|омсу|муниципалитет\w*|городск\w*|муниципальн\w*|'
    r'блок\w*|раздел\w*|систем\w*|платформ\w*|склад\w*|наличи\w*|'
    r'в|во|на|по|за|из|с|для|у|к|о|об|и|это|этом|этой|все|всех|всем|всему|всю|наш\w*|мне|нам|есть|нуж\w*)\b'
)
_PHONE = re.compile(r'(?<!\d)(?:\+7|8)(?:[\s().-]*\d){10}(?!\d)')
_EMAIL = re.compile(r'\b[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+')
_MARKED_ID = re.compile(r'(?:№|\bномер(?:ом|у|а)?\s*)\s*([\w-]*\d[\w./-]*)', re.I)
_ID = re.compile(r'(?<!\w)[a-zа-я\d-]*\d[\w./-]*(?!\w)', re.I)
_DATE = re.compile(r'\b(?:\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})\b')


def _plain(value, limit=180):
    # Return field values, never unlabelled source prose or imported links.
    value = ''.join(c for c in str(value or '')[:500] if unicodedata.category(c) not in ('Cf','Cc','Cs'))
    value = re.sub(r'\[([^]\n]{1,160})\]\([^)]*\)', r'\1', value)
    value = re.sub(r'<[^>]*>|(?:[a-z][a-z0-9+.-]{1,12}://|mailto:|www\.)\S+', '', value, flags=re.I)
    value = re.sub(r'\b(?:пароль|токен|token|password|api[_ -]?key|authorization)\s*[:=]\s*\S+',
                   '[скрыто]', value, flags=re.I)
    return ' '.join(_PHONE.sub('[контакт скрыт]', _EMAIL.sub('[контакт скрыт]', value)).split())[:limit]


def _answer(module, status, *, total=None, items=None, as_of=None, note=''):
    title, url = SOURCES[module]
    return {'title':title, 'url':url, 'scope':module, 'status':status,
            'as_of':as_of, 'total':total, 'items':items or [], 'note':note}


@lru_cache(maxsize=512)
def _city_pattern(name):
    variants = sorted((v for v in _variants(name) if v), key=len, reverse=True)
    return re.compile(r'(?<!\w)(?:' + '|'.join(map(re.escape, variants)) + r')(?!\w)')


def _city_filter(question, source_names):
    aliases, _ = cities._maps()  # Only the public municipality taxonomy.
    if len(source_names) > 1024:
        raise ValueError('unexpected municipality cardinality')
    names = set(aliases) | {cities.key(name) for name in source_names if cities.is_municipality(name)}
    selected = set()
    for name in sorted(names, key=len, reverse=True):
        # Cover compound city names before their component words.
        if not name or len(name) > 120:
            continue
        pattern = _city_pattern(name)
        if pattern.search(question):
            selected.add(cities.key(name))
            question = pattern.sub(' ', question)
    return selected, question


@lru_cache(maxsize=4)
def _inventory(directory):
    # StockService's contacts operation is never invoked. Its stock snapshot
    # is capped, deduplicated, refreshed on replacement and contains no contacts.
    return stock.StockService(directory / 'published.json', directory / 'contacts.json')


def _zips(question, limit):
    service = _inventory(db.DB_FILE.parent / 'zip_curator')
    try:
        snapshot = service._snapshot()
    except stock.StockError as error:
        return _answer('zips','unavailable',note=error.message)
    rows = snapshot['rows']
    requested, period_error = report_periods.extract(question)
    if requested or period_error:
        return _answer('zips','clarify',as_of=snapshot['published_at'],
                       note='Поиск по ЗиП использует текущую публикацию. Исторический остаток за произвольную дату по ней не подтверждён.')
    try:
        selected, rest = _city_filter(question, {row['municipality'] for row in rows})
    except ValueError:
        return _answer('zips','unavailable',note='Не удалось распознать муниципалитеты опубликованного свода.')

    available_only = bool(re.search(r'\bтолько\s+(?:в\s+наличии|актуальн\w*)\b', rest))
    rest = re.sub(r'\bтолько\s+(?:в\s+наличии|актуальн\w*)\b', ' ', rest)
    words = stock._tokens(_FILLER.sub(' ', rest), query=True)
    today = datetime.now(ZoneInfo('Europe/Moscow')).date()
    matched = []
    field_tokens = {}
    for row in rows:
        if selected and row['_municipality_key'] not in selected:
            continue
        missing = set(words).difference(row['_search_tokens'])
        if missing:
            fields = (row['category'], row['group'], row['organization'])
            if fields not in field_tokens:
                field_tokens[fields] = set(stock._tokens(' '.join(fields)))
            if not missing.issubset(field_tokens[fields]):
                continue
        availability = service._availability(row, today)
        if available_only and availability != 'available':
            continue
        matched.append((availability, row))
    order = {'available':0, 'stale':1, 'zero':2, 'unknown':3}
    matched.sort(key=lambda pair:(order[pair[0]], pair[1]['_municipality_key'], pair[1]['name'], pair[1]['id']))
    items = []
    for availability, row in matched[:limit]:
        quantity = f"{row['quantity']:g}" if row['quantity'] is not None else 'не указано'
        parts = [
            'Материал «' + _plain(row['name']) + '»',
            _plain(row['municipality']) or 'ОМСУ не указан',
            'РСО «' + _plain(row['organization']) + '»',
            'остаток ' + quantity + ' ' + _plain(row['unit'],24),
            'дата остатков ' + (row['stock_date'] or 'не указана'),
            stock._STATUS_LABELS[availability],
        ]
        items.append({'text':' · '.join(parts), 'availability':availability})
    note = 'Наличие подтверждается при положительном количестве, известной единице и дате остатков не старше 30 дней.'
    if snapshot['omitted']:
        note += ' Неполные строки публикации исключены.'
        if not matched:
            return _answer('zips','unavailable',as_of=snapshot['published_at'],
                           note=note+' Отсутствие подходящей позиции по неполному своду подтвердить нельзя.')

    return _answer('zips','ok' if matched else 'empty', total=len(matched),items=items,
                   as_of=snapshot['published_at'],note=note)


def _complaints(question, limit):
    requested, problem = report_periods.extract(question)
    if problem:
        return _answer('collective','clarify',note=problem)
    # Use one read transaction for period metadata, source numbers and safe
    # columns. Never select annotation, answers, applicant contacts, comments,
    # assignees or report text; never initialize or migrate an absent source.
    path = db.DB_FILE
    if not path.is_file():
        return _answer('collective','unavailable',note='Сохранённые данные коллективных жалоб недоступны.')
    try:
        with closing(sqlite3.connect(path.resolve().as_uri()+'?mode=ro', uri=True)) as conn:
            conn.row_factory = sqlite3.Row
            conn.execute('BEGIN')
            where = ' WHERE start=? AND end=?' if requested else ''
            args = [requested['from'], requested['to']] if requested else []
            head = conn.execute(
                'SELECT rowid AS id,start,end,fetched_at,length(rows_json) AS size,length(keys_json) AS key_size '
                'FROM collective_snapshots' + where + ' ORDER BY fetched_at DESC LIMIT 1', args).fetchone()
            if not head:
                return _answer('collective','unavailable',note='За запрошенный период нет сохранённой выборки.' if requested
                               else 'Коллективные жалобы ещё не сохранены.')
            from services.collective.store import MAX_SOURCE_BYTES, MAX_ROWS
            if max(head['size'],head['key_size']) > MAX_SOURCE_BYTES:
                raise ValueError('oversized snapshot')
            selector = ' FROM collective_snapshots s, json_each(s.rows_json) j WHERE s.rowid=?'
            names = [r[0] for r in conn.execute('SELECT DISTINCT json_extract(j.value,\'$[2]\')' + selector + ' LIMIT 2001',(head['id'],))]
            if len(names)>2000:
                raise ValueError('unexpected municipality cardinality')
            selected, rest = _city_filter(question, names)
            # A date in a period is not an appeal number.
            rest = _DATE.sub(' ',rest)
            if requested:
                rest = re.sub(r'\b(?:за\s+)?(?:последн\w*|текущ\w*|эт\w*|прошл\w*|предыдущ\w*)?\s*(?:\d+\s+)?(?:недел\w*|месяц\w*|дней|дня|день)\b',
                              ' ', rest)
            marked = _MARKED_ID.findall(rest)
            numbers = marked or _ID.findall(rest)
            if len(numbers)>1:
                return _answer('collective','clarify',note='Укажите один номер обращения для поиска.')
            number = numbers[0].casefold().rstrip('.,/') if numbers else ''
            if number:
                rest = rest.replace(number,' ')
            rest = re.sub(r'\b(?:сегодня|вчера|позавчера|период|до|статус\w*|результат\w*)\b',' ',rest)
            terms = stock._tokens(_FILLER.sub(' ',rest),query=True)
            keys = json.loads(conn.execute('SELECT keys_json FROM collective_snapshots WHERE rowid=?',(head['id'],)).fetchone()[0])
            if not isinstance(keys,list) or len(keys)>MAX_ROWS:
                raise ValueError('unexpected keys')
            statement = """SELECT j.key AS idx, json_array_length(j.value) AS width,
                       json_extract(j.value,'$[1]') AS received,json_extract(j.value,'$[2]') AS city,
                       json_extract(j.value,'$[5]') AS num,json_extract(j.value,'$[6]') AS secondary,
                       json_extract(j.value,'$[9]') AS fact,json_extract(j.value,'$[10]') AS signatures,
                       json_extract(j.value,'$[14]') AS state""" + selector
            city_keys = {name:cities.key(name) for name in names} if selected else {}
            seen, items, total, scanned = set(), [], 0, 0
            for row in conn.execute(statement,(head['id'],)):
                scanned += 1
                if row['width']!=16 or row['idx']>=len(keys) or not isinstance(keys[row['idx']],str):
                    raise ValueError('invalid source row')
                if selected and city_keys.get(row['city']) not in selected:
                    continue
                if number and not any(str(row[key] or '').casefold().removeprefix('№').strip()==number for key in ('num','secondary')):
                    continue
                values = set(stock._tokens(_plain(row['state'])+' '+_plain(row['fact'])))
                if any(term not in values for term in terms):
                    continue
                key = keys[row['idx']]
                if key in seen:
                    continue
                seen.add(key)
                total += 1
                if len(items) >= limit:
                    continue
                signatures = _signatures(row['signatures'])
                label = _plain(row['num'] or row['secondary'],80)
                parts = ['Обращение № '+label if label else 'Номер не указан', _plain(row['city']) or 'ОМСУ не указан',
                         _plain(row['received'],40), _plain(row['state'],80) or 'Статус не указан',
                         'подписей ' + (str(signatures) if signatures is not None else 'не указано')]
                if row['num'] and row['secondary'] and row['num'] != row['secondary']:
                    parts.append('номер источника '+_plain(row['secondary'],80))
                items.append({'text':' · '.join(parts)})
            if scanned != len(keys):
                raise ValueError('incomplete snapshot')
            note = f"Сохранённая выборка за {head['start']} — {head['end']}. Поиск по номеру, ОМСУ, факту и статусу."
            return _answer('collective','ok' if total else 'empty',total=total,items=items,as_of=head['fetched_at'],note=note)
    except (OSError,sqlite3.Error,ValueError,TypeError,KeyError,IndexError):
        return _answer('collective','unavailable',note='Не удалось прочитать сохранённую выборку коллективных жалоб.')


def search(user, module, query, limit=5):
    if not user:
        raise HTTPException(401,'Требуется авторизация.')
    if module not in SOURCES or not check_module_access(user,module):
        raise HTTPException(403,'Нет доступа к данным этого блока.')
    if not isinstance(query,str) or not query.strip() or len(query)>1200 or type(limit) is not int or not 1<=limit<=20:
        raise HTTPException(400,'Уточните поисковый запрос.')
    question = unicodedata.normalize('NFKC',query.replace('№',' номер ')).casefold().replace('ё','е').strip()
    if re.search(r'\b(?:парол\w*|токен\w*|почт\w*|телефон\w*|контакт\w*|заявител\w*)\b', question):
        return _answer(module,'clarify',note='Поиск в чате показывает только рабочие сведения о позициях и обращениях. Контактные и личные данные в выдачу не входят.')
    if module == 'collective' and re.search(r'\b(?:аннотаци\w*|текст\w*|доклад\w*|ответственн\w*|распределен\w*|комментари\w*)\b', question):
        return _answer(module,'clarify',note='Поиск по коллективным жалобам поддерживает номер обращения, ОМСУ, факт и статус. Тексты обращений и ручные доклады в выдачу не входят.')
    return _zips(question,limit) if module=='zips' else _complaints(question,limit)
