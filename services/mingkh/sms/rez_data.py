# -*- coding: utf-8 -*-
"""Блок 2 СМС: строки обращений МинЖКХ с портала тепловой карты для поиска
резонанса.

Сама оценка резонанса — в браузере, rez/rez_core.js (дословно со старой
страницы). Здесь только данные:

  * хранилище строк по месяцам data/rez_rows/ГГГГ-ММ.json — история адресов
    за HISTORY_DAYS (старая страница работала на двухлетнем архиве: «тянется
    больше 180 дней» без длинной истории не посчитать);
  * выбранный период и период сравнения всегда берутся заново с портала,
    более старые месяцы — из хранилища, если они уже скачаны после своего
    окончания;
  * отбор строк под фильтры страницы и истории для адресов периода.
"""
import calendar
import concurrent.futures
import datetime
import os
import re
import math
from pathlib import Path
import threading

from . import svod_obrashcheniya as svod, dataset
from .client import Pentaho, PortalError
from .filters import SOURCES, SUBTOPICS
from .context import atomic_json, json_read, process_lock, date_range, timestamp, today as local_today

STORE = os.path.join(svod.DATA, "rez_rows")
HISTORY_DAYS = 730
MAX_ROWS = 300000
MAX_FOUND = 100000
MAX_CASES = 5000

# Колонки детализации портала -> поля обращения (как ждёт makeItem в rez_core.js).
COLS = {
    "srcnum": "Номер в источнике",
    "ecur": "номер ЕЦУР",
    "innid": "Внутренний Id",
    "src": "Источник",
    "spam": "Спам (да/нет)",
    "type": "Тип сообщения - 0 проблемы, 1 - предл",
    "desc": "description",
    "omsu": "ОМСУ",
    "addr": "Адрес",
    "date": "Дата (первого взятия в работу)",
    "mail": "petitioner_email",
    "subtopic": "Подтема",
    "dept": "Исполнитель",
}
FIELDS = list(COLS)


# ---------- ключ адреса: shortAddr/addrKey из rez_core.js ----------
# Нужен серверу только чтобы выбрать историю нужных адресов. Совпадение с JS
# проверено на всех адресах хранилища (см. rez/_cmp).

def short_addr(full):
    parts = [s.strip() for s in str(full or "").split(",")]
    parts = [p for p in parts if p
             and not re.match(r"^россия$", p, re.I)
             and not re.match(r"^московская\s+область$", p, re.I)
             and not re.match(r"^\d{6}$", p)]
    out = []
    for p in parts:
        key = p.lower().replace("ё", "е")
        if not out or out[-1].lower().replace("ё", "е") != key:
            out.append(p)
    return ", ".join(out)


def addr_key(full):
    s = short_addr(full).lower().replace("ё", "е")
    s = re.sub(r"[.,]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


# ---------- хранилище по месяцам ----------

def _month_bounds(y, m):
    return datetime.date(y, m, 1), datetime.date(y, m, calendar.monthrange(y, m)[1])


def _months(a, b):
    y, m = a.year, a.month
    while (y, m) <= (b.year, b.month):
        yield y, m
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)


def _path(y, m):
    _month_bounds(y, m)  # Reject invalid dates before constructing a fixed file name.
    return os.path.join(STORE, "%04d-%02d.json" % (y, m))


def _read(y, m):
    value = json_read(_path(y, m))
    if (not isinstance(value, dict) or value.get('fields') != FIELDS
            or not isinstance(value.get('rows'), list) or len(value['rows']) > MAX_ROWS
            or any(not isinstance(row, list) or len(row) != len(FIELDS) for row in value['rows'])):
        raise ValueError('Архив резонансных обращений повреждён. Повторите сбор.')
    return value


def _complete(y, m):
    """Месяц скачан целиком: файл есть и снят позже, чем месяц закончился
    (обращения ещё неделю после конца месяца доклассифицируются — берём запас)."""
    if not os.path.exists(_path(y, m)):
        return False
    _, end = _month_bounds(y, m)
    try:
        taken = datetime.date.fromisoformat(_read(y, m)['fetched_at'][:10])
        return taken >= end + datetime.timedelta(days=7)
    except (ValueError, KeyError, TypeError):
        return False


def _excluded(subtopic):
    return subtopic in svod.EXCLUDED_TITLES


def _ensure(date_from, date_to, fresh_from, log=None):
    """Месяцы [date_from, date_to] в хранилище. Месяцы начиная с fresh_from и
    недокачанные — заново с портала. Возвращает число скачанных месяцев."""
    log = log or (lambda message: None)
    need = [(y, m) for y, m in _months(date_from, date_to)
            if (y, m) >= (fresh_from.year, fresh_from.month) or not _complete(y, m)]
    if not need:
        return 0
    today = local_today()
    p = Pentaho()
    jobs = {}
    for y, m in need:
        a, b = _month_bounds(y, m)
        b = min(b, today)
        jobs[(y, m)] = ((today - a).days, (today - b).days)
    log("с портала месяцев: %d" % len(jobs))
    base = p.periods(10, 0, 20, 11)
    tasks = [(k, s, e) for k, (s0, e0) in jobs.items() for s, e in dataset.chunks(s0, e0)]

    # Куски качаем так же, как тепловая карта (_fetch_one: повторы, переход на
    # запасной запрос при 501), но каждый кусок разбираем по ЕГО колонкам:
    # dataset._merge берёт колонки первого куска, и когда куски пришли от
    # разных запросов портала (основной/запасной), индексы съезжают —
    # 16.10.2026 так упало заполнение истории.
    def one(task):
        k, s, e = task
        cols, rows = dataset._fetch_one(p, base, s, e)
        ix = {c: i for i, c in enumerate(cols)}
        if not {COLS['innid'], COLS['date'], COLS['subtopic'], COLS['addr']} <= set(ix):
            raise PortalError('Портал не вернул обязательные колонки обращений.')
        pos = {field: ix.get(column) for field, column in COLS.items()}
        result = []
        for row in rows:
            if not isinstance(row, list) or len(row) != len(cols):
                raise PortalError('Портал вернул некорректную строку обращений.')
            packed = [row[pos[field]] if pos[field] is not None else None for field in FIELDS]
            if packed[FIELDS.index('innid')] in (None, ''):
                raise PortalError('Портал вернул обращение без идентификатора.')
            result.append(packed)
        return k, result

    got = {k: {} for k in jobs}
    with concurrent.futures.ThreadPoolExecutor(max_workers=dataset.WORKERS) as ex:
        for k, rows in ex.map(one, tasks):
            ii, si = FIELDS.index("innid"), FIELDS.index("subtopic")
            for r in rows:
                if not _excluded(r[si]):
                    got[k][str(r[ii])] = r       # без дублей по «Внутренний Id»
            if sum(len(records) for records in got.values()) > MAX_ROWS:
                raise PortalError('Слишком много обращений. Сократите период.')
    now = timestamp()
    for (y, m), rows in got.items():
        atomic_json(_path(y, m), {'fetched_at': now, 'fields': FIELDS, 'rows': list(rows.values())})
        log("  %04d-%02d: %d строк" % (y, m, len(rows)))
    return len(jobs)


def ensure(date_from, date_to, fresh_from, log=None):
    date_from, date_to = date_range(date_from, date_to, max_days=3 * 366 + HISTORY_DAYS)
    if not date_from <= fresh_from <= date_to:
        raise ValueError('Некорректная дата обновления истории.')
    with process_lock(Path(STORE).parent / '.rez-refresh.lock', timeout=0):
        return _ensure(date_from, date_to, fresh_from, log)


def rows_between(a, b):
    """Строки хранилища с датой в [a, b] — словари с полями FIELDS."""
    ia, ib = a.isoformat(), b.isoformat()
    di = FIELDS.index("date")
    out = []
    for y, m in _months(a, b):
        if not os.path.exists(_path(y, m)):
            continue
        for r in _read(y, m)["rows"]:
            d = str(r[di] or "")[:10]
            if ia <= d <= ib:
                out.append(dict(zip(FIELDS, r)))
                if len(out) > MAX_ROWS:
                    raise ValueError('Слишком много обращений в архиве. Сократите период.')
    return out


# ---------- выбор под фильтры страницы ----------

def collect(req, log=None):
    """Данные для страницы резонансных.

    req: from, to, prev_from, prev_to (ISO), subtopics (id), sources (id),
         omsu (id или -1), dept (-1 или id ФКР).
    Тематики и исполнитель режут и период, и историю; источники и ОМСУ — только
    сами периоды: история адреса — это все его жалобы по выбранной тематике
    (на старой странице источник баллов не давал).
    """
    if not isinstance(req, dict):
        raise ValueError('Некорректные параметры сбора.')
    f, t = date_range(req.get('from'), req.get('to'), max_days=366)
    pf, pt = date_range(req.get('prev_from'), req.get('prev_to'), max_days=366)
    if (max(t, pt) - min(f, pf)).days > 1096:
        raise ValueError('Периоды сравнения расположены слишком далеко друг от друга.')
    for key, allowed in (('subtopics', dict(SUBTOPICS)), ('sources', dict(SOURCES))):
        values = req.get(key) or []
        if not isinstance(values, list) or len(values) > len(allowed) or any(str(v) not in allowed for v in values):
            raise ValueError('Неизвестный фильтр обращений.')
    if str(req.get('dept') or '-1') not in ('-1', svod.DEPT_FKR):
        raise ValueError('Неизвестный исполнитель.')
    if not re.fullmatch(r'-1|[0-9]{1,10}', str(req.get('omsu') or '-1')):
        raise ValueError('Некорректный муниципалитет.')
    start = min(f, pf) - datetime.timedelta(days=HISTORY_DAYS)
    fetched = ensure(start, max(t, pt), min(f, pf), log)

    sub_ids = {str(value) for value in (req.get('subtopics') or [])}
    sub_titles = {title for sid, title in SUBTOPICS
                  if sid != '-1' and title not in svod.EXCLUDED_TITLES
                  and (not sub_ids or '-1' in sub_ids or sid in sub_ids)}
    src_ids = {str(value) for value in (req.get('sources') or [])}
    src_titles = {tt for sid, tt in SOURCES if sid in src_ids} if src_ids and '-1' not in src_ids else None
    omsu = str(req.get("omsu") or "-1")
    omsu_title = dict(svod.omsu_list()).get(omsu) if omsu != "-1" else None
    if omsu != '-1' and omsu_title is None:
        raise ValueError('Выбранный муниципалитет отсутствует в справочнике портала.')
    dept_title = svod.FKR_TITLE if str(req.get("dept") or "-1") == svod.DEPT_FKR else None

    def topical(r):
        return r["subtopic"] in sub_titles and (dept_title is None or r["dept"] == dept_title)

    def in_scope(r):
        return ((src_titles is None or r["src"] in src_titles)
                and (omsu_title is None or r["omsu"] == omsu_title))

    allrows = [r for r in rows_between(start, max(t, pt)) if topical(r)]
    fi, ti, pfi, pti = f.isoformat(), t.isoformat(), pf.isoformat(), pt.isoformat()
    cur = [r for r in allrows if fi <= str(r["date"])[:10] <= ti and in_scope(r)]
    prev = [r for r in allrows if pfi <= str(r["date"])[:10] <= pti and in_scope(r)]
    keys = {addr_key(r["addr"]) for r in cur + prev if r["addr"]}
    taken = {id(r) for r in cur + prev}
    hist = [r for r in allrows if id(r) not in taken and r["addr"] and addr_key(r["addr"]) in keys]

    pack = lambda rs: [[r[k] for k in FIELDS] for r in rs]   # noqa: E731
    return {"fields": FIELDS, "current": pack(cur), "prev": pack(prev), "history": pack(hist),
            "history_from": start.isoformat(), "fetched_months": fetched}


# ---------- память найденного ----------
# data/rez_found.json: ключ адреса -> когда впервые найден резонансным.
# «Новый» резонанс за период [from, to] — адрес, которого в базе нет или который
# впервые найден в этом же (или более позднем) периоде: повторный прогон того
# же периода даёт тот же результат.

FOUND = os.path.join(svod.DATA, "rez_found.json")
_found_lock = threading.RLock()


def found_load():
    if not os.path.exists(FOUND):
        return {}
    return validate_found_data(json_read(FOUND, max_bytes=64 * 1024 * 1024))


def _case(case):
    if not isinstance(case, dict) or set(case) - {'key', 'addr', 'omsu', 'score'}:
        raise ValueError('Некорректная запись резонансного адреса.')
    result = {}
    for key, limit in (('key', 1500), ('addr', 1500), ('omsu', 300)):
        value = case.get(key) or ''
        if not isinstance(value, str) or len(value) > limit or any(ord(char) < 32 for char in value):
            raise ValueError('Слишком длинное или некорректное название адреса.')
        result[key] = value.strip()
    if not result['key'] or not result['addr']:
        raise ValueError('Укажите адрес и его ключ.')
    score = case.get('score')
    if type(score) not in (int, float) or not math.isfinite(score) or not 0 <= score <= 1000000:
        raise ValueError('Некорректная оценка резонансного адреса.')
    result['score'] = score
    return result


def validate_found_data(value):
    """Validate a complete imported discovery store without writing anything."""
    if not isinstance(value, dict) or len(value) > MAX_FOUND:
        raise ValueError('Некорректная или слишком большая память найденных адресов.')
    result = {}
    for key, record in value.items():
        if not isinstance(record, dict) or set(record) - {'first_from', 'first_to', 'last_from', 'last_to', 'addr', 'omsu', 'score'}:
            raise ValueError('Некорректная запись в памяти найденных адресов.')
        case = _case({'key': key, 'addr': record.get('addr'), 'omsu': record.get('omsu'), 'score': record.get('score')})
        first_from, first_to = date_range(record.get('first_from'), record.get('first_to'), max_days=366)
        first = {'first_from': first_from.isoformat(), 'first_to': first_to.isoformat()}
        if 'last_from' in record or 'last_to' in record:
            last_from, last_to = date_range(record.get('last_from'), record.get('last_to'), max_days=366)
            if last_to < first_from:
                raise ValueError('Некорректный порядок дат в памяти адресов.')
            first.update(last_from=last_from.isoformat(), last_to=last_to.isoformat())
        ident = case.pop('key')
        result[ident] = {**case, **first}
    return result


def merge_found_data(existing, imported):
    """Pure validated merge for the caller's all-or-nothing archive transaction."""
    base, clean = validate_found_data(existing), validate_found_data(imported)
    for key, incoming in clean.items():
        previous = base.get(key)
        if previous is None:
            base[key] = incoming
            continue
        first = incoming if (incoming['first_from'], incoming['first_to']) < (previous['first_from'], previous['first_to']) else previous
        latest = max((previous, incoming), key=lambda row: row.get('last_to', row['first_to']))
        base[key] = dict(first, last_from=latest.get('last_from', latest['first_from']), last_to=latest.get('last_to', latest['first_to']))
    if len(base) > MAX_FOUND:
        raise ValueError('Слишком много адресов после объединения архивов.')
    return base


def merge_found(imported):
    """Merge validated history, preserving earliest discovery and latest observation."""
    clean = validate_found_data(imported)
    with _found_lock, process_lock(Path(FOUND).with_suffix('.lock')):
        base = merge_found_data(found_load(), clean)
        atomic_json(FOUND, base)
        return {'total': len(base)}


def found_upsert(date_from, date_to, cases):
    """Persist the earliest discovery; validate the entire request before writing."""
    date_from, date_to = (value.isoformat() for value in date_range(date_from, date_to, max_days=366))
    if not isinstance(cases, list) or len(cases) > MAX_CASES:
        raise ValueError('Слишком много резонансных адресов в одном запросе.')
    clean = [_case(case) for case in cases]
    with _found_lock, process_lock(Path(FOUND).with_suffix('.lock')):
        base = found_load()
        added = 0
        for case in clean:
            record = base.get(case['key'])
            if record is None or record['first_from'] > date_from:
                previous_last = {key: record[key] for key in ('last_from', 'last_to') if record and key in record}
                base[case['key']] = record = {
                    'first_from': date_from, 'first_to': date_to,
                    'addr': case['addr'], 'omsu': case['omsu'], 'score': case['score'], **previous_last,
                }
                added += 1
            if record.get('last_to', '') < date_to:
                record['last_from'], record['last_to'] = date_from, date_to
        if len(base) > MAX_FOUND:
            raise ValueError('Достигнут предел памяти адресов. Обратитесь к администратору.')
        atomic_json(FOUND, base)
        return {'added': added, 'total': len(base)}
