"""Bounded detail collection retaining the supplied report's chunk methodology."""
import collections
import concurrent.futures
import datetime
import threading
import time
from .client import PortalError
from .context import today


def year_ago(days):
    now = today()
    target = now - datetime.timedelta(days=days)
    try:
        target = target.replace(year=target.year - 1)
    except ValueError:
        target = target.replace(year=target.year - 1, day=28)
    return (now - target).days

CHUNK_DAYS = 30
WORKERS = 4                 # параллельных запросов к порталу, не больше
RETRIES = 2
RETRY_PAUSE = (1, 2)    # сек перед повторными попытками
ID_COL = "Внутренний Id"    # по нему склеиваем куски без дублей

# Запросы со строками обращений, по порядку предпочтения. 25.09.2026 основной
# начал отвечать 501 на любые параметры; запасной отдаёт те же колонки (плюс
# разбивку адреса) и после удаления дублей по ID сходится с итогом портала.
DETAIL_QUERIES = ["q_download_detalization", "q_download_detalization_address"]
_working_lock = threading.Lock()
_working = {"query": None}     # какой запрос жив — запоминаем на время работы


def chunks(start, end, size=CHUNK_DAYS):
    """Нарезать окно (start, end) на куски по size дней.

    Окно портала включает оба конца, поэтому следующий кусок начинается на
    день раньше конца предыдущего — без щелей и без перекрытий.
    """
    if (any(type(value) is not int for value in (start, end, size))
            or start < end or end < 0 or start > 15000 or not 1 <= size <= 31):
        raise ValueError('Некорректный период выгрузки.')
    out = []
    s = start
    while s >= end:
        e = max(s - size + 1, end)
        out.append((s, e))
        s = e - 1
    return out


def _fetch_one(pentaho, params, start, end):
    """Один кусок. Временные сбои повторяем, а на 501 сразу переходим к
    запасному запросу: 501 значит, что запрос сломан, и повтор не поможет."""
    p = dict(params)
    p["curr_period_start"] = "%d day" % start
    p["curr_period_end"] = "%d day" % end

    order = list(DETAIL_QUERIES)
    with _working_lock:
        preferred = _working['query']
    if preferred in order:
        order.remove(preferred)
        order.insert(0, preferred)

    last = None
    for qid in order:
        for attempt in range(RETRIES):
            try:
                res = pentaho.query(qid, p, timeout=180)
                with _working_lock:
                    _working["query"] = qid
                return res
            except PortalError as ex:
                last = ex
                if getattr(ex, "status", None) == 501:
                    break                    # этот запрос сломан — к следующему
                if not ex.retryable:
                    raise
                if attempt < RETRIES - 1:
                    time.sleep(RETRY_PAUSE[attempt])
    raise last


def _merge(parts):
    """Merge by identity and column name, including fallback query extra columns."""
    cols = list(dict.fromkeys(col for names, _ in parts for col in names))
    if not cols:
        return [], []
    seen, rows = set(), []
    for names, records in parts:
        pos = {col: index for index, col in enumerate(names)}
        index = pos.get(ID_COL)
        for record in records:
            if not isinstance(record, list) or len(record) != len(names):
                raise PortalError('Портал вернул некорректную строку детализации.')
            identity = record[index] if index is not None else None
            if identity not in (None, ''):
                identity = str(identity)
                if identity in seen:
                    continue
                seen.add(identity)
            rows.append([record[pos[col]] if col in pos else None for col in cols])
            if len(rows) > 300000:
                raise PortalError('Слишком много обращений. Сократите период.')
    return cols, rows


def fetch_periods(pentaho, base_params, periods):
    """Скачать строки для нескольких периодов сразу.

    periods: {ключ: (start, end)}. Все куски всех периодов идут через один
    общий пул, чтобы портал не получал больше WORKERS запросов одновременно.
    Возвращает ({ключ: (cols, rows)}, {ключ: текст ошибки}).
    """
    tasks = []
    for k, (s, e) in periods.items():
        for cs_, ce_ in chunks(s, e):
            tasks.append((k, cs_, ce_))

    if len(tasks) > 150:
        raise ValueError('Слишком большой объём выгрузки. Сократите период.')
    parts = collections.defaultdict(list)
    failed = collections.defaultdict(list)
    with concurrent.futures.ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futures = {ex.submit(_fetch_one, pentaho, base_params, s, e): (k, s, e)
                   for k, s, e in tasks}
        for f in concurrent.futures.as_completed(futures):
            k, s, e = futures[f]
            try:
                parts[k].append(f.result())
            except PortalError as ex_:
                failed[k].append((s, e, str(ex_)))
            except Exception:
                failed[k].append((s, e, 'Не удалось обработать ответ портала. Повторите сбор.'))

    results, errors = {}, {}
    for k in periods:
        if failed.get(k):
            # Частичные данные дали бы заниженные цифры без предупреждения,
            # поэтому период с дырой считаем неполученным целиком.
            spans = ", ".join(_span_label(s, e) for s, e, _ in failed[k])
            errors[k] = "не скачались куски %s: %s" % (spans, failed[k][0][2][:150])
            results[k] = ([], [])
        else:
            results[k] = _merge(parts[k])
    return results, errors


def _span_label(start, end):
    now = today()
    a = now - datetime.timedelta(days=start)
    b = now - datetime.timedelta(days=end)
    return "%s—%s" % (a.strftime("%d.%m.%Y"), b.strftime("%d.%m.%Y"))

