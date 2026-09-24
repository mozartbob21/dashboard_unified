"""Period comparisons and the original dashboard's compact browser dataset."""
import copy
import hashlib
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from .client import Pentaho, PortalError, periods, year_ago

DIMS=['omsu','source','direction','theme','subtopic','fact','executor']
COLS=dict(zip(DIMS,['ОМСУ','Источник','Направление','Синт. группа','Подтема','Факт','Исполнитель']))
EXCLUDE={'subtopic':{'Вне компетенции (МИНЖКХ)','Платежные документы (ЕИРЦ)','Деятельность УК\\Кооперативов'},
         'theme':{'Вне компетенции (МИНЖКХ)'},'direction':{'Вне компетенции Ведомств МО'}}
PRESETS=[('thucur','Текущая неделя с чт'),('thuwed','Неделя чт—ср'),('d7','7 дней'),('d14','14 дней'),('d30','30 дней'),('d90','Квартал'),('d365','Год')]
CACHE={}
LOCK=threading.Lock()
CHUNK_DAYS = 30
WORKERS = 4
RETRY_PAUSE = (2, 5)
MAX_ROWS = 300000


def chunks(start, end, size=CHUNK_DAYS):
    """Inclusive intervals, with neither missing nor overlapping boundary dates."""
    if size < 1 or start < end:
        raise ValueError('Некорректный период.')
    while start >= end:
        stop = max(end, start - size + 1)
        yield start, stop
        start = stop - 1


def fetch_chunk(portal, params, start, end):
    p = dict(params)
    p.update(curr_period_start=f'{start} day', curr_period_end=f'{end} day')
    for attempt in range(len(RETRY_PAUSE) + 1):
        try:
            return portal.query('q_download_detalization', p, timeout=300)
        except PortalError as exc:
            # Invalid credentials, schema and TLS errors are not helped by retries.
            if not exc.retryable or attempt == len(RETRY_PAUSE):
                raise
            time.sleep(RETRY_PAUSE[attempt])


def merge_chunks(parts):
    columns = next((c for c, _ in parts if c), [])
    if len(columns) != len(set(columns)):
        raise PortalError('В выгрузке повторяются названия столбцов.')
    rows, seen = [], set()
    id_index = columns.index('Внутренний Id') if 'Внутренний Id' in columns else None
    for names, values in parts:
        if names != columns:
            if len(names) != len(columns) or set(names) != set(columns):
                raise PortalError('Столбцы частей выгрузки не совпадают. Повторите загрузку.')
            order = [names.index(c) for c in columns]
        else:
            order = list(range(len(columns)))
        for value in values:
            if not isinstance(value, list) or len(value) != len(names):
                raise PortalError('Портал вернул неполную строку обращения.')
            row = [value[i] for i in order]
            ident = str(row[id_index]).strip() if id_index is not None and row[id_index] is not None else ''
            if ident:
                if ident in seen:
                    continue
                seen.add(ident)
            rows.append(row)
            if len(rows) > MAX_ROWS:
                raise PortalError('В периоде больше 300 000 обращений. Выберите меньший период.')
    return columns, rows


def fetch_periods(portal, params, spans):
    jobs = [(key, start, end) for key, span in spans.items() for start, end in chunks(*span)]
    parts = {key: [] for key in spans}
    errors = {}
    counts = {key: 0 for key in spans}
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(fetch_chunk, portal, params, start, end): (key, start, end)
                   for key, start, end in jobs}
        for future in as_completed(futures):
            key, start, end = futures.pop(future)
            if key in errors:
                continue
            try:
                result = future.result()
                counts[key] += len(result[1])
                if counts[key] > MAX_ROWS:
                    raise PortalError('В периоде больше 300 000 строк. Выберите меньший период.')
                parts[key].append((start, result))
            except PortalError as exc:
                errors[key] = f'{label(start, end)}: {exc}'
                parts[key].clear()
                # Release failed-period payloads and avoid unnecessary portal work.
                for pending, (period_key, _, _) in futures.items():
                    if period_key == key:
                        pending.cancel()
    # No partial period may silently turn into lower totals or zeroes.
    result = {key: merge_chunks([part for _, part in sorted(values, reverse=True)])
              for key, values in parts.items() if key not in errors}
    return result, errors


def resolve_periods(query,today=None):
    today=today or date.today()
    preset=query.get('preset','thucur')
    def offset(key):
        try:
            value=date.fromisoformat(query.get(key,''))
        except ValueError:
            raise ValueError('Укажите корректные даты периода.')
        return (today-value).days
    if preset=='thucur':cs,ce=(today.weekday()-3)%7,0
    elif preset=='thuwed':
        ce=(today.weekday()-2)%7 or 7
        cs=ce+6
    elif preset in dict(PRESETS) and preset.startswith('d'):cs,ce=int(preset[1:])-1,0
    elif preset=='custom':cs,ce=offset('curr_from'),offset('curr_to')
    else:raise ValueError('Неизвестный период.')
    if query.get('cmp','auto')=='custom':ps,pe=offset('prev_from'),offset('prev_to')
    elif query.get('cmp','auto')=='auto':
        shift=7 if preset=='thucur' else cs-ce+1
        ps,pe=cs+shift,ce+shift
    else:raise ValueError('Неизвестный период сравнения.')
    for start,end in [(cs,ce),(ps,pe)]:
        if end<0 or start<end or start-end>365 or start>1095:
            raise ValueError('Выберите период до 366 дней, без будущих дат и не старше трёх лет.')
    return cs,ce,ps,pe


def label(start,end,today=None):
    today=today or date.today()
    weekdays=['пн','вт','ср','чт','пт','сб','вс']
    def render(offset):
        d=today-timedelta(days=offset)
        return weekdays[d.weekday()]+' '+d.strftime('%d.%m.%Y')
    return render(start)+' — '+render(end)


def build(portal,windows):
    cs,ce,ps,pe=windows
    params=periods(cs,ce,ps,pe)
    spans={'curr':(cs,ce),'prev':(ps,pe),'appg':(year_ago(cs),year_ago(ce))}
    raw, errors = fetch_periods(portal, params, spans)
    if 'curr' not in raw:
        raise PortalError(errors.get('curr','Не удалось загрузить текущий период.'))
    if any(key not in raw for key in ('prev','appg')):
        raise PortalError('Не удалось получить все периоды сравнения. Повторите загрузку: неполные данные не показаны как нулевые.')
    values={dim:[] for dim in DIMS};indexes={dim:{} for dim in DIMS}
    packed={};excluded={}
    for key,(columns,rows) in raw.items():
        if not all(column in columns for column in COLS.values()):
            raise PortalError('Выгрузка портала изменилась: отсутствуют ожидаемые столбцы.')
        positions={dim:columns.index(column) for dim,column in COLS.items()}
        output=[];excluded[key]=0
        for row in rows:
            if not isinstance(row,list) or len(row)<=max(positions.values()):
                raise PortalError('Портал вернул неполную строку. Уменьшите период и повторите.')
            if any(row[positions[dim]] in items for dim,items in EXCLUDE.items()):
                excluded[key]+=1;continue
            encoded=[]
            for dim in DIMS:
                value=str(row[positions[dim]] or '—')
                if value not in indexes[dim]:
                    indexes[dim][value]=len(values[dim]);values[dim].append(value)
                encoded.append(indexes[dim][value])
            output.append(encoded)
        packed[key]=output
    population={}
    try:
        columns,rows=portal.query('q_map',params)
        title,pop=columns.index('title'),columns.index('population')
        for row in rows:
            value=float(row[pop] or 0)
            population[str(row[title])]=max(0,value)
    except (PortalError,ValueError,IndexError,TypeError):
        errors['population']='Нет данных о населении.'
    try:updated,not_mapped=portal.last_updated()
    except PortalError:updated,not_mapped='Нет данных',None
    return {'dims':values,'rows':packed,'population':population,'not_municipal':['Москва','Московская область'],
            'errors':errors,'counts':{key:len(packed.get(key,[])) if key in packed else None for key in spans},
            'excluded':excluded,'curr_label':label(cs,ce),'prev_label':label(ps,pe),
            'appg_label':label(year_ago(cs),year_ago(ce)), 'updated':updated,'not_mapped':not_mapped,
            'fetched_at':datetime.now().strftime('%H:%M:%S')}


def get_dataset(creds,query):
    windows=resolve_periods(query)
    identity=hashlib.sha256((creds['username']+'\0'+creds['password']).encode()).hexdigest()
    key=(identity,windows,date.today())
    with LOCK:
        cached=CACHE.get(key)
        if cached and ((query.get('fresh')!='1' and time.monotonic()-cached[0]<600) or time.monotonic()-cached[0]<10):
            return copy.deepcopy(cached[1])
        result=build(Pentaho(**creds),windows)
        if len(CACHE)>=8:CACHE.pop(next(iter(CACHE)))
        CACHE[key]=(time.monotonic(),result)
        return copy.deepcopy(result)
