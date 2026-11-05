"""Ground a complaint breakdown in one authorised, saved MinZhKH dataset at a time.

Only reviewed classification labels leave this adapter. In particular an imported
fact cell is data from a file, never a user or system instruction.
"""
from collections import Counter
from datetime import date, timedelta
from functools import lru_cache
import json
import re
import zlib

from fastapi import HTTPException
from core.roles import check_module_access
from services import report_municipalities as cities
from services.aichat import local_sources
from services.mingkh import dashboard, water_map
from services.mingkh.sms.filters import SUBTOPICS

_TITLES = {'dashboard': ('Дашборд по жалобам МинЖКХ', '/mingkh'),
           'map': ('Карта жалоб МинЖКХ (Добродел)', '/mingkh/water-map')}
_UNASSIGNED = 'Без классификации'
_UNALLOCATED = 'Без подтверждённого ОМСУ'
MAX_GROUPS = 10
MAX_PERIOD_BUCKETS = 4096
MAX_PROJECTION_BYTES = 4 * 1024 * 1024
_REPORT_PATH = local_sources.BASE / 'data/mingkh/report-summary.json'


def _plain(value):
    return ' '.join(value.casefold().replace('ё','е').split()) if isinstance(value,str) else ''


def _read(path):
    try:
        return local_sources._load(path)
    except (EOFError, zlib.error, RecursionError):
        # A truncated gzip, invalid stream, or excessively nested JSON is unavailable.
        return {}


@lru_cache(maxsize=1)
def _taxonomy():
    # Only the bundled dashboard and the source-controlled SMS taxonomy supply labels.
    packaged = _read(water_map.SEED_FILE)
    labels = [title for ident,title in SUBTOPICS if ident != '-1']
    for field in ('fact','subcat'):
        labels.extend(packaged.get(field, ()))
    allowed = {_plain(label): label for label in labels
               if isinstance(label,str) and 1 < len(label) <= 320
               and _plain(label) not in ('нет данных','прочие','прочее')}
    city_names = {cities.key(name): cities.display(name) for name in packaged.get('omsu',[])
                  if cities.is_municipality(name)}
    return allowed, city_names


def _known_cities():
    return {**_taxonomy()[1], **cities._maps()[1]}


def _place(name, known):
    return known.get(cities.key(name), _UNALLOCATED)


def _label(fact, subtopic):
    labels, _ = _taxonomy()
    return labels.get(_plain(fact)) or labels.get(_plain(subtopic)) or _UNASSIGNED


def _indexed(values, index):
    if type(index) is not int or not 0 <= index < len(values):
        raise ValueError('Некорректный индекс справочника.')
    return values[index]


def project_dashboard(dataset):
    """Add bounded period/city/fact counts to the existing atomic dashboard projection."""
    known = _known_cities()
    dimensions = dataset.get('dims', {})
    if not isinstance(dimensions,dict) or not all(isinstance(dimensions.get(key),list) for key in ('omsu','fact','subtopic')):
        return {}
    indexes = {name: dashboard.DIMS.index(name) for name in ('omsu','fact','subtopic')}
    result = {}
    remaining = MAX_PROJECTION_BYTES - 64
    for key in ('curr','prev','appg'):
        raw = dataset.get('rows',{}).get(key)
        if not isinstance(raw,list) or len(raw)>dashboard.MAX_ROWS:
            continue
        counts = {}
        buckets = 0
        try:
            for row in raw:
                city = _place(_indexed(dimensions['omsu'],row[indexes['omsu']]),known)
                label = _label(_indexed(dimensions['fact'],row[indexes['fact']]),
                               _indexed(dimensions['subtopic'],row[indexes['subtopic']]))
                group = counts.setdefault(city,Counter())
                if label not in group:
                    buckets += 1
                    if buckets > MAX_PERIOD_BUCKETS:
                        raise ValueError('Слишком много сочетаний ОМСУ и факта.')
                group[label] += 1
        except (ValueError,IndexError,TypeError,KeyError):
            continue  # A partial period must never be reported as a complete snapshot.
        # Use the writing serializer's default separators, with room for wrapper keys.
        size = len(json.dumps(counts, ensure_ascii=False).encode('utf-8'))
        if size <= remaining:
            result[key] = counts
            remaining -= size
    return result


def _response(status, source, city, note, *, period=None, as_of=None, counts=None):
    title, url = _TITLES[source]
    total = sum(counts.values()) if counts is not None else None
    unknown = counts.get(_UNASSIGNED,0) if counts is not None else None
    classified = sorted(((label,count) for label,count in (counts or {}).items()
                         if label != _UNASSIGNED and count), key=lambda item: (-item[1],item[0]))
    groups = [{'label':label,'count':count} for label,count in classified[:MAX_GROUPS]]
    return {'status':status, 'scope':'mingkh', 'title':title, 'url':url, 'as_of':as_of,
            'municipality':city, 'period':period, 'total':total, 'groups':groups,
            'unknown':unknown, 'other':sum(count for _,count in classified[MAX_GROUPS:]), 'note':note}


def _dates(first,last):
    first,last=date.fromisoformat(first),date.fromisoformat(last)
    if first>last:
        raise ValueError('Некорректные границы периода')
    return {'from':first.isoformat(),'to':last.isoformat()}


def _stamp(raw):
    # Only dates from the source, no arbitrary imported metadata in chat.
    if isinstance(raw,str) and re.fullmatch(r'(?:\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})(?:[T ]\d{2}:\d{2}(?::\d{2}(?:[.]\d{1,6})?)?(?:[+]\d{2}:\d{2}|Z)?)?',raw):
        return raw
    return None


def _dashboard_result(city, requested):
    summary=_read(_REPORT_PATH)
    periods=summary.get('periods',{})
    stored=summary.get('causes_v1',{})
    if summary.get('schema_version')!=1 or not isinstance(stored,dict) or not isinstance(periods,dict):
        return None
    for key in (('curr','prev','appg') if requested else ('curr',)):
        info=periods.get(key)
        counts_by_city=stored.get(key)
        if not isinstance(info,dict) or not isinstance(counts_by_city,dict) or info.get('invalid_rows'):
            continue
        try:
            bounds=_dates(*info['bounds'])
            if requested and requested != bounds:
                continue
            counts=Counter()
            recorded=0
            for name, grouped in counts_by_city.items():
                if not isinstance(grouped,dict):
                    raise ValueError
                for label,value in grouped.items():
                    if type(value) is not int or value<0:
                        raise ValueError
                    recorded+=value
                    if not city or name==city:
                        counts[_label(label,None)]+=value
            if info.get('total') != recorded:
                continue
            note='Сохранённая выборка дашборда с его фильтрами. Группы по факту ЕЦУР, при отсутствии известного факта — по подтеме.'
            if city and not counts and counts_by_city.get(_UNALLOCATED):
                return _response('unavailable','dashboard',city,'В выборке есть обращения без подтверждённого ОМСУ. Отсутствие жалоб здесь не установлено.',period=bounds)
            return _response('ok','dashboard',city,note,period=bounds,counts=counts,
                             as_of=_stamp(summary.get('source_updated')) or _stamp(summary.get('collected_at')))
        except (ValueError,TypeError,KeyError,IndexError):
            continue
    return None


def _dashboard_missing(city, requested):
    summary = _read(_REPORT_PATH)
    stamp = _stamp(summary.get('source_updated')) or _stamp(summary.get('collected_at'))
    newest = (summary.get('periods') or {}).get('curr') if isinstance(summary.get('periods'),dict) else None
    available = None
    if isinstance(newest,dict):
        try:
            available = _dates(*newest['bounds'])
        except (ValueError, TypeError, KeyError):
            pass
    if available and requested and requested != available:
        note = ('Последний сохранённый свод дашборда охватывает '+available['from']+' — '+available['to']+
                '. Подтверждённой разбивки причин за запрошенный период нет.')
    elif available:
        note = ('Свод дашборда за указанный период сохранён без подтверждённой разбивки по фактам. '
                'Следующее штатное обновление дашборда сохранит её вместе со сводом.')
    else:
        note = 'Сохранённый свод дашборда с разбивкой по фактам недоступен.'
    return _response('unavailable','dashboard',city,note,period=requested or available,as_of=stamp)


def _map_result(city, requested):
    # Never prime/migrate a store or start a portal check when reading for chat.
    current = water_map.DATA_FILE
    path = current if current.is_file() else water_map.SEED_FILE
    data = _read(path)
    if not isinstance(data,dict) or not data:
        return _response('unavailable','map',city,'Сохранённую карту прочитать не удалось.')
    meta = data.get('meta',{})
    try:
        period=_dates(meta['from'],meta['to'])
        as_of=_stamp(meta.get('live_updated')) or _stamp(meta.get('updated'))
        if requested and (requested['from']<period['from'] or requested['to']>period['to']):
            return _response('unavailable','map',city,'Запрошенный период выходит за границы карты: '
                             +period['from']+' — '+period['to']+'.',period=requested,as_of=as_of)
        if requested:
            period=requested
        else:
            first=max(date.fromisoformat(period['from']),date.fromisoformat(period['to'])-timedelta(days=29))
            period={**period,'from':first.isoformat()}
        cols=data['cols']
        rows=data['rows']
        if not isinstance(cols,list) or not isinstance(rows,list) or len(rows)>water_map.MAX_RECORDS or meta.get('count',len(rows))!=len(rows):
            raise ValueError
        index={key:cols.index(key) for key in ('id','created','omsu','fact')}
        known = _known_cities()
        city_labels = [_place(name, known) for name in data['omsu']]
        sub_index=cols.index('subcat') if 'subcat' in cols else None
        counts=Counter()
        unknown_city=0
        seen=set()
        for row in rows:
            ident=row[index['id']]
            if type(ident) is not int or ident in seen:
                raise ValueError
            seen.add(ident)
            day=date.fromisoformat(row[index['created']][:10]).isoformat()
            if not period['from']<=day<=period['to']:
                continue
            place=_indexed(city_labels,row[index['omsu']])
            if place==_UNALLOCATED:
                unknown_city+=1
            if city and place!=city:
                continue
            fact=_indexed(data['fact'],row[index['fact']])
            subtopic=_indexed(data['subcat'],row[sub_index]) if sub_index is not None else ''
            counts[_label(fact,subtopic)] += 1
        note='Карта Добродела по дате подачи. Группы по факту ЕЦУР или подтеме; это указанные в обращениях проблемы.'
        if not requested:
            note+=' Период по умолчанию — последние 30 дней доступной карты.'
        if type(meta.get('omitted')) is int and meta['omitted']>0:
            note+=' Архив исключил записей: '+str(meta['omitted'])+'.'
        if city and not counts and unknown_city:
            return _response('unavailable','map',city,'В периоде есть обращения без известного ОМСУ. Отсутствие жалоб здесь не установлено.',period=period,as_of=as_of)
        return _response('ok','map',city,note,period=period,as_of=as_of,counts=counts)
    except (ValueError,TypeError,IndexError,KeyError,AttributeError):
        return _response('unavailable','map',city,'Структура сохранённой карты не позволяет подтвердить итог.')


def analyze(user, municipality=None, period=None, *, source='auto'):
    if not user or not user.get('is_active',True):
        raise HTTPException(401,'Требуется авторизация.')
    if not check_module_access(user,'mingkh'):
        raise HTTPException(403,'Нет доступа к жалобам МинЖКХ.')
    if source not in ('auto','dashboard','map'):
        raise ValueError('Неизвестный источник.')
    hint = 'dashboard' if source == 'dashboard' else 'map'
    city=None
    if municipality:
        city=_place(municipality,_known_cities())
        if city==_UNALLOCATED:
            return _response('clarify',hint,None,'Не удалось однозначно определить ОМСУ. Уточните его название.')
    try:
        requested=_dates(period['from'],period['to']) if period else None
    except (ValueError,TypeError,KeyError):
        return _response('clarify',hint,city,'Не удалось определить границы запрошенного периода.')
    primary=_dashboard_result(city,requested) if source!='map' else None
    if source=='dashboard':
        return primary or _dashboard_missing(city, requested)
    if requested and primary and primary['status']=='ok':
        return primary
    mapped=_map_result(city,requested)
    if not primary or primary['status'] != 'ok':
        return mapped
    if mapped['status'] == 'ok' and mapped['period']['to'] > primary['period']['to']:
        return mapped
    return primary
