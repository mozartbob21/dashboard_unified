"""Read complaint classifications under the grant of the chosen source."""
from collections import Counter
from contextlib import closing
from datetime import date
import sqlite3
import re
import zlib

from fastapi import HTTPException
from services.aichat import complaint_causes as taxonomy, catalog_lookup, local_sources, remaining_sources, report_context
from services.aichat.treatment_lookup import _linked
from services.collective.client import HEADERS
from services.collective.store import MAX_ROWS, MAX_SOURCE_BYTES
from services.mingkh import treatment, water_map
from core.roles import check_module_access
from utils import db

SOURCES = {
    'cds': ('cds','Обращения 1С','/cds'),
    'collective': ('collective','Коллективные жалобы МинЖКХ','/mingkh/collective'),
    'ecur': ('ecur','ЕЦУР — активные жалобы','/ecur'),
    'edds': ('edds','ЕДДС — жалобы Добродела по воде','/edds'),
    'treatment': ('mingkh','Жалобы по очистным сооружениям','/mingkh/treatment'),
}
THEMES = {'ii':'Инженерная инфраструктура','kr':'Капитальный ремонт МКД','pr':'Прочее'}

def _result(module, city, status, note, *, counts=None, total=None, period=None, as_of=None):
    counts = counts or {}
    scope,title,url = SOURCES[module]
    groups = sorted(((label,count) for label,count in counts.items() if count and label != taxonomy._UNASSIGNED),
                    key=lambda item: (-item[1],item[0]))
    return {'scope':scope, 'title':title,'url':url,'status':status,'municipality':city,
            'note':note, 'as_of':taxonomy._stamp(as_of), 'period':period, 'total':total,
            'groups':[{'label':label,'count':count} for label,count in groups[:taxonomy.MAX_GROUPS]],
            'other':sum(count for _,count in groups[taxonomy.MAX_GROUPS:]),'unknown':counts.get(taxonomy._UNASSIGNED,0) if total is not None else None}

def project_ecur_causes(rows):
    """Optional extension of the ECUR aggregate; a legacy/incomplete grid stays readable."""
    if not isinstance(rows,list) or not 1<=len(rows)<=remaining_sources.MAX_ROWS+1:
        return {}
    fields=('Район','ЕЦУР факт','Подкатегория ЕЦУР')
    header=rows[0]
    if not isinstance(header,list) or any(header.count(key)!=1 for key in fields):
        return {}
    idx, fact_idx, sub_idx = map(header.index,fields)
    counts = {}
    known = taxonomy._known_cities()
    groups=0
    try:
        for row in rows[1:]:
            if not isinstance(row,list):
                return {}
            city=taxonomy._place(row[idx],known)
            label=taxonomy._label(row[fact_idx],row[sub_idx])
            grouped=counts.setdefault(city,Counter())
            if label not in grouped:
                groups+=1
                if groups>taxonomy.MAX_PERIOD_BUCKETS:
                    return {}
            grouped[label]+=1
    except (IndexError,ValueError,TypeError):
        return {}
    return {'by_city':counts,'total':len(rows)-1}


def _ecur(city, requested):
    data=remaining_sources._read(remaining_sources.DATA_DIR/'ecur/report-summary.json')
    if data.get('schema_version')!=1:
        return _result('ecur',city,'unavailable','Сохранённая выборка активных жалоб ЕЦУР отсутствует.')
    day=data['as_of']
    bounds=taxonomy._dates(day,day)
    stamp=data.get('collected_at')
    if requested and requested!=bounds:
        return _result('ecur',city,'unavailable','Есть срез активных жалоб на '+day+
                       '; исторические сведения за запрошенный период он не подтверждает.',period=requested,as_of=stamp)
    by_city=data['municipalities']
    totals=data['totals']
    if not isinstance(by_city,dict) or type(totals.get('total')) is not int:
        raise ValueError
    group = by_city.get(taxonomy.cities.key(city)) if city else totals
    if city and group is None and data.get('unknown_municipality_rows'):
        return _result('ecur',city,'unavailable','У части активных жалоб не определён район.',period=bounds,as_of=stamp)
    if city and group is None:
        return _result('ecur',city,'unavailable','В выборке активных жалоб не подтверждена отдельная строка ОМСУ.',period=bounds,as_of=stamp)
    total=group.get('total')
    if type(total) is not int or total<0:
        raise ValueError
    saved=data.get('causes_v1') or {}
    counts=Counter()
    if saved.get('total')==totals['total'] and isinstance(saved.get('by_city'),dict):
        for name,buckets in saved['by_city'].items():
            if not city or name==city:
                for label,n in buckets.items():
                    if type(n) is not int or n<0:
                        raise ValueError
                    counts[taxonomy._label(label,None)]+=n
        if sum(counts.values())!=total:
            raise ValueError
        return _result('ecur',city,'ok','Активные задачи ЕЦУР в сохранённом срезе. Разбивка по фактам и подтемам на дату проверки.',
                       counts=counts,total=total,period=bounds,as_of=stamp)
    # Older snapshots hold aggregate categories, not source facts.
    allowed = {taxonomy._plain(name):name for name in taxonomy._read(water_map.SEED_FILE).get('cat',[]) if isinstance(name,str)}
    for label,n in (group or {}).get('category_counts',{}).items():
        if type(n) is not int or n<0:
            raise ValueError
        counts[allowed.get(taxonomy._plain(label),taxonomy._UNASSIGNED)]+=n
    if sum(counts.values())>total:
        raise ValueError
    counts[taxonomy._UNASSIGNED]+=total-sum(counts.values())
    return _result('ecur',city,'limited','Сохранены категории активных жалоб. Детальные факты по ним в этом срезе отсутствуют.',
                   counts=counts,total=total,period=bounds,as_of=stamp)


def _edds(city, requested):
    data=taxonomy._read(remaining_sources.DATA_DIR/'edds/water_daily.json')
    projection=local_sources._daily(data,city,requested)
    metrics=projection.get('metrics',[])
    bounds=requested or projection.get('period')
    stamp=projection.get('collected_at')
    if not metrics or any(type(m['value']) is not int for m in metrics):
        return _result('edds',city,'unavailable','За выбранный период или ОМСУ нет полного подтверждённого свода жалоб ЕДДС.',
                       period=bounds,as_of=stamp)
    total=sum(m['value'] for m in metrics)
    parts='; '.join(m['label']+' — '+str(m['value']) for m in metrics)
    note='Свод классифицирует обращения только по виду услуги: '+parts+'. Причины по отдельным жалобам в ЕДДС не сохраняются.'
    if projection.get('period',{}).get('days_present')!=projection.get('period',{}).get('all_saved_days') and city:
        note+=' Сумма только за дни с записями по выбранному ОМСУ.'
    return _result('edds',city,'limited',note,counts={taxonomy._UNASSIGNED:total},
                   total=total,period=bounds,as_of=stamp)


def _treatment(city, requested):
    snap=treatment.overview(False)  # excludes the independently protected EDDS events
    if not snap:
        return _result('treatment',city,'unavailable','Снимок жалоб по очистным недоступен.')
    bounds=taxonomy._dates(snap['from'],snap['as_of'])
    period=requested or bounds
    if not bounds['from']<=period['from']<=period['to']<=bounds['to']:
        return _result('treatment',city,'unavailable','Снимок охватывает '+bounds['from']+' — '+bounds['to']+
                       '. Запрошенный период выходит за эти границы.',period=period,as_of=snap['as_of'])
    objects={row['n']:row for row in snap['kos']}
    counts=Counter()
    known=taxonomy._known_cities()
    unknown_place=0
    for row in snap['complaints']:
        arrived=date.fromisoformat(row['created'][:10]).isoformat()
        if not period['from']<=arrived<=period['to']:
            continue
        obj=_linked(row,objects,snap['radii'])
        place=taxonomy._place((obj or row).get('omsu'),known)
        unknown_place+=place==taxonomy._UNALLOCATED
        if city and city!=place:
            continue
        counts[treatment.TOPICS.get(row.get('topic'),taxonomy._UNASSIGNED)]+=1
    if city and not counts and unknown_place:
        return _result('treatment',city,'unavailable','У части жалоб неизвестен ОМСУ.',period=period,as_of=snap['as_of'])
    return _result('treatment',city,'ok','Три темы снимка по дате поступления. Для связанной жалобы ОМСУ определяется по объекту КОС, для остальных — по самой жалобе.',
                   counts=counts,total=sum(counts.values()),period=period,as_of=snap['as_of'])


def _collective(city, requested):
    if not db.DB_FILE.is_file():
        return _result('collective',city,'unavailable','Сохранённая выборка коллективных жалоб недоступна.',period=requested)
    with closing(sqlite3.connect(db.DB_FILE.resolve().as_uri()+'?mode=ro',uri=True)) as conn:
        conn.row_factory=sqlite3.Row
        conn.execute('BEGIN')
        where=' WHERE start<=? AND end>=?' if requested else ''
        order=' ORDER BY (start=? AND end=?) DESC, fetched_at DESC' if requested else ' ORDER BY fetched_at DESC'
        values=[requested['from'],requested['to']]*2 if requested else []
        meta=conn.execute('SELECT rowid AS snapshot_id,start,end,fetched_at,length(rows_json) AS bytes,'
                          'json_array_length(rows_json) AS rows FROM collective_snapshots'+where+order+' LIMIT 1',values).fetchone()
        if not meta:
            return _result('collective',city,'unavailable','В базе нет подходящей сохранённой выборки коллективных жалоб.',period=requested)
        if not 0<=meta['bytes']<=MAX_SOURCE_BYTES or not 0<=meta['rows']<=MAX_ROWS:
            raise ValueError
        bounds=taxonomy._dates(meta['start'],meta['end'])
        period=requested or bounds
        idx={k:HEADERS.index(k) for k in ('Факт','Дата поступления обращения','ОМСУ')}
        # Project only the three operational cells plus the optional reviewed theme.
        query='''SELECT json_extract(r.value,?) AS fact,json_extract(r.value,?) AS received,
                    json_extract(r.value,?) AS omsu,w.theme AS manual
                 FROM collective_snapshots AS s,json_each(s.rows_json) AS r
                 LEFT JOIN collective_workflows w ON w.complaint_key=json_extract(s.keys_json, '$['||r.key||']')
                 WHERE s.rowid=?'''
        counts=Counter()
        known=taxonomy._known_cities()
        missing_days=missing_city=0
        for row in conn.execute(query, ['$['+str(idx[k])+']' for k in ('Факт','Дата поступления обращения','ОМСУ')]+[meta['snapshot_id']]):
            place=taxonomy._place(row['omsu'],known)
            if city and place not in (city,taxonomy._UNALLOCATED):
                continue
            if period!=bounds:
                day=remaining_sources._date(row['received'])
                if not day:
                    missing_days+=1
                    continue
                if not period['from']<=day.isoformat()<=period['to']:
                    continue
            if city and place==taxonomy._UNALLOCATED:
                missing_city+=1
                continue
            fact=taxonomy._label(row['fact'],None)
            if fact==taxonomy._UNASSIGNED:
                fact=THEMES.get(row['manual'],taxonomy._UNASSIGNED)
            counts[fact]+=1
        if missing_days or (city and not counts and missing_city):
            return _result('collective',city,'unavailable','У части строк не подтверждены даты или ОМСУ; точный итог неизвестен.',period=period,as_of=meta['fetched_at'])
    return _result('collective',city,'ok','Факты из сохранённого реестра коллективных обращений. Если факт не классифицирован, учтена сохранённая ручная тема; она могла быть уточнена позднее.',
                   counts=counts,total=sum(counts.values()),period=period,as_of=meta['fetched_at'])


_ADDRESS_CITY = re.compile(
    r'^(?:городской округ|муниципальный округ|город|г\s*\.\s*о\s*\.|г\.?)\s+'
    r'([а-яё][а-яё0-9 .-]{0,70})$', re.I)


def _cds_place(row, known):
    explicit = report_context._municipality(row)
    if explicit:
        return taxonomy._place(explicit, known)
    # Only explicit administrative components, never street or person names.
    address = row.get('address') or row.get('Адрес обращения') or ''
    if not isinstance(address,str):
        return taxonomy._UNALLOCATED
    places = {taxonomy._place(match[1], known) for component in address.split(',')
              if (match := _ADDRESS_CITY.fullmatch(component.strip()))}
    places.discard(taxonomy._UNALLOCATED)
    return places.pop() if len(places)==1 else taxonomy._UNALLOCATED


def _cds(city, requested):
    data,error = report_context._load_result('cds')
    if data is None:
        return _result('cds',city,'unavailable','Сохранённая выгрузка обращений 1С недоступна.',period=requested)
    rows = catalog_lookup._shape(data)
    if rows is None or len(rows)>300_000 or any(not isinstance(row,dict) for row in rows):
        raise ValueError('Недостоверная структура результата 1С.')
    report = report_context._module_report('cds','',loaded=(data,error))
    stamp = catalog_lookup._date(report.get('collected_at'))
    if not stamp and isinstance(data,dict):
        stamp = catalog_lookup._date(data.get('timestamp'))
    if report.get('status')!='saved' or (isinstance(data,dict) and 'count' in data and data['count']!=len(rows)):
        return _result('cds',city,'unavailable','Выгрузка 1С не подтвердила полный результат проверки.',as_of=stamp,period=requested)
    bounds = None
    available = report.get('period')
    if available:
        first=remaining_sources._date(available.get('from'))
        last=remaining_sources._date(available.get('to'))
        if first and last:
            bounds=taxonomy._dates(first.isoformat(),last.isoformat())
    if requested and (not bounds or not bounds['from']<=requested['from']<=requested['to']<=bounds['to']):
        return _result('cds',city,'unavailable','Сохранённая выгрузка 1С не подтверждает весь запрошенный период.',
                       period=requested,as_of=stamp)
    counts=Counter()
    known=taxonomy._known_cities() if city else {}
    unallocated = missing_days = 0
    for row in rows:
        if requested and requested!=bounds:
            received = remaining_sources._date(row.get('date') or row.get('Дата'))
            if not received:
                missing_days+=1
                continue
            if not requested['from']<=received.isoformat()<=requested['to']:
                continue
        if city:
            place=_cds_place(row,known)
            unallocated+=place==taxonomy._UNALLOCATED
            if place!=city:
                continue
        # Both the normalised scraper and older Excel-column exports occur on disk.
        reason = row.get('reason') or row.get('Причина обращения')
        counts[taxonomy._label(reason, row.get('type') or row.get('Тип обращения'))]+=1
    period=requested or bounds
    if missing_days or (city and not counts and unallocated):
        return _result('cds',city,'unavailable','В выгрузке 1С есть строки без подтверждённой даты или ОМСУ; '
                       'итог по запрошенному срезу не установлен.',period=period,as_of=stamp)
    total=sum(counts.values())
    classified=total-counts.get(taxonomy._UNASSIGNED,0)
    note = ('Использованы распознанные причины и типы обращений из сохранённой выгрузки 1С.'
            if classified else 'В этой выборке 1С нет распознанной классификации причин обращений.')
    if unallocated:
        note+=' Показаны только подтверждённые строки ОМСУ; часть адресов в выборке сопоставить не удалось.'
    if not period:
        note+=' Границы периода выгрузки не указаны.'
    return _result('cds',city,'ok' if classified else 'limited',note,counts=counts,total=total,period=period,as_of=stamp)


def analyze(user, module_id, municipality=None, period=None):
    if module_id not in SOURCES:
        raise ValueError('Неизвестный раздел жалоб.')
    if not user or not user.get('is_active',True):
        raise HTTPException(401,'Требуется авторизация.')
    if not check_module_access(user, SOURCES[module_id][0]):
        raise HTTPException(403,'Нет доступа к этому разделу жалоб.')
    city=taxonomy._place(municipality,taxonomy._known_cities()) if municipality else None
    if city==taxonomy._UNALLOCATED:
        return _result(module_id,None,'clarify','Не удалось однозначно определить ОМСУ.')
    try:
        requested=taxonomy._dates(period['from'],period['to']) if period else None
    except (ValueError,TypeError,KeyError):
        return _result(module_id,city,'clarify','Уточните границы периода.')
    readers={'cds':_cds,'collective':_collective,'treatment':_treatment,'ecur':_ecur,'edds':_edds}
    try:
        return readers[module_id](city,requested)
    except (ValueError,TypeError,KeyError,IndexError,AttributeError,OSError,sqlite3.Error,EOFError,zlib.error):
        return _result(module_id,city,'unavailable','Сохранённая выборка недоступна или имеет неподтверждённую структуру.',period=requested)
