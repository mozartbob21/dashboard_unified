"""Validated application operations preserving the imported calculation contracts."""
from datetime import timedelta
import re

from services.mingkh.client import PortalError
from . import context, svod_obrashcheniya as svod, rez_data, rm_data
from .filters import SOURCES, SUBTOPICS

HEAVY = {'build', 'daily', 'rez/data', 'sys/run', 'rm/measures'}
PRESETS = [
    {'id':'fkr','title':'Жалобы по ФКР','hint':'Исполнитель — ФКР, все тематики','subtopics':'all','dept':svod.DEPT_FKR},
    {'id':'water','title':'Жалобы по ВС/ВО','hint':'Четыре подтемы водоснабжения','subtopics':svod.WATER,'dept':'-1'},
    {'id':'ministry','title':'Жалобы по министерству','hint':'Все тематики','subtopics':'all','dept':'-1'},
    {'id':'kapremont','title':'Жалобы по капремонту','hint':'Фонд капитального ремонта','subtopics':['548'],'dept':'-1'},
]


def options():
    end = svod.last_wednesday(context.today())
    from . import sys_data
    registry_dates = sys_data.registries()
    system_date = end
    # An imported snapshot is the baseline, not a previous-period result. Avoid
    # selecting that same day when there is no older snapshot to compare with.
    if not any(day < end for day in registry_dates) and any(day < context.today() for day in registry_dates):
        system_date = context.today()
    warning = ''
    try:
        omsu = svod.omsu_list()
    except Exception:
        omsu = []
        warning = 'Справочник ОМСУ пока недоступен. Проверьте доступ к источнику или перенесите данные из архива. Доступен расчёт по всем муниципалитетам.'
    return {'date':end.isoformat(),'system_date':system_date.isoformat(),'week':{'from':(end-timedelta(days=6)).isoformat(),'to':end.isoformat()},
            'sources':[{'id':str(i),'title':t} for i,t in SOURCES if i!='-1'],
            'subtopics':[{'id':str(i),'title':t} for i,t in SUBTOPICS if i!='-1' and t not in svod.EXCLUDED_TITLES],
            'omsu':[{'id':str(i),'title':t} for i,t in omsu],'presets':PRESETS,'warning':warning}


def validate(action, req):
    if not isinstance(action, str) or action not in HEAVY or not isinstance(req, dict):
        raise ValueError('Неизвестный расчёт.')
    result = {}
    if action in {'build', 'daily', 'sys/run'}:
        end = context.iso_date(req.get('date'))
        start = req.get('from')
        if start is None or start == '':
            start = svod.year_start(end).isoformat() if action=='build' else end.isoformat()
        start, end = context.date_range(start, end, max_days=1096 if action=='build' else 366)
        result.update({'from':start.isoformat(),'date':end.isoformat()})
    else:
        start, end = context.date_range(req.get('from'), req.get('to'), max_days=366)
        result.update({'from':start.isoformat(),'to':end.isoformat()})
    if action=='rez/data':
        a,b=context.date_range(req.get('prev_from'),req.get('prev_to'),max_days=366)
        if (max(end, b) - min(start, a)).days > 1096:
            raise ValueError('Периоды сравнения расположены слишком далеко друг от друга.')
        result.update(prev_from=a.isoformat(),prev_to=b.isoformat())
    if action in {'build','rez/data'}:
        for key, choices in (('sources',SOURCES),('subtopics',SUBTOPICS)):
            allowed={str(i) for i,title in choices if str(i)!='-1' and (key!='subtopics' or title not in svod.EXCLUDED_TITLES)}
            values=req.get(key)
            if not isinstance(values,list) or not values or len(values)>len(allowed) or any(type(v) is not str or v not in allowed for v in values):
                raise ValueError('Выберите корректные источники и тематики обращений.')
            result[key]=list(dict.fromkeys(values))
        omsu=req.get('omsu', '-1'); dept=req.get('dept', '-1')
        if type(omsu) is not str or not re.fullmatch(r'-1|[0-9]{1,10}',omsu):
            raise ValueError('Некорректный муниципалитет.')
        if type(dept) is not str or dept not in ('-1',svod.DEPT_FKR):
            raise ValueError('Неизвестный исполнитель.')
        result.update(omsu=omsu,dept=dept)
    return result


def execute(action, req):
    req=validate(action,req)
    if action in {'build', 'rez/data'} and req['omsu'] != '-1':
        if req['omsu'] not in dict(svod.omsu_list()):
            raise ValueError('Выбранный муниципалитет отсутствует в справочнике портала.')
    if action=='build':
        src=req['sources']; all_src={str(i) for i,_ in SOURCES if str(i)!='-1'}
        d=svod.build(context.iso_date(req['from']),context.iso_date(req['date']),[] if set(src)==all_src else src,req['omsu'],req['dept'],log=lambda _:None)
        return {'text':svod.render(d,req['subtopics']),'portal_updated':d['portal_updated']}
    if action=='daily':
        d=svod.daily(context.iso_date(req['from']),context.iso_date(req['date']))
        return {'text':svod.render_daily(d),'date_from':d['date_from'],'date':d['date'],
                'prev_from':d['prev_from'],'prev':d['prev'],'rows':d['rows'],
                'portal_updated':svod.Pentaho().last_updated()[0]}
    if action=='rez/data':return rez_data.collect(req,log=lambda _:None)
    if action=='rm/measures':return rm_data.measures(req['from'],req['to'])
    if action=='sys/run':
        from . import sys_data
        return sys_data.run(req,log=lambda _:None)
