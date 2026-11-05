"""Row search against actual local collector JSON layouts and report projections."""
import json

from fastapi import HTTPException
import pytest

from services import report_municipalities
from services.aichat import catalog_lookup as lookup, report_context


@pytest.fixture(autouse=True)
def isolated(tmp_path,monkeypatch):
    monkeypatch.setattr(report_context,'DATA_DIR',tmp_path)
    monkeypatch.setattr(report_municipalities,'REGISTRY',tmp_path/'registry.json')
    (tmp_path/'registry.json').write_text(json.dumps({'municipalities':{'Дубна':{},'Химки':{}}}))


def write(module,payload):
    target=report_context.DATA_DIR/report_context.MODULES[module][1]
    target.parent.mkdir(parents=True,exist_ok=True)
    target.write_text(json.dumps(payload))


def search(module,q,limit=5):
    return lookup.search({'id':1,'modules':[module]},module,q,limit=limit)


@pytest.mark.parametrize(('module','payload','query','fragment'),[
    ('edo', {'rows':[{'municipality':'Дубна','organization':'Котельная','status':'risk','empty_fields':['Период']}]},
     'Найди в ЭДО Котельная в Дубне','Котельная'),
    ('overdue',{'items':[{'municipality':'Дубна','overdue_count':0}]},'Покажи просроченные задачи в Дубне','Просроченные задачи 0'),
    ('watercontrol',{'rows':[{'municipality':'Дубна','task_id':'768','status':'risk'}]},'Какой статус задачи №768 в контроле воды','768'),
    ('utnkr',{'violators':[{'task_id':22,'municipality':'Дубна','name':'Школа','overdue_days':4,'situation':'private-annotation'}]},
     'Найди школу в УТНКР по номеру 22','Школа'),
    ('mgkh_rm',{'buckets':{'close':[{'id':55,'subject':'private original title','url':'https://internal.invalid/login'}],'extend':[],'rework':[]}},
     'ЖКХ Redmine номер 55','Закрыть'),
    ('cds',{'data':[{'number':'А-145','status':'В работе','date':'01.01.2026','address':'private residence','applicant':'private person'}]},
     'Обращение 1С №А-145','А-145'),
    ('cameras',{'items':[{'id':'5','municipality':'Химки','camera_status':'not_working','owner':'private owner'}]},
     'Какие камеры не работают в Химках','не работает'),
])
def test_supported_sources(module,payload,query,fragment):
    write(module,{'updated_at':'2026-11-04T16:00:00',**payload})
    result=search(module,query)
    assert result['status']=='ok' and result['total']==1, result
    assert result['scope']==module and result['as_of']=='2026-11-04T16:00:00'
    text=result['items'][0]['text']
    assert fragment in text and 'private' not in text and 'https://' not in text


def test_real_total_beyond_preview_and_exact_identifier_not_substring():
    write('watercontrol',{'rows':[{'municipality':'Химки','task_id':str(i),'status':'risk'} for i in range(20)]})
    assert search('watercontrol','Контроль воды в Химках',limit=2)['total']==20
    result=search('watercontrol','номер 1 в Химках',limit=1)
    assert result['total']==1 and 'Номер задачи 1' in result['items'][0]['text']


def test_no_false_zero_for_missing_unsupported_empty_failed_or_scratch():
    assert search('edo','ЭДО')['total'] is None
    write('edo',{'items':[None]})
    assert search('edo','ЭДО')['total'] is None
    write('edo',{'items':[{'secret':'private-annotation'}]})
    assert search('edo','ЭДО')['total'] is None
    write('edo',{'items':[]})
    assert search('edo','ЭДО')['total']==0
    write('edo',{'ok':False,'items':[{'municipality':'Дубна','organization':'Old'}]})
    assert search('edo','ЭДО')['total'] is None
    write('utnkr',{'source_mode':'local_rows','rows':[{'name':'Из debug'}]})
    assert search('utnkr','УТНКР')['total'] is None


def test_incompatible_period_and_missing_id_or_city_are_unavailable():
    write('cds',{'data':[{'status':'Открыто'}],'date_from':'01.01.2026','date_to':'01.01.2026'})
    assert search('cds','Обращения 1С за 02.01.2026')['total'] is None
    assert search('cds','Обращения 1С за 01.01.2026')['total']==1
    assert search('cds','Обращения 1С в Дубне')['total'] is None
    assert search('cds','Обращение 1С №123')['total'] is None


def test_sources_and_contacts_are_excluded_before_search_and_output():
    write('cameras',{'items':[{'id':3,'municipality':'Дубна','address':'Станция rtsp://private.host/stream',
                              'responsible':'private name','stream_url':'rtsp://password',
                              'phone':'8 (999) 123-45-67','token':'secret-key','raw_row_text':'special-narrative'}]})
    result=search('cameras','Камеры в Дубне')
    assert result['total']==1
    assert not any(s in repr(result) for s in ('rtsp:','private','secret-key','special-narrative','123-45-67'))
    assert search('cameras','Камеры special-narrative')['total']==0
    assert search('cameras','Телефоны ответственных')['status']=='clarify'


def test_grants_checked_before_opening_source(monkeypatch):
    def unsafe(*a,**kw):
        raise AssertionError('unauthorized source opened')
    monkeypatch.setattr(report_context,'_load_result',unsafe)
    for module in ('edo','collective','edds','unknown-file'):
        with pytest.raises(HTTPException) as error:
            lookup.search({'id':1,'role':'admin','modules':[]},module,'Записи')
        assert error.value.status_code==403
