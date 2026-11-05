"""Real SQLite/JSON schema with synthetic inputs. No sources or accounts are modified."""
import json
import sqlite3
import pytest
from fastapi import HTTPException

from services.aichat import other_complaint_causes as causes, remaining_sources
from services.collective.client import HEADERS
from test.test_aichat_treatment_lookup import snapshot
from utils import db

FACT='Устранить повреждение на сети централизованного водоотведения'
HIDDEN='SECRET private surname, download credentials http://example.test'
def user(*modules):
    return {'id':2,'role':'user','modules':list(modules)}

@pytest.fixture
def root(tmp_path,monkeypatch):
    monkeypatch.setattr(causes.remaining_sources,'DATA_DIR',tmp_path)
    monkeypatch.setattr(db,'DB_FILE',tmp_path/'db.sqlite')
    return tmp_path

def save(root,path,value):
    p=root/path
    p.parent.mkdir(exist_ok=True)
    p.write_text(json.dumps(value,ensure_ascii=False),encoding='utf-8')

def test_no_access_before_reading_storage_and_mingkh_does_not_grant_others(root,monkeypatch):
    monkeypatch.setattr(causes.taxonomy,'_known_cities',lambda:pytest.fail('unauthorized read'))
    for name in causes.SOURCES:
        actor=user('collective') if name=='treatment' else user('mingkh')
        with pytest.raises(HTTPException) as denied:
            causes.analyze(actor,name,'Чехов')
        assert denied.value.status_code==403

def test_ecur_verified_new_fact_and_old_category_with_limited_date(root):
    from datetime import date
    rows=[['Район','Категория ЕЦУР','Срок','Статус','ЕЦУР факт','Подкатегория ЕЦУР','Текст'],
          ['Чехов','Инженерная инфраструктура','2026-10-11','В работе',FACT,'—',HIDDEN],
          ['Чехов','Инженерная инфраструктура','2026-10-11','В работе',HIDDEN,'—',HIDDEN],
          ['Химки','Инженерная инфраструктура','2026-10-11','Закрыто',FACT,'—',HIDDEN]]
    source=remaining_sources.project_ecur(rows,today=date(2026,10,12))
    source['causes_v1']=causes.project_ecur_causes(rows)
    save(root,'ecur/report-summary.json',source)
    result=causes.analyze(user('ecur'),'ecur','Чехов')
    assert result['total']==2 and result['groups']==[{'label':FACT,'count':1}] and result['unknown']==1
    assert result['scope']=='ecur' and HIDDEN not in json.dumps(result)
    assert causes.analyze(user('ecur'),'ecur',period={'from':'2026-10-10','to':'2026-10-12'})['total'] is None
    source.pop('causes_v1')
    save(root,'ecur/report-summary.json',source)
    result=causes.analyze(user('ecur'),'ecur','Чехов')
    assert result['status']=='limited' and result['total']==2 and 'Детальные факты' in result['note']

def test_edds_reports_service_counts_without_claiming_incident_causes(root):
    save(root,'edds/water_daily.json',{'updated':'2026-10-11','days':{'2026-10-10':{'Чехов':[0,2,4]},'2026-10-11':{'Чехов':[1,1,0]}}})
    result=causes.analyze(user('edds'),'edds','Чехов')
    assert result['status']=='limited' and result['total']==8 and result['groups']==[]
    assert 'водоотведения' in result['note'].lower() and '— 3' in result['note']
    unavailable=causes.analyze(user('edds'),'edds','Чехов',{'from':'2026-10-09','to':'2026-10-11'})
    assert unavailable['total'] is None


def test_collective_readonly_sql_uses_verified_fact_or_allowed_theme(root):
    path=db.DB_FILE
    assert not path.exists()
    assert causes.analyze(user('collective'),'collective','Чехов')['total'] is None
    assert not path.exists()
    rows=[]
    for i,(city,day,fact) in enumerate([('Чехов','01.10.2026',FACT),('Дубна','01.10.2026',FACT),('Чехов','03.10.2026',HIDDEN)]):
        row=['']*len(HEADERS)
        for key,value in [('ОМСУ',city),('Дата поступления обращения',day),('Факт',fact),('Почта заявителя',HIDDEN)]:
            row[HEADERS.index(key)]=value
        rows.append(row)
    with sqlite3.connect(path) as conn:
        conn.execute('CREATE TABLE collective_snapshots(start,end,fetched_at,rows_json,keys_json)')
        conn.execute('CREATE TABLE collective_workflows(complaint_key,theme)')
        conn.execute('INSERT INTO collective_snapshots VALUES(?,?,?,?,?)',('2026-10-01','2026-10-04','2026-10-04T19:00:00+03:00',json.dumps(rows),json.dumps(['a','b','c'])))
        conn.execute("INSERT INTO collective_workflows VALUES('c','kr')")
    result=causes.analyze(user('collective'),'collective','Чехов')
    assert result['status']=='ok' and result['total']==2
    assert {x['label'] for x in result['groups']} == {FACT,'Капитальный ремонт МКД'}
    assert HIDDEN not in json.dumps(result)
    subset=causes.analyze(user('collective'),'collective','Чехов',{'from':'2026-10-01','to':'2026-10-01'})
    assert subset['groups']==[{'label':FACT,'count':1}] and subset['total']==1


def test_treatment_source_uses_facility_municipality_and_only_three_topics(snapshot):
    result=causes.analyze(user('mingkh'),'treatment','Дубна')
    assert result['scope']=='mingkh' and result['status']=='ok' and result['total']==1
    assert result['groups']==[{'label':'Запах от очистных','count':1}] and 'SECRET' not in json.dumps(result)
    with pytest.raises(HTTPException):
        causes.analyze(user('edds'),'treatment','Дубна')
    assert causes.analyze(user('mingkh'),'treatment',period={'from':'2026-01-01','to':'2026-11-04'})['total'] is None


def test_cds_scraper_shape_dates_and_private_addresses(root,monkeypatch):
    monkeypatch.setattr(causes.report_context,'DATA_DIR',root)
    data={'success':True,'timestamp':'2026-10-10T12:00:00','date_from':'01.10.2026','date_to':'10.10.2026',
          'count':3,'data':[{'date':'02.10.2026','address':'Московская область, г.о. Чехов, ул. Чехова',
                             'reason':FACT,'applicant':HIDDEN},
                            {'Дата':'05.10.2026','ОМСУ':'Чехов','Причина обращения':HIDDEN,'Тип обращения':'—'},
                            {'date':'02.10.2026','address':'г. Химки, ул. Чехова','reason':FACT}]}
    save(root,'cds/result.json',data)
    result=causes.analyze(user('cds'),'cds','Чехов')
    assert result['groups']==[{'label':FACT,'count':1}] and result['total']==2 and result['unknown']==1
    assert result['period']=={'from':'2026-10-01','to':'2026-10-10'} and result['as_of']=='2026-10-10T12:00:00'
    assert HIDDEN not in json.dumps(result) and 'ул. Чехова' not in str(result)
    subset=causes.analyze(user('cds'),'cds','Чехов',{'from':'2026-10-05','to':'2026-10-05'})
    assert subset['status']=='limited' and subset['total']==1 and subset['groups']==[]
    assert causes.analyze(user('cds'),'cds',period={'from':'2026-09-15','to':'2026-09-16'})['total'] is None
