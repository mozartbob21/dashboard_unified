"""Actual chat, SQLite and XLSX flow. No provider call or network access."""
from io import BytesIO
import json
import time
from unittest.mock import patch
from openpyxl import Workbook, load_workbook
from fastapi.testclient import TestClient
import pytest
from test.test_check_dynamics import users
from services.aichat.agents import ratings
from services.auth import security
from utils import db

def workbook(sheets):
    wb=Workbook()
    for index,(name,rows) in enumerate(sheets):
        ws=wb.active if not index else wb.create_sheet()
        ws.title=name
        for row in rows: ws.append(row)
    out=BytesIO(); wb.save(out); wb.close()
    return out.getvalue()

def samples():
    geo=workbook([('Справочник',[
        ['mkd_code','address_eds','point','address_yandex','address_yandex_components'],
        ['10','Чехов г, Мира ул, 1','37.5 55.15','Чехов, улица Мира, 1',''],
        ['11','Чехов г, Мира ул, 2','37.501 55.15','Чехов, улица Мира, 2','']])])
    complaints=workbook([('ЕДС',[
        ['id заявки','Источник','Адрес','КОД АИС','Дата создания'],
        ['-101','Инцидент','Чехов г, Мира ул, 1','10','2026-10-02'],
        ['-102','Инцидент','Чехов г, Мира ул, 2','11','2026-10-02']]),
        ('Тепловая карта ЕЦУР',[
        ['Дубль заявки ЕДС','Факт ГВС','Подтема','Источник','Адрес','Внутренний Id','created'],
        [None,None,'Отопление','ЕЦУР','Чехов, улица Мира, 1','101','2026-10-02']])])
    return {'geo_file':('geo.xlsx',geo), 'complaints_file':('complaints.xlsx',complaints)}

@pytest.fixture()
def client(users,monkeypatch):
    with patch('services.scheduler.start'):
        import app
    monkeypatch.setattr(app,'aichat_ask',lambda *a,**kw:pytest.fail('An agent calculation must not send workbooks to a model'))
    with TestClient(app.app) as c:
        c.cookies.set('access_token',security.create_access_token({'sub':'one'}))
        yield c

def start(c,key,files=None,**params):
    return c.post('/aichat/api/send',data={'agent_id':ratings.ID,'background':'true','request_id':key,**params},files=files)

def done(c,ticket):
    deadline=time.monotonic()+9
    while time.monotonic()<deadline:
        response=c.get(ticket['status_url'])
        assert response.status_code==200
        state=response.json()
        if state['status'] not in ('running','queued'): return state
        time.sleep(.015)
    pytest.fail('Rating did not finish')

def test_shared_rating_file_survives_restart_and_checks_grants(client):
    c=client
    assert c.get('/aichat/api/agents').json()['agents'][0]['id']==ratings.ID
    assert 'id="agentsBtn"' in c.get('/aichat').text
    files=samples();key='neighbors_rating_123456'
    r=start(c,key,files)
    assert r.status_code==202,r.text
    ticket=r.json()
    state=done(c,ticket)
    assert state['status']=='succeeded',state['result']
    result=state['result']
    assert 'скрытых: 1' in result['answer'] and 'К домам привязано: 2' in result['answer']
    assert result['required_modules']==['mingkh']
    assert start(c,key,files).json()['job_id']==ticket['job_id']
    assert start(c,key,files,agent_radius='50').status_code==409
    ident=result['artifact']['id']
    url='/aichat/api/agents/artifacts/'+ident
    db._SCHEMA_READY=False
    with db.get_db_connection() as conn:
        conn.execute('UPDATE users SET modules=? WHERE username=?',(json.dumps(['mingkh']),'two'))
    c.cookies.set('access_token',security.create_access_token({'sub':'two'}))
    history=c.get('/aichat/api/dialogs/'+ticket['dialog_id']).json()
    assert history['messages'][-1]['artifact']==result['artifact']
    file=c.get(url)
    assert file.status_code==200 and 'attachment' in file.headers['content-disposition']
    book=load_workbook(BytesIO(file.content),data_only=True)
    assert book.sheetnames==['Рейтинг','Только отопление','Не привязаны','Как считали']
    assert book['Рейтинг']['A3'].value is not None or book['Рейтинг']['A2'].value is not None
    book.close()
    with db.get_db_connection() as conn:
        conn.execute("UPDATE users SET modules='[]' WHERE username='two'")
    assert c.get('/aichat/api/agents').json()['agents']==[]
    assert 'id="agentsBtn"' not in c.get('/aichat').text
    assert c.get(url).status_code==403 and c.get(ticket['status_url']).status_code==403
    assert all('artifact' not in m for m in c.get('/aichat/api/dialogs/'+ticket['dialog_id']).json()['messages'])
    c.cookies.set('access_token',security.create_access_token({'sub':'one'}))
    assert c.delete('/aichat/api/dialogs/'+ticket['dialog_id']).status_code==200
    assert c.get(url).status_code==404
    with db.get_db_connection() as conn:
        assert not conn.execute('SELECT 1 FROM aichat_agent_artifacts').fetchone()

def test_denied_invalid_and_failed_agent_never_publishes_a_download(client):
    c=client
    files=samples()
    c.cookies.set('access_token',security.create_access_token({'sub':'three'}))
    assert start(c,'neighbors_denied_1234',files).status_code==403
    c.cookies.set('access_token',security.create_access_token({'sub':'one'}))
    assert start(c,'neighbors_empty_1234').status_code==422
    assert start(c,'neighbors_bbox_1234',files,agent_city='Химки').status_code==422
    assert start(c,'neighbors_radii_1234',files,agent_radius='500',agent_radius_wide='250').status_code==422
    files['complaints_file']=('broken.xlsx',workbook([('Лист', [['Без нужных колонок']])]))
    bad=start(c,'neighbors_badbook_1234',files)
    assert bad.status_code==202
    ticket=bad.json(); state=done(c,ticket)
    assert state['status']=='failed' and not state['result'].get('artifact')
    assert c.get('/aichat/api/agents/artifacts/'+ticket['job_id']).status_code==404
