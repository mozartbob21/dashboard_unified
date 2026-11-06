"""Normal browser transport: signed account cookie, durable send and polling, isolated database."""
import gzip
import json
import time
from datetime import date
from unittest.mock import patch
from uuid import uuid4

from fastapi.testclient import TestClient
import pytest

from services.aichat import complaint_causes, jobs, private_storage, report_context
from services.auth import security
from services.mingkh import water_map
from utils import db

QUESTION = 'что там по чехову по причинам жалоб'
FACT = 'Устранить повреждение на сети централизованного водоотведения'
SECOND = 'Восстановить работу внешней системы водоснабжения'
SECRET = 'НЕ_ПОКАЗЫВАТЬ_АДРЕС_И_ТЕЛЕФОН'
ORDINARY = 'Обычный ответ без чтения данных платформы.'


@pytest.fixture()
def agent_client(tmp_path, monkeypatch):
    monkeypatch.setattr(db,'DB_FILE',tmp_path/'agent.sqlite3')
    monkeypatch.setattr(db,'_SCHEMA_READY',False)
    monkeypatch.setattr(private_storage,'LEGACY_FILE',tmp_path/'missing.json')
    private_storage.initialize()
    with db.get_db_connection() as conn:
        for username,modules in [('manager',[]),('analyst',['mingkh']),('visitor',[])]:
            conn.execute('INSERT INTO users(username,password_hash,role,modules) VALUES (?,?,?,?)',
                         (username,'synthetic hash','user',json.dumps(modules)))
    monkeypatch.setattr(security,'get_jwt_secret',lambda:'synthetic-secret-for-chat-integration-only')
    monkeypatch.setattr(complaint_causes,'_REPORT_PATH',tmp_path/'no-report.json')
    monkeypatch.setattr(report_context,'DATA_DIR',tmp_path)
    monkeypatch.setattr(water_map,'SEED_FILE',tmp_path/'seed.json.gz')
    monkeypatch.setattr(water_map,'DATA_FILE',tmp_path/'current.json.gz')
    complaint_causes._taxonomy.cache_clear()
    points={}
    for ident,city,fact in [(1,'Чехов',FACT),(2,'Чехов',FACT),(3,'Химки',SECOND)]:
        points[str(ident)]={'id':ident,'created':'2026-10-19','kind':'hvs','omsu':city,
                            'fact':fact,'address':SECRET}
    dataset=water_map.compact(points,{'updated':'19.10.2026 19:00'})
    water_map.SEED_FILE.write_bytes(gzip.compress(json.dumps(dataset,ensure_ascii=False).encode('utf-8')))
    with patch('services.scheduler.start'):
        import app
    # Direct model and skills transport stand-ins allow ordinary chat to work
    # offline. The requested analytical result must instead come from source rows.
    monkeypatch.setattr(app,'aichat_ask',lambda *args, **kwargs: ORDINARY)
    from services.aichat.skills import runtime
    monkeypatch.setattr(runtime,'_model',lambda *args, **kwargs:'{}')
    with TestClient(app.app) as client:
        client.cookies.set('access_token',security.create_access_token({'sub':'analyst'}))
        yield client
    complaint_causes._taxonomy.cache_clear()


def _ask(client, text, *, dialog_id=None, include_context=None):
    form={'text':text,'request_id':uuid4().hex,'background':'true','skills_mode':'auto'}
    if dialog_id: form['dialog_id']=dialog_id
    if include_context is not None: form['include_context']=include_context
    response=client.post('/aichat/api/send',data=form)
    assert response.status_code==202, response.text
    ticket=response.json()
    deadline=time.monotonic()+5
    while time.monotonic()<deadline:
        poll=client.get(ticket['status_url'])
        assert poll.status_code==200,poll.text
        state=poll.json()
        if state['status'] not in jobs.ACTIVE:
            assert state['status']=='succeeded',state
            return ticket,state['result']
        time.sleep(0.01)
    pytest.fail('The local read-only chat job did not finish.')


def test_exact_user_phrase_and_followup_read_same_source_with_current_grants(agent_client, monkeypatch):
    client=agent_client
    from services.aichat import analytical_tools, live_mingkh_tool, report_periods
    from services.mingkh.sms.client import Pentaho
    monkeypatch.setattr(report_periods, 'today', lambda: date(2026,11,6))
    monkeypatch.setattr(analytical_tools, '_snapshot', lambda *args: pytest.fail('An archive replaced the live source'))
    calls=[]

    class Portal:
        periods=staticmethod(Pentaho.periods)

        def last_updated(self):
            return '2026-11-05 12:00', None

        def query(self, name, params=None, **kwargs):
            calls.append((name, dict(params or {})))
            if name=='q_omsu':
                return ['id','title'], [['123','Чехов'],['456','Химки']]
            assert name=='q_download_detalization'
            assert params['curr_period_start']==params['curr_period_end']=='18 day'
            assert params['omsu'] in ('123','456')
            city, topic, count=('Чехов','Канализация',2) if params['omsu']=='123' else ('Химки','Отсутствие воды',1)
            rows=[[ident,city,'Добродел','Водоснабжение/Водоотведение',
                   'Водоотведение' if city=='Чехов' else 'Качество воды',topic,SECRET,SECRET,'2026-10-19']
                  for ident in range(1,count+1)]
            return ['Внутренний Id']+list(live_mingkh_tool.dashboard.COLS.values()),rows

    monkeypatch.setattr(live_mingkh_tool, 'Pentaho', Portal)
    ticket,result=_ask(client,QUESTION)
    assert result['response_kind']=='clarification' and result['required_modules']==['mingkh']
    assert result['agent_pending']['municipality']=='Чехов'
    assert result['agent_pending']['tools']==['query:mingkh-complaints']
    assert result['agent_pending']['periods']==[]
    assert 'период' in result['answer'].lower() and 'action' not in result and not calls

    dated_ticket, dated=_ask(client,'за 19.10.2026',dialog_id=ticket['dialog_id'])
    answer=dated['answer']
    assert dated['report_scope']['agent_action']=='execution_plan' and dated['required_modules']==['mingkh']
    first=dated['agent_plan']['steps'][0]
    assert first['tool']=='query:mingkh-complaints'
    assert first['params']=={'start':'2026-10-19','end':'2026-10-19','municipality':'Чехов','topic':'all'}
    assert first['result']['metrics'][0]['value']==2
    assert first['result']['topic_breakdown']==[{'label':'Канализация','count':2}]
    assert '/mingkh' in answer and '2026-10-19' in answer
    assert 'Тематика обращений по классификации источника' in answer and 'Канализация — 2.' in answer
    assert ORDINARY not in answer and SECRET not in answer
    assert SECRET not in json.dumps(dated,ensure_ascii=False)
    saved = client.get('/aichat/api/dialogs/'+ticket['dialog_id']).json()
    assert saved['messages'][-1]['report_sources'][0]['source_url']=='/mingkh'
    next_ticket,follow=_ask(client,'А по Химкам?',dialog_id=ticket['dialog_id'])
    assert follow['report_scope']['agent_action']=='execution_plan' and follow['required_modules']==['mingkh']
    second=follow['agent_plan']['steps'][0]
    assert second['tool']=='query:mingkh-complaints'
    assert second['params']=={**first['params'],'municipality':'Химки'}
    assert second['result']['metrics'][0]['value']==1
    assert second['result']['topic_breakdown']==[{'label':'Отсутствие воды','count':1}]
    assert 'Отсутствие воды — 1.' in follow['answer'] and 'Канализация' not in follow['answer']
    assert SECRET not in json.dumps(follow,ensure_ascii=False)
    assert [params['omsu'] for name,params in calls if name=='q_download_detalization']==['123','456']

    with db.get_db_connection() as conn:
        conn.execute("UPDATE users SET modules='[]' WHERE username='analyst'")
    for token in (ticket,dated_ticket,next_ticket):
        assert client.get(token['status_url']).status_code==403
    dialog=client.get('/aichat/api/dialogs/'+ticket['dialog_id'])
    assert dialog.status_code==200 and '/mingkh' not in dialog.text
    assert 'Канализация' not in dialog.text and 'Отсутствие воды' not in dialog.text
    assert all('agent_plan' not in message and 'agent_pending' not in message for message in dialog.json()['messages'])
    count_before=len(calls)
    _,denied=_ask(client,'Посмотри причины жалоб в МинЖКХ по Чехову')
    assert '/mingkh' not in denied['answer'] and SECRET not in denied['answer']
    assert len(calls)==count_before


def test_explicit_context_off_skips_all_source_reads(agent_client,monkeypatch):
    def forbidden(*args, **kwargs):
        pytest.fail('A request with platform context off read the complaint archive')
    monkeypatch.setattr(complaint_causes,'analyze',forbidden)
    _,result=_ask(agent_client,QUESTION,include_context='false')
    assert result['answer']==ORDINARY and result['required_modules']==[]


def test_cds_access_and_partial_classification_through_normal_agent(agent_client,tmp_path):
    path=tmp_path/'cds'/'result.json'
    path.parent.mkdir()
    path.write_text(json.dumps({'timestamp':'2026-10-10T12:00:00','date_from':'01.10.2026',
        'date_to':'10.10.2026','success':True,'count':1,
        'data':[{'address':'г. Чехов, ул. '+SECRET,'reason':SECRET,'applicant':SECRET,'date':'10.10.2026'}]}))
    question='По каким причинам жалуются в Чехове по обращениям 1С?'
    with db.get_db_connection() as conn:
        conn.execute("UPDATE users SET modules='[]' WHERE username='analyst'")
    _,denied=_ask(agent_client,question)
    assert '/cds' not in denied['answer'] and SECRET not in denied['answer']
    with db.get_db_connection() as conn:
        conn.execute('UPDATE users SET modules=? WHERE username=?',(json.dumps(['cds']),'analyst'))
    _,limited=_ask(agent_client,question)
    assert limited['report_scope']['agent_action']=='complaint_causes'
    assert limited['required_modules']==['cds'] and '/cds' in limited['answer']
    assert 'Количество обращений: 1' in limited['answer'] and 'нет распознанной классификации' in limited['answer']
    assert SECRET not in limited['answer']
