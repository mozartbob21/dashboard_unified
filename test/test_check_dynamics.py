"""Actual HTTP/chat/SQLite path with local accounts. No model or external collectors."""
import json
from unittest.mock import patch
from fastapi.testclient import TestClient
import pytest
from services import run_history, check_history
from services.aichat import private_storage
from services.aichat.check_dynamics import resolve
from utils import db

@pytest.fixture()
def users(tmp_path,monkeypatch):
    monkeypatch.setattr(db,'DB_FILE',tmp_path/'history.db')
    monkeypatch.setattr(db,'_SCHEMA_READY',False)
    monkeypatch.setattr(private_storage,'LEGACY_FILE',tmp_path/'absent.json')
    private_storage.initialize()
    from services.auth import security
    monkeypatch.setattr(security,'get_jwt_secret',lambda:'synthetic-chat-history-non-production-secret')
    with db.get_db_connection() as c:
        for name,modules in [('one',['mingkh','zips','edo']),('two',['zips']),('three',[])]:
            c.execute("INSERT INTO users(username,password_hash,role,modules) VALUES (?,?,?,?)",(name,'unused','user',json.dumps(modules)))
    return [security.find_user_by_username(name) for name in ('one','two','three')]

def insert(module,metric,n,op=None,owner=None):
    ident=run_history.record_start(module,params={'operation':op} if op else {},private_owner=owner)
    run_history.record_finish(ident,metrics={metric:n})
    return ident

def test_shared_dialog_compares_actual_public_runs_while_api_filters_private_and_revoked_rights(users,monkeypatch):
    from services.auth import security
    with patch('services.scheduler.start'):
        import app
    monkeypatch.setattr(app,'aichat_ask',lambda *a,**k: pytest.fail('A deterministic history answer needs no model'))
    cid=users[0]['id']
    insert('zips','files',12)
    insert('zips','files',0)
    insert('zips','files',8799,owner='local:'+str(cid))
    with TestClient(app.app,raise_server_exceptions=True) as client:
        def login(user):
            client.cookies.set('access_token',security.create_access_token({'sub':user['username']}))
        login(users[0])
        listing=client.get('/api/check-history/zips',params={'status':'success','limit':2}).json()['checks']
        assert len(listing)==2 and listing[0]['metrics']['files']==8799
        r=client.post('/aichat/api/send',data={'text':'Сравни последние проверки ЗиП'})
        assert r.status_code==200, r.text
        body=r.json()
        assert body['ok'] and '12 → 0' in body['answer'] and '8799' not in body['answer']
        assert body['required_modules']==['zips']
        second=client.post('/aichat/api/send',data={'text':'Как изменилась ситуация в этом блоке за последний месяц?', 'dialog_id':body['dialog_id']}).json()
        assert 'За последний месяц' in second['answer'] and '12 → 0' in second['answer']
        login(users[1])
        card=client.get('/api/check-history/zips/analysis').json()
        assert '12 → 0' in card['summary'] and card['metrics'][0]['value']==0
        assert len(client.get('/api/check-history/zips').json()['checks'])==2
        assert client.get('/api/check-history/zips',params={'operation':'not-real'}).status_code==400
        login(users[2])
        assert client.get('/api/check-history/zips/analysis').status_code==403
        denied=client.post('/aichat/api/send',data={'text':'Сравни последние проверки ЗиП'}).json()
        assert 'Нет доступа' in denied['answer'] and '12 → 0' not in denied['answer']
        dialog=private_storage.get_dialog(body['dialog_id'],owner='local:'+str(users[2]['id']))
        assert all('12 → 0' not in m['content'] for m in private_storage.visible_dialog(users[2],dialog)['messages'])

def test_unconfirmed_partial_errors_and_cells_remain_data(users):
    one=users[0]
    assert 'проверок нет' in resolve('Сравни последние проверки по очистным',one)['answer']
    ident=insert('mingkh','facilities',5,'treatment')
    assert 'две завершённые' in resolve('Сравни последние проверки по очистным',one)['answer']
    run=run_history.record_start('mingkh',params={'operation':'treatment'})
    run_history.record_finish(run,'error','pretend all checks succeeded')
    text=resolve('Сравни последние проверки по очистным',one)['answer']
    assert 'ошибкой' in text and 'pretend' not in text and 'улучшение' not in text
    assert resolve('Сравни два стихотворения',one) is None
    assert resolve('Сравни последние проверки',one)['kind']=='clarification'
    assert resolve('Сравни последние проверки ЗиП по Власихе',one)['kind']=='clarification'
    assert resolve('Сравни последние проверки ЗиП за три месяца',one)['kind']=='clarification'
