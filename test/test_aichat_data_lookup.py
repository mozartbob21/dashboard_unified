"""Questions via the actual authenticated durable chat endpoint, no model or provider."""
from unittest.mock import patch
import time
from fastapi.testclient import TestClient
from test.test_check_dynamics import users
from test.test_aichat_treatment_lookup import snapshot
from services.aichat import data_lookup, jobs
from services.auth import security

def test_background_search_returns_confirmed_snapshot_values_and_hides_from_other_grants(users,snapshot,monkeypatch):
    with patch('services.scheduler.start'):
        import app
    monkeypatch.setattr(app,'aichat_ask',lambda *a,**kw: (_ for _ in ()).throw(AssertionError('unexpected model invocation')))
    with TestClient(app.app) as c:
        c.cookies.set('access_token',security.create_access_token({'sub':users[0]['username']}))
        r=c.post('/aichat/api/send',data={'text':'Какие жалобы по очистным в работе в Дубне?', 'request_id':'lookup_test_1234567890','background':'true'})
        assert r.status_code==202,r.text
        ticket=r.json()
        for _ in range(200):
            state=c.get(ticket['status_url']).json()
            if state['status'] not in jobs.ACTIVE: break
            time.sleep(0.01)
        assert state['status']=='succeeded',state
        result=state['result']
        assert result['required_modules']==['mingkh']
        assert 'Найдено записей: 1.' in result['answer'] and '№89' in result['answer'].replace(' ','')
        assert '2026-11-03' in result['answer'] and '/mingkh/treatment' in result['answer'] and 'SECRET' not in result['answer']
        c.cookies.set('access_token',security.create_access_token({'sub':users[1]['username']}))
        private=c.get('/aichat/api/dialogs/'+ticket['dialog_id']).text
        assert '2026-11-03' not in private and 'SECRET' not in private
        denied=c.post('/aichat/api/send',data={'text':'Найди жалобу по очистным №89'}).json()
        assert 'Нет доступа' in denied['answer'] and '2026-11-03' not in denied['answer']

def test_read_only_intent_leaves_reports_and_normal_chat_unchanged(users):
    user=users[0]
    assert data_lookup.resolve('Расскажи сказку о воде',user) is None
    assert data_lookup.resolve('Подготовь отчёт по остаткам ЗиП',user) is None
    assert data_lookup.resolve('Как найти кнопку для ЗиП?',user) is None
    assert 'В каком' in data_lookup.resolve('Найди в системе строку с артикулом',user)['answer']
