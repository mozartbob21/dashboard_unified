"""Synthetic encrypted keys and HTTP transports only; never contact a model."""
import json
import pytest
from fastapi import HTTPException
from test.test_zip_cloud_integration import isolated_db, api
from services.auth import ai_credentials as tokens, integrations
from utils import db
from core.ai_errors import AIServiceError

HEADERS = {'X-Test-User': 'manager'}
BASE = '/api/users/ai-tokens'


def test_per_module_encryption_status_blank_and_delete(api, monkeypatch):
    monkeypatch.setenv('QWEN_API_KEY', 'fallback-for-test')
    secret = 'only-test-secret-aichat'
    assert api.put(BASE+'/aichat', headers=HEADERS, json={'token': secret}).status_code == 200
    with db.get_db_connection() as conn:
        row = conn.execute("SELECT encrypted_value FROM integration_credentials WHERE service='ai:aichat'").fetchone()
        notifications = str(conn.execute('SELECT message FROM account_notifications').fetchall())
    assert secret not in row[0] and secret not in notifications
    result = api.get(BASE, headers=HEADERS)
    assert secret not in result.text and 'fallback-for-test' not in result.text
    item = next(x for x in result.json()['items'] if x['module_id'] == 'aichat')
    assert item['source'] == 'saved' and item['override']
    assert tokens.resolve_key('aichat') == secret
    assert tokens.resolve_key('appeals') == 'fallback-for-test'
    assert api.put(BASE+'/aichat', headers=HEADERS, json={'token': ''}).status_code == 200
    assert tokens.resolve_key('aichat') == secret
    assert api.delete(BASE+'/aichat', headers=HEADERS).status_code == 200
    assert tokens.resolve_key('aichat') == 'fallback-for-test'


def test_permissions_are_current_for_all_operations(api):
    for method, url, kwargs in [('get', BASE, {}), ('put', BASE+'/aichat', {'json':{'token':'test'}}), ('delete', BASE+'/aichat', {})]:
        assert getattr(api, method)(url, **kwargs).status_code == 403
    assert api.get(BASE, headers=HEADERS).status_code == 200
    with db.get_db_connection() as conn:
        conn.execute("UPDATE users SET is_active=0 WHERE username='manager'")
    assert api.get(BASE, headers=HEADERS).status_code == 403
    assert api.put(BASE+'/aichat', headers=HEADERS, json={'token':'test'}).status_code == 403
    assert api.delete(BASE+'/aichat', headers=HEADERS).status_code == 403


@pytest.mark.parametrize('payload', [ {'token':'test-secret\nheader'}, {'token':'x'*8193}, {'token': ['test-secret']}, {'token': 'test-secret', 'url':'https://example.invalid'}, ['test-secret'] ])
def test_invalid_payload_never_echoes_token(api, payload):
    response=api.put(BASE+'/aichat', headers=HEADERS, json=payload)
    assert response.status_code == 400
    assert 'test-secret' not in response.text and 'xxxxx' not in response.text


def test_unknown_module_is_rejected(api):
    assert api.put(BASE+'/not-a-module', headers=HEADERS, json={'token':'test'}).status_code == 404
    assert api.delete(BASE+'/not-a-module', headers=HEADERS).status_code == 404


def test_invalid_encryption_fails_closed_and_can_be_replaced(isolated_db, monkeypatch):
    monkeypatch.setenv('QWEN_API_KEY','fallback-test')
    tokens.save('tools','saved-test')
    with db.get_db_connection() as conn:
        conn.execute("UPDATE integration_credentials SET encrypted_value='corrupted' WHERE service='ai:tools'")
    with pytest.raises(AIServiceError, match='AI_CREDENTIALS_ERROR'):
        tokens.resolve_key('tools')
    assert tokens.status('tools')['error']
    tokens.save('tools','recovered-test')
    assert tokens.resolve_key('tools') == 'recovered-test'


def test_runtime_read_does_not_create_or_migrate_db(tmp_path, monkeypatch):
    monkeypatch.setattr(db,'DB_FILE',tmp_path/'missing.db')
    monkeypatch.setenv('QWEN_API_KEY','env-test')
    monkeypatch.setenv('SUMMARIZER_QWEN_API_KEY','summarizer-test')
    assert tokens.resolve_key('aichat')=='env-test'
    assert tokens.resolve_key('summarizer')=='summarizer-test'
    assert not db.DB_FILE.exists()


def test_transport_uses_module_key_without_changing_endpoint_or_policy(isolated_db, monkeypatch):
    import httpx
    from services.summarizer.engine import _qwen_chat
    tokens.save('aichat','override-test')
    monkeypatch.setenv('QWEN_API_KEY','env-test')
    monkeypatch.setenv('QWEN_API_BASE','https://aiplatform.mosreg.ru/api/user-models/v1')
    monkeypatch.setenv('QWEN_MODEL','existing-model')
    calls=[]
    def post(url,**kwargs):
        calls.append((url,kwargs))
        return httpx.Response(200, json={'choices':[{'message':{'content':'test-answer'}}]}, request=httpx.Request('POST',url))
    monkeypatch.setattr(httpx,'post',post)
    assert _qwen_chat([],module_id='aichat')=='test-answer'
    url,kw=calls[-1]
    assert url=='https://aiplatform.mosreg.ru/api/user-models/v1/chat/completions'
    assert kw['json']['model']=='existing-model'
    assert kw['headers']['Authorization']=='Bearer override-test'
    assert kw['trust_env'] is False and kw['follow_redirects'] is False and kw['verify'] is not False
    _qwen_chat([],module_id='tools')
    assert calls[-1][1]['headers']['Authorization']=='Bearer env-test'
    tokens.save('aichat','updated-test')
    _qwen_chat([],module_id='aichat')
    assert calls[-1][1]['headers']['Authorization']=='Bearer updated-test'


def test_override_does_not_allow_external_endpoint(isolated_db, monkeypatch):
    from services.summarizer.engine import _qwen_chat
    from core.privacy import PrivacyError
    import httpx
    tokens.save('aichat','test-only')
    monkeypatch.setattr(httpx,'post',lambda *a,**kw:pytest.fail('network forbidden'))
    with pytest.raises(PrivacyError):
        _qwen_chat([],module_id='aichat',base='https://api.openai.com/v1')


def test_rating_source_key_is_separate_encrypted_and_permission_checked(api, monkeypatch):
    from services.auth import source_api_keys as source
    route='/api/users/source-api-keys/omsu_rating'
    monkeypatch.setenv('OMSU_RATING_API_KEY','source-env-test')
    for method, options in [('get',{}),('put',{'json':{'token':'test-secret'}}),('delete',{})]:
        assert getattr(api,method)(route,**options).status_code==403
    assert api.put(route,headers=HEADERS,json={'token':'source-saved-test'}).status_code==200
    assert source.get_source_api_key('omsu_rating')=='source-saved-test'
    with db.get_db_connection() as conn:
        encrypted=conn.execute("SELECT encrypted_value FROM integration_credentials WHERE service='source-api:omsu_rating'").fetchone()[0]
    assert 'source-saved-test' not in encrypted
    response=api.get(route,headers=HEADERS)
    assert response.json()['override'] and 'source-saved-test' not in response.text
    assert api.put(route,headers=HEADERS,json={'token':''}).status_code==200
    assert source.get_source_api_key('omsu_rating')=='source-saved-test'
    assert api.delete(route,headers=HEADERS).status_code==200
    assert source.get_source_api_key('omsu_rating')=='source-env-test'
    assert api.put(route,headers=HEADERS,json={'token':'test-secret\ninjected'}).status_code==400
    assert api.put(route,headers=HEADERS,content='not-json-test-secret').status_code==400
    assert api.get('/api/users/source-api-keys/unknown',headers=HEADERS).status_code==404


def test_rating_source_corruption_fails_closed(isolated_db, monkeypatch):
    from services.auth import source_api_keys as source
    monkeypatch.setenv('OMSU_RATING_API_KEY','source-env-test')
    source.save('omsu_rating','source-saved-test')
    with db.get_db_connection() as conn:
        conn.execute("UPDATE integration_credentials SET encrypted_value='corrupted' WHERE service='source-api:omsu_rating'")
    with pytest.raises(HTTPException) as error:
        source.get_source_api_key('omsu_rating')
    assert error.value.status_code==503
    assert source.status('omsu_rating')['error']
