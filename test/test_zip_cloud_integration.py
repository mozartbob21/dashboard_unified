"""Isolated admin/storage checks. Never use the office cloud or real credentials."""
import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from services.auth import integrations, accounts
from services.zip_curator import source
from utils import db


@pytest.fixture
def isolated_db(tmp_path, monkeypatch):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'auth.sqlite')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(integrations, 'KEY_FILE', tmp_path / 'integrations.key')
    with db.get_db_connection() as conn:
        conn.execute("INSERT INTO users(username,password_hash,role,modules) VALUES ('manager','unused','Пользователь','[]')")
        conn.execute("INSERT INTO users(username,password_hash,role,modules) VALUES ('ordinary','unused','Администратор','[\"zips\"]')")
        conn.execute("UPDATE account_control SET manager_user_id=(SELECT id FROM users WHERE username='manager'),grants_migrated=1")


@pytest.fixture
def api(isolated_db):
    from routers.users import router
    app = FastAPI()
    @app.middleware('http')
    async def identity(request, call_next):
        # Deterministic test authentication; permissions themselves are real DB checks.
        request.state.user = {'username': request.headers.get('X-Test-User', 'ordinary')}
        return await call_next(request)
    app.include_router(router)
    with TestClient(app) as client:
        yield client


def save(client, **extra):
    return client.put('/api/users/integrations/zip_cloud', headers={'X-Test-User': 'manager'},
                      json={'username': 'office-user', 'password': 'secret-for-test',
                            'source_url': source.DEFAULT_CLOUD_URL, **extra})


def test_cloud_settings_encrypted_and_password_retained(api):
    assert save(api, folder_path='/Shared/ЗиП').status_code == 200
    with db.get_db_connection() as conn:
        encrypted = conn.execute("SELECT encrypted_value FROM integration_credentials WHERE service='zip_cloud'").fetchone()[0]
    assert 'secret-for-test' not in encrypted
    assert 'office-user' not in encrypted
    assert integrations.credentials('zip_cloud')['password'] == 'secret-for-test'
    status = api.get('/api/users/integrations/zip_cloud', headers={'X-Test-User': 'manager'})
    assert status.status_code == 200
    assert 'secret-for-test' not in status.text
    assert status.json()['configured'] is True
    # A second device sees persisted URL/path; omission does not erase either.
    response = api.put('/api/users/integrations/zip_cloud', headers={'X-Test-User': 'manager'},
                       json={'username': 'office-user', 'password': ''})
    assert response.status_code == 200
    saved = integrations.credentials('zip_cloud')
    assert saved['password'] == 'secret-for-test'
    assert saved['folder_path'].rstrip('/') == '/Shared/ЗиП'
    assert '10968596' in saved['source_url']
    assert save(api, username='someone-else', password='').status_code == 400


def test_permissions_reject_read_write_and_probe(api, monkeypatch):
    monkeypatch.setattr(source, 'check_cloud_access', lambda: pytest.fail('unauthorized network call'))
    assert api.get('/api/users/integrations/zip_cloud').status_code == 403
    assert api.put('/api/users/integrations/zip_cloud', json={'username':'x','password':'x'}).status_code == 403
    assert api.post('/api/users/integrations/zip_cloud/check').status_code == 403


@pytest.mark.parametrize('url', [
    'http://cloud.mosreg.ru/apps/files/?fileid=10968596',
    'https://other.invalid/apps/files/?fileid=10968596',
    'https://cloud.mosreg.ru/apps/files/?fileid=0',
])
def test_invalid_source_does_not_replace_saved_settings(api, url):
    assert save(api).status_code == 200
    before = integrations.credentials('zip_cloud')
    assert save(api, source_url=url).status_code == 400
    assert integrations.credentials('zip_cloud') == before


def test_probe_safe_errors_and_success(api, monkeypatch):
    monkeypatch.setattr(source, 'check_cloud_access', lambda: {'ok':True,'files':3,'message':'Папка доступна'})
    response = api.post('/api/users/integrations/zip_cloud/check', headers={'X-Test-User':'manager'})
    assert response.status_code == 200 and response.json()['files'] == 3
    def fail():
        raise RuntimeError('https://office-user:secret-for-test@cloud.mosreg.ru/private')
    monkeypatch.setattr(source, 'check_cloud_access', fail)
    response = api.post('/api/users/integrations/zip_cloud/check', headers={'X-Test-User':'manager'})
    assert response.status_code == 502
    assert 'secret-for-test' not in response.text and '/private' not in response.text


def test_public_status_keeps_both_sources_without_credentials(isolated_db):
    integrations.save_credentials('office-user','secret-for-test','zip_cloud',source_url=source.DEFAULT_CLOUD_URL)
    status = source.public_status()
    assert status['provider'] == 'combined' and status['configured']
    assert [item['provider'] for item in status['sources']] == ['yandex', 'mosreg']
    assert all(item['configured'] for item in status['sources'])
    assert 'secret-for-test' not in str(status) and 'office-user' not in str(status)
    selected = source.resolve_sources(source.DEFAULT_YANDEX_URL)
    assert [item['provider'] for item in selected] == ['yandex', 'mosreg']
    assert selected[0]['url'] == source.DEFAULT_YANDEX_URL
    assert selected[1]['account']['username'] == 'office-user'
    assert source.resolve_source()['provider'] == 'yandex'


def test_unconfigured_cloud_url_is_not_sent_to_yandex(isolated_db, monkeypatch):
    monkeypatch.setenv('ZIP_FOLDER_A_URL', source.DEFAULT_CLOUD_URL)
    selected = source.resolve_sources()
    assert selected[0]['provider'] == 'yandex'
    assert selected[0]['url'] == source.DEFAULT_YANDEX_URL
    assert selected[1]['provider'] == 'mosreg' and 'Сохраните доступ' in selected[1]['error']


def test_unreadable_cloud_does_not_block_yandex(isolated_db, monkeypatch):
    from fastapi import HTTPException
    def broken(service):
        raise HTTPException(503, 'sensitive config detail')
    monkeypatch.setattr(source, 'credentials', broken)
    monkeypatch.setenv('ZIP_FOLDER_A_URL', 'existing-yandex-link')
    selected = source.resolve_sources()
    assert selected[0]['url'] == 'existing-yandex-link'
    assert selected[1]['provider'] == 'mosreg' and selected[1]['error']
    assert 'sensitive' not in str(source.public_status())
    assert source.public_status()['configured'] is True


def test_cloud_probe_does_not_probe_yandex(api, monkeypatch):
    assert save(api).status_code == 200
    from services.zip_curator import mosreg_cloud
    calls = []
    class FakeCloud:
        def __init__(self, session, url, **kwargs): calls.append(url)
        def list_files(self): return [{'name':'test.xlsx'}]
    monkeypatch.setattr(mosreg_cloud, 'CloudClient', FakeCloud)
    result = source.check_cloud_access()
    assert result['ok'] and result['files'] == 1
    assert calls == [source.DEFAULT_CLOUD_URL]
