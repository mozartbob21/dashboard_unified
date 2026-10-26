"""Administrator source settings use encrypted storage; probes do not send files."""
import pytest
from test.test_zip_cloud_integration import isolated_db, api
from services.auth import integrations
from services.max_bot import presentation_config
from utils import db
from services.max_bot.presentation_source import SourceError


def test_credentials_stay_server_side_and_encrypted(api):
    payload = {'username': 'presentation-test', 'password': 'only-a-test-secret'}
    route = '/api/users/integrations/presentations'
    assert api.get(route).status_code == 403
    assert api.put(route, json=payload).status_code == 403
    assert api.post(route + '/check').status_code == 403
    headers = {'X-Test-User': 'manager'}
    assert api.put(route, headers=headers, json=payload).status_code == 200
    with db.get_db_connection() as conn:
        raw = conn.execute("SELECT encrypted_value FROM integration_credentials WHERE service='presentations'").fetchone()[0]
    assert payload['username'] not in raw and payload['password'] not in raw
    public = api.get(route, headers=headers)
    assert public.json()['configured']
    assert payload['password'] not in public.text
    assert api.put(route, headers=headers, json={**payload, 'password': ''}).status_code == 200
    assert integrations.credentials('presentations')['password'] == payload['password']


def test_probe_only_reads_list_and_sanitizes_unknown_failures(api, monkeypatch):
    class Source:
        def __enter__(self): return self
        def __exit__(self, *_): pass
        def list_attachments(self): return [object(), object()]
        def download(self, *_): raise AssertionError('probe must not download')
    monkeypatch.setattr(presentation_config, 'source_factory', Source)
    route = '/api/users/integrations/presentations/check'
    headers = {'X-Test-User': 'manager'}
    response = api.post(route, headers=headers)
    assert response.status_code == 200
    assert 'Файлов в источнике: 2' in response.json()['message']
    def fail(): raise RuntimeError('private-secret private-file-name')
    monkeypatch.setattr(presentation_config, 'source_factory', fail)
    response = api.post(route, headers=headers)
    assert response.status_code == 502 and 'private-' not in response.text


def test_worker_reads_saved_credentials_without_database_migrations(isolated_db, monkeypatch):
    integrations.save_credentials('test-user', 'test-password', 'presentations')
    def forbidden(*_, **__): raise AssertionError('worker must not migrate accounts')
    monkeypatch.setattr(integrations, 'credentials', forbidden)
    monkeypatch.setattr(db, 'get_db_connection', forbidden)
    monkeypatch.setattr(presentation_config, 'RedmineSource', lambda user, password: (user, password))
    assert presentation_config.source_factory() == ('test-user', 'test-password')


def test_missing_platform_database_is_not_created(tmp_path, monkeypatch):
    missing = tmp_path / 'missing.sqlite'
    monkeypatch.setattr(db, 'DB_FILE', missing)
    with pytest.raises(SourceError) as error:
        presentation_config.source_factory()
    assert error.value.code == 'SOURCE_NOT_CONFIGURED'
    assert not missing.exists()
