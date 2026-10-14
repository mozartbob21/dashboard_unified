"""The scan endpoint must never expose temporary download credentials."""
import ast
from pathlib import Path
from types import SimpleNamespace

import pytest
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from services.zip_curator import core
from services.zip_curator import source


@pytest.fixture(autouse=True)
def no_real_credentials(monkeypatch):
    monkeypatch.setattr(source, 'credentials', lambda service: None)
    monkeypatch.delenv('ZIP_FOLDER_A_URL', raising=False)


def client_with(sync):
    tree = ast.parse((Path(__file__).parents[1] / 'app.py').read_text(encoding='utf-8-sig'))
    handlers = ast.Module(body=[node for node in tree.body if isinstance(node, ast.AsyncFunctionDef)
                               and node.name == 'zc_scan'], type_ignores=[])
    app = FastAPI()
    zc = SimpleNamespace(sync_and_scan_folder_a=sync, scan_folder=lambda: (1, 0),
                         load_state=lambda: {'pending': [], 'clean': {}},
                         FolderSyncBusy=core.FolderSyncBusy, DiskSyncError=core.DiskSyncError)
    exec(compile(handlers, 'app.py', 'exec'), {'app': app, 'zc': zc, 'HTTPException': HTTPException})
    return TestClient(app)


@pytest.mark.parametrize('exception,status', [
    (core.FolderSyncBusy('Папка А уже обновляется.'), 409),
    (core.DiskSyncError('Яндекс Диск не отвечает.'), 502),
    (RuntimeError('500 https://downloader.disk.yandex.ru/private?token=SECRET'), 502),
])
def test_errors_are_safe(exception, status):
    def failing_sync(url):
        raise exception
    response = client_with(failing_sync).post('/zip_curator/api/scan', json={'public_url': 'synthetic'})
    assert response.status_code == status
    assert response.json()['ok'] is False
    assert 'SECRET' not in response.text
    assert 'https://' not in response.text
    if isinstance(exception, core.DiskSyncError):
        assert response.json()['error'] == str(exception)


def test_partial_success_returns_good_state():
    expected = {'ok': True, 'partial': True, 'failed': 1, 'downloaded': 2,
                'errors': [{'name': 'failed.xlsx', 'message': 'HTTP 500', 'retained': True}],
                'state': {'pending': [{'rso': 'РСО'}], 'clean': {}}}
    response = client_with(lambda url: expected).post('/zip_curator/api/scan', json={'public_url': 'synthetic'})
    assert response.status_code == 200
    assert response.json() == expected


def test_empty_request_uses_server_source(monkeypatch):
    calls = []
    monkeypatch.setenv('ZIP_FOLDER_A_URL', 'configured-yandex-folder')
    def sync(url):
        calls.append(url)
        return {'ok': True, 'partial': False}
    response = client_with(sync).post('/zip_curator/api/scan', json={})
    assert response.status_code == 200
    assert response.json()['ok'] is True
    assert response.json()['partial'] is False
    assert calls == ['configured-yandex-folder']


def test_saved_cloud_overrides_stale_frontend_yandex_url(monkeypatch):
    monkeypatch.setattr(source, 'credentials', lambda service: {
        'username': 'office-user', 'password': 'PRIVATE', 'source_url': source.DEFAULT_CLOUD_URL,
        'folder_path': '/shared/stock'})
    calls = []
    def sync(url, **kwargs):
        calls.append((url, kwargs))
        return {'ok': True, 'partial': False}
    response = client_with(sync).post('/zip_curator/api/scan', json={'public_url': source.DEFAULT_YANDEX_URL})
    assert response.status_code == 200
    assert calls == [(source.DEFAULT_CLOUD_URL, {'provider': 'mosreg',
        'account': {'username': 'office-user', 'password': 'PRIVATE'}, 'folder_path': '/shared/stock'})]
    assert 'PRIVATE' not in response.text
