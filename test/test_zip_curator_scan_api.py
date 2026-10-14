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
    zc = SimpleNamespace(sync_and_scan_sources=sync, scan_folder=lambda: (1, 0),
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
    assert calls == [[{'provider':'yandex','label':'Яндекс Диск','url':'configured-yandex-folder'}]]


def test_saved_cloud_complements_existing_yandex_url(monkeypatch):
    monkeypatch.setattr(source, 'credentials', lambda service: {
        'username': 'office-user', 'password': 'PRIVATE', 'source_url': source.DEFAULT_CLOUD_URL,
        'folder_path': '/shared/stock'})
    calls = []
    def sync(selected):
        calls.append(selected)
        return {'ok': True, 'partial': False}
    response = client_with(sync).post('/zip_curator/api/scan', json={'public_url': source.DEFAULT_YANDEX_URL})
    assert response.status_code == 200
    assert [item['provider'] for item in calls[0]] == ['yandex','mosreg']
    assert calls[0][0]['url'] == source.DEFAULT_YANDEX_URL
    assert calls[0][1]['url'] == source.DEFAULT_CLOUD_URL
    assert calls[0][1]['account'] == {'username':'office-user','password':'PRIVATE'}
    assert 'PRIVATE' not in response.text


def test_partial_source_failure_still_returns_success_and_results():
    result = {'ok':True,'partial':True,'downloaded':1,'duplicates':1,
              'sources':[{'provider':'yandex','label':'Яндекс Диск','ok':True},
                         {'provider':'mosreg','label':'Облако МО','ok':False,'error':'Недоступно'}],
              'state':{'pending':[{'rso':'Успешно загруженное РСО'}],'clean':{}}}
    response = client_with(lambda selected: result).post('/zip_curator/api/scan',json={})
    assert response.status_code == 200 and response.json() == result
    assert response.headers['cache-control'] == 'private, no-store'


def test_total_failure_has_per_source_diagnostics_and_502():
    result = {'ok':False,'error':'Не удалось обновить источники.',
              'sources':[{'provider':'yandex','label':'Яндекс Диск','ok':False,'error':'Таймаут'},
                         {'provider':'mosreg','label':'Облако МО','ok':False,'error':'Проверьте папку'}]}
    response = client_with(lambda selected: result).post('/zip_curator/api/scan',json={})
    assert response.status_code == 502 and response.json() == result


# Exercise the actual resolver, sync, parser, state persistence and response together.
from test.test_zip_curator_combined_sync import sources, add_source, ACCOUNT
from test.test_zip_curator_folder_sync import local_store, workbook, PUBLIC_URL


def test_scan_end_to_end_merges_sources_deduplicates_and_survives_cloud_failure(monkeypatch, local_store, sources):
    monkeypatch.setenv('ZIP_FOLDER_A_URL', PUBLIC_URL)
    monkeypatch.setattr(source, 'credentials', lambda service: {
        **ACCOUNT, 'source_url': source.DEFAULT_CLOUD_URL, 'folder_path':'/FolderA'})
    data = workbook(local_store, 'Общее РСО', 3)
    add_source(sources, 'yandex', 'original.xlsx', data)
    add_source(sources, 'mosreg', 'copy.xlsx', data)
    with client_with(core.sync_and_scan_sources) as api:
        first = api.post('/zip_curator/api/scan', json={})
        assert first.status_code == 200
        assert first.json()['duplicates'] == 1
        assert first.json()['added'] == 1
        assert len(first.json()['state']['pending']) == 1
        sources['mosreg']['files'] = core.DiskSyncError('Папка облака недоступна')
        add_source(sources, 'yandex', 'fresh.xlsx', workbook(local_store, 'Новое РСО', 5))
        second = api.post('/zip_curator/api/scan', json={})
        assert second.status_code == 200 and second.json()['partial']
        assert second.json()['sources'][0]['ok'] is True
        assert second.json()['sources'][1]['ok'] is False
        assert {row['rso'] for row in second.json()['state']['pending']} == {'Общее РСО','Новое РСО'}
        state_before_failure = core.STATE_FILE.read_bytes()
        sources['yandex']['files'] = core.DiskSyncError('Яндекс временно недоступен')
        failed = api.post('/zip_curator/api/scan', json={})
        assert failed.status_code == 502 and failed.json()['ok'] is False
        assert core.STATE_FILE.read_bytes() == state_before_failure
        assert ACCOUNT['password'] not in first.text + second.text + failed.text
