"""Isolated end-to-end contracts; real SQLite, synthetic portal and sessions."""
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from datetime import date
import threading

import httpx
import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from routers import collective
from services.collective import client, store, refresh
from utils import db


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'dashboard.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(refresh, 'kick', lambda: None)
    monkeypatch.setattr(refresh, '_thread', None)
    monkeypatch.setattr(refresh, '_wake', threading.Event())
    monkeypatch.setattr(collective, 'is_account_manager', lambda user: bool(user and user.get('id') == 1))
    monkeypatch.setattr(httpx.HTTPTransport, 'handle_request', lambda *a, **kw: pytest.fail('Live network forbidden'))


def api(ident=2, modules=None):
    app = FastAPI()
    @app.middleware('http')
    async def session(request, call_next):
        request.state.user = {'id': ident, 'modules': modules if modules is not None else ['collective']}
        return await call_next(request)
    app.include_router(collective.router)
    return TestClient(app)


def source(number='123', status='В работе'):
    row = [''] * len(client.HEADERS)
    row[0], row[1], row[2], row[5], row[8], row[10], row[14] = 'ЕЦУР', '01.01.2026', 'Тестовый округ', number, 'Вода', '15', status
    return row


def snapshot():
    return store.save_snapshot('2026-01-01', '2026-01-02', [('1', 'МинЖКХ')], [source()])


def test_snapshot_saved_shared_between_accounts_and_no_source_request(monkeypatch):
    c = api()
    assert c.get('/mingkh/collective/api/data?start=2026-01-01&end=2026-01-02').status_code == 404
    saved = snapshot()
    monkeypatch.setattr(refresh.CollectivePortal, 'fetch', lambda *a: pytest.fail('Read must never fetch source'))
    for ident in [1, 2, 3]:
        response = api(ident).get('/mingkh/collective/api/data?start=2026-01-01&end=2026-01-02')
        assert response.status_code == 200
        assert response.json()['keys'] == saved['keys']
        assert response.headers['cache-control'] == 'private, no-store'


def test_source_jobs_concurrent_dedup_idempotent_after_completion_and_failure(monkeypatch):
    monkeypatch.setattr(refresh, 'credentials', lambda key: {'username': 'synthetic', 'password': 'synthetic'})
    monkeypatch.setattr(refresh.CollectivePortal, 'fetch', lambda *a: ([('1', 'МинЖКХ')], [source()]))
    def submit(i):
        return refresh.submit('2026-01-01', '2026-01-02', f'request-{i:020d}')
    with ThreadPoolExecutor(max_workers=4) as pool:
        tickets = list(pool.map(submit, range(8)))
    assert len({t['job_id'] for t in tickets}) == 1
    with pytest.raises(refresh.BusyError):
        refresh.submit('2026-01-01', '2026-01-03', 'different-request-01')
    refresh._drain()
    assert refresh.status(tickets[0]['job_id'])['status'] == 'success'
    assert submit(0)['job_id'] == tickets[0]['job_id']
    key = store.get_snapshot('2026-01-01', '2026-01-02')['keys'][0]
    store.update_workflow(key, 0, {'report': 'Сохранённый доклад'}, 'local:2')
    def unavailable(*args):
        raise client.PortalError('Источник временно недоступен.')
    monkeypatch.setattr(refresh.CollectivePortal, 'fetch', unavailable)
    ticket = submit(20)
    refresh._drain()
    assert refresh.status(ticket['job_id'])['error'] == 'Источник временно недоступен.'
    assert store.get_workflow(key)['report'] == 'Сохранённый доклад'


def test_restart_and_os_lock_leave_previous_snapshot(monkeypatch):
    saved = snapshot()
    ticket = refresh.submit('2026-01-01', '2026-01-02', 'restart-request-key')
    with db.get_db_connection() as conn:
        conn.execute("UPDATE collective_refresh_jobs SET status='running' WHERE id=?", (ticket['job_id'],))
    with refresh._collector_lock():
        refresh._drain()
        assert refresh.status(ticket['job_id'])['status'] == 'running'
    refresh._drain()
    assert refresh.status(ticket['job_id'])['status'] == 'error'
    assert store.get_snapshot('2026-01-01', '2026-01-02')['keys'] == saved['keys']


def test_shared_job_api_and_credentials_failure(monkeypatch):
    monkeypatch.setattr(refresh, 'credentials', lambda key: None)
    c = api()
    response = c.post('/mingkh/collective/api/refresh', json={'start':'2026-01-01', 'end':'2026-01-02'}, headers={'Idempotency-Key':'synthetic-request-01'})
    assert response.status_code == 202
    url = response.json()['status_url']
    assert api(3).get(url).json()['status'] == 'queued'
    refresh._drain()
    assert 'Администратору' in api(3).get(url).json()['error']
    assert c.post('/mingkh/collective/api/refresh', json={'start':'2026-01-01','end':'2026-01-02'}).status_code == 400


def test_directory_admin_only_workflow_shared_and_concurrent_conflict():
    saved = snapshot()
    key = saved['keys'][0]
    path = '/mingkh/collective/api/appeals/' + key
    ordinary, admin = api(2), api(1)
    assert ordinary.post('/mingkh/collective/api/assignees', json={'name':'Тестовый сотрудник'}).status_code == 403
    person = admin.post('/mingkh/collective/api/assignees', json={'name':'Тестовый сотрудник'}).json()['assignee']
    changed = ordinary.patch(path, json={'version':0, 'report':'<script>alert(1)</script>', 'assignee_id':person['id']})
    assert changed.status_code == 200
    assert api(3).patch(path, json={'version':0, 'report':'Запоздавшая правка'}).status_code == 409
    assert store.get_workflow(key)['report'] == '<script>alert(1)</script>'
    assert ordinary.patch('/mingkh/collective/api/assignees/' + str(person['id']), json={'version':person['version'], 'active':False}).status_code == 403
    assert admin.patch('/mingkh/collective/api/assignees/' + str(person['id']), json={'version':person['version'], 'active':False}).status_code == 200
    current = store.get_workflow(key)
    assert current['assignee_archived']
    assert ordinary.patch(path, json={'version':current['version'], 'report':'x'*20001}).status_code == 400
    assert ordinary.patch(path, json={'version':current['version'], 'user_id':1}).status_code == 400
    assert ordinary.post('/mingkh/collective/api/import-distribution', json={}).status_code == 403
    with ThreadPoolExecutor(max_workers=2) as pool:
        results = list(pool.map(lambda i: api(i+2).patch(path, json={'version':current['version'], 'report':'Правка '+str(i)}).status_code, range(2)))
    assert sorted(results) == [200, 409]


def test_all_endpoints_require_module_access_and_actor_from_session():
    key = snapshot()['keys'][0]
    denied = api(modules=['mingkh'])
    for url in ['/api/data', '/api/assignees', '/api/jobs/'+'a'*32]:
        assert denied.get('/mingkh/collective'+url).status_code == 403
    for url in ['/api/refresh', '/api/assignees', '/api/import-distribution', '/api/pptx']:
        assert denied.post('/mingkh/collective'+url, json={}).status_code == 403
    assert denied.patch('/mingkh/collective/api/appeals/'+key, json={}).status_code == 403
    assert api(None).patch('/mingkh/collective/api/appeals/'+key, json={'version':0,'report':'x'}).status_code == 401


def test_export_uses_server_report_and_assignee_not_client(monkeypatch):
    saved = snapshot()
    key = saved['keys'][0]
    person = store.create_assignee('Ответственный', 'local:1')
    store.update_workflow(key, 0, {'report':'Реальный доклад','assignee_id':person['id'],'theme':'kr'}, 'local:2')
    seen = []
    monkeypatch.setattr(presentation := collective.presentation, 'build', lambda rows, *a: seen.extend(rows) or b'presentation')
    response = api().post('/mingkh/collective/api/pptx', json={'start':'2026-01-01','end':'2026-01-02','rows':[{'_key':key,'_theme':'ii','_report':'Подмена','ОМСУ':'Подмена'}]})
    assert response.status_code == 200
    assert seen[0]['_report'] == 'Реальный доклад'
    assert seen[0]['_assignee'] == 'Ответственный'
    assert seen[0]['_theme'] == 'kr'
    assert seen[0]['ОМСУ'] == 'Тестовый округ'


def test_invalid_unicode_report_and_deep_json_are_rejected_without_writes():
    key = snapshot()['keys'][0]
    c = api()
    response = c.patch('/mingkh/collective/api/appeals/'+key, content=b'{"version":0,"report":"\\ud800"}', headers={'Content-Type':'application/json'})
    assert response.status_code == 400
    assert store.get_workflow(key)['version'] == 0
    assert c.post('/mingkh/collective/api/import-distribution', content=b'{' + b'"x":[' + b'['*1100 + b']'*1100 + b']}').status_code == 403
    assert api(1).post('/mingkh/collective/api/import-distribution', content=b'{"x":'+b'['*1100+b']'*1100+b'}').status_code == 400
