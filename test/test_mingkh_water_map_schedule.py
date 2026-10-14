"""Offline integration of incremental map refresh with persisted scheduling."""
from datetime import date, datetime, timedelta, timezone
import json
import threading
from types import SimpleNamespace
from unittest.mock import Mock

from fastapi import FastAPI
from fastapi.testclient import TestClient
import pytest

from core.roles import allowed_history_modules, check_module_access
from services import scheduler, run_history
from services.mingkh import water_map as maps, water_map_runner as runner
from services.edds import dobrodel
from utils import db

JOB = 'mingkh-water-map'


@pytest.fixture
def isolated(monkeypatch, tmp_path):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'dashboard.sqlite')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(scheduler, '_SEEDED', False)
    monkeypatch.setattr(maps, 'DATA_FILE', tmp_path / 'map.json.gz')
    monkeypatch.setattr(maps, 'SEED_FILE', tmp_path / 'absent-seed.gz')
    monkeypatch.setattr(maps, '_seed_checked', None)
    monkeypatch.setattr(maps, 'LOCK', threading.Lock())
    monkeypatch.setattr(maps, 'STATE', {'busy': False, 'message': '', 'error': None, 'completed_at': None})
    monkeypatch.setattr(runner, 'credentials', Mock(return_value={'username': 'synthetic', 'password': 'not-real'}))
    monkeypatch.setattr(dobrodel.DobrodelClient, 'login', Mock(side_effect=AssertionError('Network prohibited')))
    fixed = datetime(2026, 10, 14, 9, tzinfo=timezone.utc)
    monkeypatch.setattr(scheduler, '_now', lambda: fixed)
    with db.get_db_connection():
        pass
    return SimpleNamespace(root=tmp_path, now=fixed)


def seed_map():
    points = {'1': {'id': 1, 'created': '2026-10-12 10:00', 'omsu': 'Округ', 'kind': 'hvs',
                    'address': 'Адрес', 'fact': 'Водоснабжение', 'status': 'В работе',
                    'lat': 55.7, 'lon': 37.6, 'geo': 'ok'},
              '2': {'id': 2, 'created': '2024-01-01 10:00', 'omsu': 'Округ', 'kind': 'hvs',
                    'address': 'Старая жалоба', 'fact': 'Водоснабжение', 'status': 'Решено',
                    'lat': 55.8, 'lon': 37.7, 'geo': 'ok'}}
    maps.import_upload(json.dumps({'updated': '12.10.2026 10:00', 'points': points}).encode())


def test_new_job_is_enabled_once_and_existing_admin_settings_survive_seed(isolated):
    job = scheduler.get_job(JOB)
    assert job['enabled'] is True and job['interval_minutes'] == 120
    assert datetime.fromisoformat(job['next_run_at']) == isolated.now + timedelta(hours=2)
    assert job['module_label'] == 'МИНЖКХ — карта жалоб'
    assert all(not item['enabled'] for item in scheduler.get_all_jobs() if item['module_id'] != JOB)
    scheduler.set_interval(JOB, 360)
    scheduler.set_enabled(JOB, False)
    previous = scheduler.get_job(JOB)
    scheduler._SEEDED = False
    current = scheduler.get_job(JOB)
    for key in ('enabled', 'interval_minutes', 'next_run_at', 'last_run_at', 'last_status'):
        assert current[key] == previous[key]


def test_existing_enabled_schedule_is_not_reset_on_restart(isolated):
    scheduler.get_job(JOB)
    with db.get_db_connection() as conn:
        conn.execute("UPDATE scheduler_jobs SET interval_minutes=77, next_run_at='2026-10-14T10:17:00+00:00', last_status='error', last_error='previous' WHERE module_id=?", (JOB,))
    scheduler._SEEDED = False
    job = scheduler.get_job(JOB)
    assert job['enabled'] and job['interval_minutes'] == 77
    assert job['next_run_at'] == '2026-10-14T10:17:00+00:00'
    assert job['last_status'] == 'error' and job['last_error'] == 'previous'


def test_due_tick_launches_existing_inprocess_collector_and_persists_outcome(isolated, monkeypatch):
    scheduler.get_job(JOB)
    with db.get_db_connection() as conn:
        conn.execute('UPDATE scheduler_jobs SET next_run_at=? WHERE module_id=?', ((isolated.now - timedelta(seconds=1)).isoformat(), JOB))
    calls = []
    monkeypatch.setattr(runner, 'run', lambda: calls.append(JOB) or {'status': 'success', 'error': ''})
    class InlineThread:
        def __init__(self, *, target, **kwargs):
            self.target = target
        def start(self):
            self.target()
    monkeypatch.setattr(scheduler, 'threading', SimpleNamespace(Thread=InlineThread))
    monkeypatch.setattr(scheduler, '_get_run_status_dict', Mock(side_effect=AssertionError('Map must not import the app subprocess runner')))
    scheduler._tick()
    assert calls == [JOB]
    job = scheduler.get_job(JOB)
    assert job['last_status'] == 'success' and job['last_error'] == ''
    assert datetime.fromisoformat(job['next_run_at']) == isolated.now + timedelta(hours=2)
    scheduler._tick()
    assert calls == [JOB]


def test_live_map_lock_skips_duplicate_schedule(isolated, monkeypatch):
    scheduler.get_job(JOB)
    maps._state(busy=True)
    launch = Mock()
    monkeypatch.setattr(runner, 'run', launch)
    scheduler._launch_module(JOB)
    launch.assert_not_called()
    assert scheduler.get_job(JOB)['last_run_at'] is None


def test_adapter_uses_same_incremental_refresh_and_records_history(isolated, monkeypatch):
    seed_map()
    calls = []
    class Portal:
        def __init__(self, **account):
            assert account == {'username': 'synthetic', 'password': 'not-real'}
        def __enter__(self):
            return self
        def __exit__(self, *args):
            pass
        def login(self):
            calls.append('login')
        def fetch_all(self, filters):
            calls.append(filters)
            return [{'cardId': 1, 'created': '12.10.2026 10:00:00', 'district': 'Округ',
                     'ecurFact': 'Водоснабжение', 'address': 'Адрес', 'status': 'Решено'}]
        def request(self, *args, **kwargs):
            raise AssertionError('Known coordinates must not be fetched again')
    monkeypatch.setattr(dobrodel, 'DobrodelClient', Portal)
    # Fix the end date while exercising the actual merge/filter/atomic save code.
    refresh = maps.refresh
    monkeypatch.setattr(maps, 'refresh', lambda client, **kwargs: refresh(client, date(2026, 10, 14), **kwargs))
    assert runner.run() == {'status': 'success', 'error': ''}
    points = maps._expand(maps.load())
    assert set(points) == {'1', '2'}
    assert points['1']['status'] == 'Решено' and points['1']['lat'] == 55.7
    assert points['2']['created'].startswith('2024-01-01')
    assert calls[1]['filters.createdAfter'] == '2026-10-05'
    assert calls[1]['filters.curators'] == maps.CURATOR
    with db.get_db_connection() as conn:
        history = conn.execute('SELECT * FROM run_history').fetchall()
    assert len(history) == 1 and history[0]['module_id'] == JOB and history[0]['status'] == 'success'
    assert history[0]['module_label'] == 'МИНЖКХ — карта жалоб'
    assert not maps.status()['busy'] and not maps.LOCK.locked()


def test_missing_credentials_and_failed_portal_keep_snapshot_and_finish_history(isolated, monkeypatch):
    seed_map()
    before = maps.DATA_FILE.read_bytes()
    monkeypatch.setattr(runner, 'credentials', lambda service: None)
    result = runner.run()
    assert result['status'] == 'error' and 'Добродел' in result['error']
    assert maps.DATA_FILE.read_bytes() == before
    monkeypatch.setattr(runner, 'credentials', lambda service: {'username': 'synthetic', 'password': 'not-real'})
    monkeypatch.setattr(dobrodel.DobrodelClient, 'login', Mock(side_effect=dobrodel.DobrodelError('login')))
    result = runner.run()
    assert result['status'] == 'error'
    assert maps.DATA_FILE.read_bytes() == before
    assert not maps.status()['busy'] and not maps.LOCK.locked()
    with db.get_db_connection() as conn:
        assert [row['status'] for row in conn.execute('SELECT status FROM run_history')] == ['error', 'error']


def test_process_lock_blocks_other_writer_without_losing_thread_lock(isolated):
    seed_map()
    before = maps.DATA_FILE.read_bytes()
    with maps._process_lock():
        assert runner.run()['status'] == 'skipped'
        assert not maps.LOCK.locked()
        with pytest.raises(maps.StoreBusy):
            maps.import_upload(b'invalid')
    assert maps.DATA_FILE.read_bytes() == before
    assert not maps.LOCK.locked()


def test_mingkh_alias_grants_history_but_only_managers_can_control_job(isolated, monkeypatch):
    from services.auth import accounts
    identity = {'user': {'username': 'ordinary', 'modules': ['mingkh']}, 'manager': False}
    monkeypatch.setattr(accounts, 'is_account_manager', lambda user: bool(user and identity['manager']))
    monkeypatch.setattr(scheduler, 'get_user_from_token', lambda token: None)
    app = FastAPI()
    @app.middleware('http')
    async def user_state(request, call_next):
        request.state.user = identity['user']
        return await call_next(request)
    app.include_router(scheduler.router)
    assert check_module_access(identity['user'], JOB)
    assert JOB in allowed_history_modules(identity['user'])
    with TestClient(app) as client:
        assert client.get('/api/scheduler/jobs').json()['jobs'] == []
        for tail in ['', '/enable', '/disable', '/interval', '/run-now']:
            response = client.get('/api/scheduler/jobs/' + JOB) if not tail else client.post('/api/scheduler/jobs/' + JOB + tail, json={'interval_minutes': 30})
            assert response.status_code == 403
        identity['manager'] = True
        assert [job['module_id'] for job in client.get('/api/scheduler/jobs').json()['jobs']] == [JOB]
        assert client.post('/api/scheduler/jobs/' + JOB + '/interval', json={'interval_minutes': 180}).status_code == 200
        assert client.post('/api/scheduler/jobs/' + JOB + '/disable').json()['job']['enabled'] is False
        identity['user'] = {'username': 'manager', 'modules': []}
        assert client.get('/api/scheduler/jobs/' + JOB).status_code == 403
        identity['user'] = None
        assert client.get('/api/scheduler/jobs/' + JOB).status_code == 401


def test_another_process_cannot_replace_map_during_refresh(isolated):
    import os
    import subprocess
    import sys
    seed_map()
    before = maps.DATA_FILE.read_bytes()
    child = """
import importlib.util, sys
from pathlib import Path
spec = importlib.util.spec_from_file_location('map_child', sys.argv[1])
maps = importlib.util.module_from_spec(spec)
spec.loader.exec_module(maps)
maps.DATA_FILE = Path(sys.argv[2])
try:
    maps.import_upload(b'invalid')
except maps.StoreBusy:
    raise SystemExit(0)
raise SystemExit('Import did not respect another process lock')
"""
    with maps._process_lock():
        assert maps.status()['busy']
        result = subprocess.run([sys.executable, '-c', child, maps.__file__, str(maps.DATA_FILE)],
                                env={**os.environ, 'PYTHONDONTWRITEBYTECODE': '1'},
                                capture_output=True, text=True, timeout=10)
    assert result.returncode == 0, result.stderr
    assert maps.DATA_FILE.read_bytes() == before
    assert not maps.status()['busy']


def test_status_is_readable_without_maintenance_and_reflects_saved_schedule(isolated, monkeypatch):
    from fastapi.staticfiles import StaticFiles
    from pathlib import Path
    from routers import mingkh
    from services.auth import accounts
    identity = {'manager': False}
    monkeypatch.setattr(accounts, 'is_account_manager', lambda user: identity['manager'])
    app = FastAPI()
    @app.middleware('http')
    async def user_state(request, call_next):
        request.state.user = {'username': 'reader', 'modules': ['mingkh']}
        return await call_next(request)
    # core.web templates can render the route without starting the application.
    app.mount('/static', StaticFiles(directory=str(Path(mingkh.__file__).resolve().parents[1] / 'static'), check_dir=False), name='static')
    app.include_router(mingkh.router)
    with TestClient(app) as client:
        job = scheduler.get_job(JOB)
        scheduler.mark_run_finished(JOB, 'error', 'Администратору нужно сохранить доступ к Доброделу.')
        response = client.get('/mingkh/api/water-map/status')
        data = response.json()
        assert response.status_code == 200
        assert response.headers['cache-control'] == 'private, no-store'
        assert data['automatic_enabled'] is True and data['interval_minutes'] == 120
        assert data['next_run_at'] == job['next_run_at'] and data['last_status'] == 'error'
        assert 'последние сохранённые данные' in data['error']
        assert 'Добродел' not in data['error']
        assert data['dataset_version'] is None
        scheduler.set_enabled(JOB, False)
        assert client.get('/mingkh/api/water-map/status').json()['automatic_enabled'] is False
        for endpoint in ('import', 'refresh'):
            assert client.post('/mingkh/api/water-map/' + endpoint, content=b'').status_code == 403
        identity['manager'] = True
        assert 'Добродел' in client.get('/mingkh/api/water-map/status').json()['error']
        seed_map()
        data_response = client.get('/mingkh/api/water-map')
        assert data_response.status_code == 200
        assert data_response.headers['x-water-map-version'] == client.get('/mingkh/api/water-map/status').json()['dataset_version']


def test_manual_scheduler_action_is_attributed_to_mingkh(isolated):
    from fastapi import Request
    from fastapi.responses import JSONResponse
    from services.auth import activity
    request = Request({'type': 'http', 'method': 'POST',
                       'path': '/api/scheduler/jobs/mingkh-water-map/run-now', 'headers': []})
    assert activity._event(request, JSONResponse({'ok': True}), {}) == ('mingkh', 'action')
