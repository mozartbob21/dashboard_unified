"""Real directory, SQLite and authenticated HTTP routes with synthetic local accounts."""
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from fastapi.staticfiles import StaticFiles

from utils import db
from services.auth import admin_directory as directory
from services.auth import service_center
from services.auth.registration import ensure_registration_tables


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'admin.sqlite')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    ensure_registration_tables()
    from routers import users, auth
    from services.auth import security
    monkeypatch.setattr(security, 'get_jwt_secret', lambda: 'synthetic-test-secret-32-characters')
    with db.get_db_connection() as conn:
        accounts = [('parent','Управление пользователями'), ('delegated','Пользователь'),
                    ('admin','Администратор'), ('externalrole','Администратор'), ('ordinary','Пользователь')]
        accounts += [('сотрудник%02d' % i, 'Контроль') for i in range(40)]
        for username, role in accounts:
            conn.execute("INSERT INTO users(username,password_hash,email,role,modules) VALUES (?,?,?,?,'[]')",
                         (username, 'unused', username+'@example.test', role))
        conn.execute("UPDATE account_control SET manager_user_id=1,grants_migrated=1")
        conn.executemany("INSERT INTO account_managers(user_id) VALUES (?)", [(2,), (3,)])
        conn.execute("INSERT INTO account_archives(user_id) VALUES (45)")
    app = FastAPI()
    app.mount("/static", StaticFiles(directory=Path(__file__).resolve().parents[1]/"static"), name="static")
    @app.middleware('http')
    async def session(request, handler):
        request.state.user = security.get_user_from_token(request.cookies.get('access_token'))
        return await handler(request)
    app.include_router(auth.router)
    app.include_router(users.router)
    return app


def client(app, username=None):
    from services.auth.security import create_access_token
    return TestClient(app, cookies={'access_token':create_access_token({'sub':username})} if username else {})


def test_search_pagination_and_counts(app):
    with client(app,'parent') as c:
        found = c.get('/api/users/admin-directory', params={'q':'СОТРУДНИК', 'page':2, 'size':10}).json()
        assert (found['page'],found['matched'],found['pages'],found['total']) == (2,39,4,45)
        assert [u['username'] for u in found['users']] == ['сотрудник%02d' % i for i in range(10,20)]
        assert (found['online'], found['offline']) == (0,45)
        assert c.get('/api/users/admin-directory?q=КОНТРОЛЬ').json()['matched'] == 39
        assert c.get('/api/users/admin-directory?q=ordinary%40example').json()['matched'] == 1
        assert c.get('/api/users/admin-directory?q=%25').json()['matched'] == 0
        assert c.get('/api/users/admin-directory?archive=archived').json()['matched'] == 1
        assert c.get('/api/users/admin-directory?size=150').status_code == 422


def test_presence_is_per_account_durable_and_expires_after_disconnect(app,monkeypatch):
    now = [10000]
    # Patch only the presence clock, keeping JWT times valid.
    from types import SimpleNamespace
    monkeypatch.setattr(directory, 'time', SimpleNamespace(time=lambda: now[0]))
    def ping(_):
        with client(app,'ordinary') as tab:
            # Client fields cannot select another identity or timestamp.
            return tab.post('/api/me/presence',json={'user_id':1,'last_seen':999999}).status_code
    with ThreadPoolExecutor(max_workers=3) as pool:
        assert list(pool.map(ping,range(3))) == [204]*3
    with client(app,'parent') as c:
        result = c.get('/api/users/admin-directory?q=ordinary').json()
        assert (result['online'],result['offline']) == (1,44)
        assert result['users'][0]['online']
        seen = result['users'][0]['last_online_at']
        assert seen
        with db.get_db_connection() as conn:
            assert conn.execute("SELECT count(*) FROM account_presence").fetchone()[0] == 1
        # Simulate reopening schema/DB after a process restart; no data reset.
        monkeypatch.setattr(db,'_SCHEMA_READY',False)
        now[0] += 60
        assert c.get('/api/users/admin-directory').json()['online'] == 1
        now[0] += 31
        assert c.get('/api/users/admin-directory').json()['online'] == 0
        with client(app,'delegated') as manager:
            stale = manager.get('/api/users/page?q=ordinary').json()['users'][0]
            assert stale['online'] is False and stale['last_online_at']==seen
            assert manager.get('/api/users/page?q=сотрудник00').json()['users'][0]['last_online_at'] is None
        assert ping(0) == 204
        assert c.get('/api/users/admin-directory').json()['online'] == 1
        with db.get_db_connection() as conn:
            conn.execute("UPDATE users SET is_active=0 WHERE username='ordinary'")
        assert c.get('/api/users/admin-directory').json()['online'] == 0
    assert client(app).post('/api/me/presence').status_code == 401


@pytest.mark.parametrize('actor,expected', [(None,403),('ordinary',403),('delegated',403),
                                             ('externalrole',403),('admin',200),('parent',200)])
def test_service_and_directory_authorization(app,actor,expected):
    with client(app,actor) as c:
        for route in ['/api/users/admin-directory','/api/users/service-center']:
            response = c.get(route)
            assert response.status_code == expected
            if expected == 200: assert response.headers['cache-control'] == 'private, no-store'
        if actor == 'delegated':
            basic = c.get('/api/users/page?size=3').json()
            assert len(basic['users']) == 3 and basic['online'] + basic['offline'] == basic['total']
            assert all('online' in u and 'last_online_at' in u for u in basic['users'])
            assert c.get('/api/users/page?q=Контроль').json()['matched'] == 0
            assert c.get('/api/users/record?username=delegated').status_code == 200
        if actor in ('parent','admin','delegated'):
            html=c.get('/users').text
            assert ('data-admin-panel="service"' in html) == (expected == 200)


def test_service_reports_actual_database_queue_and_unavailable_metrics(app,monkeypatch):
    monkeypatch.setattr(service_center, 'psutil', None)
    monkeypatch.setattr(service_center, '_cached', None)
    with db.get_db_connection() as conn:
        conn.execute("CREATE TABLE aichat_jobs(status TEXT)")
        conn.executemany("INSERT INTO aichat_jobs VALUES (?)", [('running',),('queued',),('queued',),('done',)])
    with client(app,'parent') as c:
        result=c.get('/api/users/service-center').json()
        assert result['cpu_percent'] is None and result['ram_total'] is None
        assert result['disk_total']>0 and result['database']=='Доступна'
        assert result['chat_jobs']=={'queued':2,'running':1}
        assert result['uptime_seconds']>=0 and result['updated_at']
        assert result['scope'] and str(db.DB_FILE) not in str(result)
        assert c.get('/api/users/service-center').json() == result
