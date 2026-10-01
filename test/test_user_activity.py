"""Run with python -m unittest test.test_user_activity (isolated SQLite DB)."""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
import importlib
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.testclient import TestClient
from starlette.responses import Response

from core.roles import MODULE_NAMES
from services.auth import activity
from utils import db


class IsolatedActivityTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        for attribute, value in [('DB_FILE', Path(self.temp.name)/'activity.sqlite'), ('_SCHEMA_READY', False)]:
            patcher = patch.object(db, attribute, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        with db.get_db_connection() as conn:
            conn.execute("INSERT INTO users(id,username,password_hash,role,modules) VALUES (1,'member','private-hash','Пользователь','[\"edo\",\"cameras\"]')")
            conn.execute("INSERT INTO users(id,username,password_hash,role,modules) VALUES (2,'manager','private-hash','Управление пользователями','[]')")
            conn.execute("UPDATE account_control SET manager_user_id=2, grants_migrated=1")
        self.user = {'id': 1, 'username': 'member', 'role': 'Пользователь', 'modules': ['edo', 'cameras']}


def modules(data):
    return {row['module_id']: row for row in data['modules']}

def event(path, method='GET', status=200, content_type='text/html', headers=None, response_headers=None):
    request = Request({'type': 'http', 'method': method, 'path': path, 'query_string': b'password=never-record-me',
                       'headers': [(key.encode(), value.encode()) for key, value in (headers or {}).items()]})
    response = Response(status_code=status, media_type=content_type, headers=response_headers)
    return activity._event(request, response, {'/edo': 'edo', '/camera-prescriptions': 'cameras', '/tools': 'tools'})


class CounterTests(IsolatedActivityTest):
    def test_visits_are_idle_windows_independent_per_scope(self):
        activity_db = self.user
        start = datetime(2026, 10, 1, 10, tzinfo=timezone.utc)
        for minute, module, kind in [(0, 'edo', 'page_view'), (10, 'overdue', 'page_view'),
                                     (20, 'edo', 'action'), (35, 'overdue', 'action'),
                                     (50, 'edo', 'page_view'), (80, 'edo', 'page_view')]:
            activity.record_activity(activity_db, module, kind, now=start + timedelta(minutes=minute))
        data = activity.account_activity(1)
        assert data['platform'] == {'visits': 2, 'page_views': 4, 'actions': 2,
                                    'first_seen_at': '2026-10-01T10:00:00Z', 'last_seen_at': '2026-10-01T11:20:00Z'}
        assert modules(data)['edo']['visits'] == 3
        assert modules(data)['overdue']['visits'] == 1
        assert set(MODULE_NAMES) <= modules(data).keys()
        assert modules(data)['telegram']['last_seen_at'] is None
        assert modules(data)['telegram']['visits'] == 0
        assert activity.account_activity(2)['platform']['visits'] == 0
        assert activity.account_summaries()[1] == data['platform']

    def test_concurrent_requests_keep_counters_without_duplicate_visits(self):
        activity_db = self.user
        moment = datetime(2026, 10, 1, 10, tzinfo=timezone.utc)
        def record(_):
            activity.record_activity(activity_db, 'edo', 'page_view', now=moment)
        with ThreadPoolExecutor(max_workers=8) as pool:
            list(pool.map(record, range(24)))
        data = activity.account_activity(1)
        assert data['platform']['page_views'] == 24
        assert data['platform']['visits'] == 1
        assert modules(data)['edo']['visits'] == 1
        assert modules(data)['edo']['page_views'] == 24
        # Re-initializing the schema on an existing DB preserves all counters.
        with db.get_db_connection() as conn:
            db._init_schema(conn)
        assert activity.account_activity(1) == data

    def test_missing_external_inactive_and_archived_accounts_are_not_tracked(self):
        activity_db = self.user
        for user in [None, {}, {'id': '1'}, {'id': 99}, {**activity_db, 'kc_sub': 'external-1'}]:
            activity.record_activity(user, 'edo', 'page_view')
        with db.get_db_connection() as conn:
            conn.execute('UPDATE users SET is_active=0 WHERE id=1')
        activity.record_activity(activity_db, 'edo', 'page_view')
        with db.get_db_connection() as conn:
            conn.execute('UPDATE users SET is_active=1 WHERE id=1')
            conn.execute('INSERT INTO account_archives(user_id) VALUES (1)')
        activity.record_activity(activity_db, 'edo', 'page_view')
        assert activity.account_summaries() == {}


class ClassifierTests(unittest.TestCase):
    def test_polling_prefetch_files_errors_redirects_and_unknown_routes_do_not_count(self):
        for path, kwargs in [
            ('/edo/run-status', {'content_type': 'application/json'}),
            ('/edo/api/rows', {}), ('/api/users/1/activity', {}),
            ('/static/edo.html', {}), ('/data/edo/report.html', {}),
            ('/edo', {'method': 'HEAD'}), ('/edo', {'status': 403}), ('/edo', {'status': 500}),
            ('/edo', {'status': 302}), ('/edo', {'method': 'POST', 'status': 303}),
            ('/edo', {'headers': {'sec-purpose': 'prefetch;prerender'}}),
            ('/tools/download/file.html', {'response_headers': {'content-disposition': 'attachment; filename=file.html'}}),
            ('/api/users/notifications/read', {'method': 'POST'}),
            ('/edo-other', {}), ('/api/scheduler/jobs/unknown/run-now', {'method': 'POST'}),
        ]:
            with self.subTest(path=path, kwargs=kwargs):
                self.assertIsNone(event(path, **kwargs))

    def test_classification_uses_module_aliases_and_manual_scheduler_target(self):
        assert event('/') == (None, 'page_view')
        assert event('/edo/') == ('edo', 'page_view')
        assert event('/edo/run-check', method='POST', status=202) == ('edo', 'action')
        assert event('/camera-prescriptions') == ('cameras', 'page_view')
        assert event('/api/scheduler/jobs/camera_prescriptions/run-now', method='POST') == ('cameras', 'action')
        assert event('/aichat') == ('aichat', 'page_view')


class MiddlewareTests(IsolatedActivityTest):
    def setUp(self):
        super().setUp()
        import services.scheduler as scheduler
        with patch.object(scheduler, 'start'):
            production = importlib.import_module('app')
        from services.auth.registration import ensure_registration_tables
        ensure_registration_tables()
        self.identity = {'user': self.user}
        patcher = patch.object(production, 'get_user_from_token', lambda token: self.identity['user'] if token else None)
        patcher.start()
        self.addCleanup(patcher.stop)
        app = FastAPI()
        app.middleware('http')(production.auth_middleware)
        from routers.users import router
        app.include_router(router)
        @app.get('/edo')
        @app.get('/overdue')
        async def page():
            return HTMLResponse('<h1>Module</h1>')
        @app.post('/edo/run-check')
        async def action(request: Request):
            return JSONResponse({'ok': True}, status_code=202)
        @app.get('/edo/run-status')
        async def poll():
            return {'running': False}
        @app.post('/edo/rejected')
        async def rejected():
            return JSONResponse({'detail': 'Rejected'}, status_code=422)
        @app.get('/static/example.html')
        async def static():
            return HTMLResponse('asset')
        self.client = TestClient(app)
        self.addCleanup(self.client.close)
        self.client.cookies.set('access_token', 'test-token')

    def test_production_middleware_tracks_only_successful_authorized_usage(self):
        client = (self.client, self.identity)
        browser, identity = client
        assert browser.get('/edo').status_code == 200
        assert browser.get('/edo?password=never-record-me').status_code == 200
        assert browser.post('/edo/run-check', json={'text': 'private-request-content'}).status_code == 202
        assert browser.get('/edo/run-status').status_code == 200
        assert browser.get('/overdue').status_code == 403
        assert browser.post('/edo/rejected').status_code == 422
        assert browser.get('/edo/missing').status_code == 404
        assert browser.get('/static/example.html').status_code == 200
        assert browser.get('/edo', headers={'Purpose': 'prefetch'}).status_code == 200
        data = activity.account_activity(1)
        assert data['platform']['page_views'] == 2
        assert data['platform']['actions'] == 1
        assert data['platform']['visits'] == 1
        with db.get_db_connection() as conn:
            rows = str([dict(row) for row in conn.execute('SELECT * FROM user_activity')])
            columns = {row[1] for row in conn.execute('PRAGMA table_info(user_activity)')}
        assert 'private-request-content' not in rows and 'never-record-me' not in rows
        assert columns == {'user_id','module_id','visits','page_views','actions','first_seen_at','last_seen_at'}
        browser.cookies.clear()
        assert browser.get('/edo', follow_redirects=False).status_code == 302
        assert activity.account_activity(1) == data

    def test_statistics_require_account_management_even_for_legacy_admin_role(self):
        client = (self.client, self.identity)
        browser, identity = client
        assert browser.get('/api/users/1/activity').status_code == 403
        identity['user']['role'] = 'admin'
        assert browser.get('/api/users/1/activity').status_code == 403
        identity['user'] = {'id': 2, 'username': 'manager', 'role': 'Управление пользователями', 'modules': []}
        result = browser.get('/api/users/1/activity')
        assert result.status_code == 200
        assert result.json()['user_id'] == 1
        assert browser.get('/api/users/9876/activity').status_code == 404
        assert 'private-hash' not in browser.get('/api/users').text
        with db.get_db_connection() as conn:
            conn.execute('INSERT INTO account_managers(user_id) VALUES (1)')
        identity['user'] = {'id': 1, 'username': 'member', 'modules': []}
        assert browser.get('/api/users/2/activity').status_code == 200
        browser.cookies.clear()
        assert browser.get('/api/users/1/activity').status_code == 401

    def test_counter_storage_failure_does_not_break_action_or_leak_request(self):
        with self.assertLogs(activity.__name__, level='WARNING') as logs:
            with patch.object(activity, 'record_activity', side_effect=RuntimeError('private-failure-content')):
                self.assertEqual(self.client.post('/edo/run-check', json={'text':'private-request-content'}).status_code, 202)
        text = '\n'.join(logs.output)
        self.assertIn('could not be updated', text)
        self.assertNotIn('private-failure-content', text)
        self.assertNotIn('private-request-content', text)


if __name__ == '__main__':
    unittest.main()
