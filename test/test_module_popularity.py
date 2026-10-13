"""All-time module ranking uses existing counters and account-manager permissions."""
from datetime import datetime, timedelta, timezone
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from fastapi.testclient import TestClient

from core.roles import MODULE_NAMES
from services.auth import activity
from utils import db


class ModulePopularityTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        for attribute, value in [('DB_FILE', Path(temp.name) / 'ranking.sqlite'), ('_SCHEMA_READY', False)]:
            patcher = patch.object(db, attribute, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        with db.get_db_connection() as conn:
            for user_id, username in [(1, 'first-member'), (2, 'second-member'), (3, 'manager')]:
                conn.execute("INSERT INTO users(id,username,password_hash,role,modules) VALUES (?,?,'private-hash','Пользователь','[]')",
                             (user_id, username))
            conn.execute('UPDATE account_control SET manager_user_id=3, grants_migrated=1')
            conn.execute("UPDATE user_activity_collection SET started_at='2026-10-01T07:05:38Z'")
        self.users = [{'id': 1, 'username': 'first-member'}, {'id': 2, 'username': 'second-member'}]
        self.start = datetime(2026, 10, 2, 10, tzinfo=timezone.utc)

    def test_empty_store_includes_every_catalog_and_shared_section_in_stable_order(self):
        result = activity.module_popularity()
        assert result['period'] == 'all_time'
        assert result['tracking_since'] == '2026-10-01T07:05:38Z'
        assert result['visit_gap_minutes'] == 30
        assert datetime.fromisoformat(result['updated_at'].replace('Z', '+00:00')).tzinfo is not None
        assert [row['module_id'] for row in result['modules']] == list(activity.SECTION_NAMES)
        assert set(MODULE_NAMES) <= {row['module_id'] for row in result['modules']}
        for rank, row in enumerate(result['modules'], 1):
            assert row == {'rank': rank, 'module_id': row['module_id'], 'label': activity.SECTION_NAMES[row['module_id']],
                           'visits': 0, 'unique_users': 0, 'page_views': 0, 'actions': 0, 'last_seen_at': None}

    def test_totals_rank_across_users_and_retain_archived_or_disabled_history(self):
        for user, module, event, minute in [
            (0, 'edo', 'page_view', 0), (0, 'edo', 'action', 10), (0, 'edo', 'page_view', 40),
            (1, 'edo', 'page_view', 5), (1, 'mingkh', 'page_view', 0),
            (1, 'aichat', 'action', 0), (1, 'aichat', 'action', 30),
        ]:
            activity.record_activity(self.users[user], module, event, now=self.start + timedelta(minutes=minute))
        before = [activity.account_activity(user['id']) for user in self.users]
        with db.get_db_connection() as conn:
            conn.execute('UPDATE users SET is_active=0 WHERE id=1')
            conn.execute('INSERT INTO account_archives(user_id) VALUES (2)')
            # Legacy/deleted section IDs are not invented as visible blocks.
            conn.execute("INSERT INTO user_activity VALUES (3,'unknown-old-module',500,500,0,?,?)", ('2026-10-02T10:00:00Z',)*2)
        result = activity.module_popularity()
        rows = {row['module_id']: row for row in result['modules']}
        assert list(rows)[:3] == ['edo', 'aichat', 'mingkh']
        assert rows['edo'] == {'rank': 1, 'module_id': 'edo', 'label': MODULE_NAMES['edo'], 'visits': 3,
                               'unique_users': 2, 'page_views': 3, 'actions': 1, 'last_seen_at': '2026-10-02T10:40:00Z'}
        assert rows['aichat']['visits'] == 2
        assert rows['aichat']['actions'] == 2
        assert rows['aichat']['unique_users'] == 1
        assert activity.PLATFORM not in rows and 'unknown-old-module' not in rows
        zero_modules = [module for module in activity.SECTION_NAMES if module not in {'edo', 'aichat', 'mingkh'}]
        assert list(rows)[3:] == zero_modules
        assert [row['rank'] for row in result['modules']] == list(range(1, len(activity.SECTION_NAMES) + 1))
        assert [activity.account_activity(user['id']) for user in self.users] == before
        assert set(result) == {'period', 'tracking_since', 'visit_gap_minutes', 'updated_at', 'modules'}
        assert all(set(row) == {'rank', 'module_id', 'label', 'visits', 'unique_users', 'page_views', 'actions', 'last_seen_at'} for row in rows.values())

    def test_ties_are_catalog_order_not_last_seen_or_unique_users(self):
        for module, minute in [('users', 0), ('edo', 5), ('mingkh', 10)]:
            activity.record_activity(self.users[0], module, 'page_view', now=self.start + timedelta(minutes=minute))
        result = activity.module_popularity()
        assert [row['module_id'] for row in result['modules']][:3] == ['mingkh', 'edo', 'users']

    def test_read_endpoint_is_protected_and_does_not_change_activity(self):
        from routers.users import router
        identity = {'user': self.users[0]}
        app = FastAPI()
        @app.middleware('http')
        async def identity_middleware(request, call_next):
            request.state.user = identity['user']
            return await call_next(request)
        app.include_router(router)
        with TestClient(app) as client:
            assert client.get('/api/users/module-popularity').status_code == 403
            identity['user'] = {**self.users[0], 'role': 'admin'}
            assert client.get('/api/users/module-popularity').status_code == 403
            identity['user'] = {'id': 3, 'username': 'manager'}
            for _ in range(2):
                response = client.get('/api/users/module-popularity')
                assert response.status_code == 200
                assert response.headers['cache-control'] == 'no-store'
                assert len(response.json()['modules']) == len(activity.SECTION_NAMES)
                assert 'private-hash' not in response.text and 'private.example' not in response.text
                assert 'first-member' not in response.text
            assert client.get('/api/users/1/activity').status_code == 200
            with db.get_db_connection() as conn:
                conn.execute('INSERT INTO account_managers(user_id) VALUES (1)')
            identity['user'] = self.users[0]
            assert client.get('/api/users/module-popularity').status_code == 200
            with db.get_db_connection() as conn:
                conn.execute('DELETE FROM account_managers WHERE user_id=1')
            assert client.get('/api/users/module-popularity').status_code == 403
            identity['user'] = None
            assert client.get('/api/users/module-popularity').status_code == 403
        assert activity.account_summaries() == {}
        request = Request({'type': 'http', 'method': 'GET', 'path': '/api/users/module-popularity', 'headers': []})
        assert activity._event(request, JSONResponse({}), {}) is None
