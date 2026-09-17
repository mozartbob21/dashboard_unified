import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi import FastAPI
from fastapi.testclient import TestClient
from routers import auth
from services.auth import home_preferences as preferences
from utils import db


class FavoritesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        patcher = patch.object(db, 'DB_FILE', Path(self.temp.name) / 'test.sqlite')
        patcher.start(); self.addCleanup(patcher.stop)
        self.user = {'id': 1, 'modules': ['edo', 'overdue', 'edds', 'zips']}
        self.app = FastAPI()
        @self.app.middleware('http')
        async def identity(request, call_next):
            request.state.user = self.user
            return await call_next(request)
        self.app.include_router(auth.router)
        self.client = TestClient(self.app)
        self.url = '/api/me/home-favorites'

    def test_order_is_persistent_and_scoped_to_account(self):
        self.assertEqual(self.client.put(self.url, json={'favorites': ['zips', 'edo', 'zips']}).status_code, 200)
        self.assertEqual(self.client.get(self.url).json()['favorites'], ['zips', 'edo'])
        self.user = {**self.user, 'id': 2}
        self.assertEqual(self.client.get(self.url).json()['favorites'], [])
        self.client.put(self.url, json={'favorites': ['edds'], 'user_id': 1})
        self.user = {**self.user, 'id': 1}
        self.assertEqual(self.client.get(self.url).json()['favorites'], ['zips', 'edo'])

    def test_three_permissions_and_duplicates_do_not_enable_favorites(self):
        self.user['modules'] = ['edo', 'overdue', 'zips', 'zips']
        data = self.client.get(self.url).json()
        self.assertFalse(data['can_favorite'])
        self.assertEqual(data['eligible'], [])
        self.assertEqual(self.client.put(self.url, json={'favorites': ['zips']}).status_code, 403)

    def test_fixed_unauthorized_and_unknown_modules_cannot_be_saved(self):
        self.user['modules'] += ['tools', 'water-dashboard', 'municipality-report']
        for module in ['tools', 'aichat', 'water-dashboard', 'municipality-report', 'telegram', 'unknown']:
            with self.subTest(module=module):
                self.assertEqual(self.client.put(self.url, json={'favorites': [module]}).status_code, 403)
        self.assertEqual(self.client.get(self.url).json()['favorites'], [])

    def test_revoked_permissions_are_filtered_and_restored_after_regrant(self):
        self.client.put(self.url, json={'favorites': ['zips', 'edo']})
        self.user['modules'] = ['edo', 'overdue', 'edds', 'cameras']
        self.assertEqual(self.client.get(self.url).json()['favorites'], ['edo'])
        self.user['modules'] = ['edo', 'overdue', 'edds']
        self.assertEqual(self.client.get(self.url).json()['favorites'], [])
        self.user['modules'] = ['edo', 'overdue', 'edds', 'zips']
        self.assertEqual(self.client.get(self.url).json()['favorites'], ['zips', 'edo'])

    def test_requires_identity_and_valid_list(self):
        self.assertEqual(self.client.put(self.url, json={'favorites': 'zips'}).status_code, 422)
        self.user = None
        self.assertEqual(self.client.get(self.url).status_code, 401)
        self.assertEqual(self.client.put(self.url, json={'favorites': []}).status_code, 401)

    def test_keycloak_identity_is_separate_from_local_and_other_subjects(self):
        kc = {'kc_sub': '1', 'roles': ['edo','overdue','edds','zips'], 'role': 'admin'}
        preferences.preferences(kc, ['edds'])
        self.assertEqual(preferences.preferences(self.user)['favorites'], [])
        self.assertEqual(preferences.preferences({**kc, 'kc_sub': '2'})['favorites'], [])
        self.assertEqual(preferences.preferences(kc)['favorites'], ['edds'])

    def test_response_is_not_cached(self):
        self.assertEqual(self.client.get(self.url).headers['cache-control'], 'no-store')
