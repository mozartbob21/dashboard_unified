"""Synthetic imports and protected routes; never call the live complaint portal."""
import gzip
import io
import json
import zipfile
from datetime import date
from types import SimpleNamespace
from unittest.mock import patch
from fastapi import FastAPI
from fastapi.testclient import TestClient
from services.mingkh import water_map as maps
from routers.mingkh import router

def point(cid=7, **overrides):
    return {'id': cid, 'created': '2026-09-29 09:50:40', 'omsu': 'Тестовый округ', 'kind': 'hvs', 'address': '<b>адрес</b>', 'fact': 'Восстановить работу внешней системы водоснабжения', 'status': 'В работе', 'lat': 55.75, 'lon': 37.61, 'geo': 'ok', 'text': 'This private raw narrative must not be exported', 'author': 'Private Name', **overrides}

def source(*points):
    return json.dumps({'updated': '29.09.2026 15:05', 'points': {str(p['id']): p for p in points}}, ensure_ascii=False).encode()

def archive(member, data):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w') as z:
        z.writestr(member, data)
        z.writestr('scripts/never-run.py', 'raise Exception("Archive script executed")')
    return stream.getvalue()

def client_for(user):
    app = FastAPI()

    @app.middleware('http')
    async def user_state(request, call_next):
        request.state.user = user
        return await call_next(request)
    app.include_router(router)
    return TestClient(app)
import tempfile
import unittest
from pathlib import Path

class WaterMapTests(unittest.TestCase):

    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.tmp_path = Path(temp.name)
        self.setattr(maps, 'SEED_FILE', self.tmp_path / 'no-seed.json.gz')

    def setattr(self, owner, name, value):
        patcher = patch.object(owner, name, value)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_import_keeps_only_map_fields_and_recomputes_geo(self):
        tmp_path = self.tmp_path
        monkeypatch = self
        monkeypatch.setattr(maps, 'DATA_FILE', tmp_path / 'water-map.json.gz')
        result = maps.import_upload(archive('Жалобы/water_points.json', source(point(), point(8, lon=45), point(9, status='Опубликовано'))))
        assert result['count'] == 2 and result['omitted'] == 1
        stored = maps.load()
        assert stored['rows'][0][9] == 0 and stored['rows'][1][9] == 1
        assert stored['meta']['from'] == '2026-09-29'
        assert stored['meta']['archive_updated'] == '29.09.2026 15:05'
        raw = gzip.decompress(maps.DATA_FILE.read_bytes())
        assert b'private raw narrative' not in raw and b'Private Name' not in raw
        assert stored['rows'][0][4] == '<b>адрес</b>'
        assert not list(tmp_path.glob('*.py'))

    def test_bad_import_cannot_replace_good_dataset(self):
        tmp_path = self.tmp_path
        monkeypatch = self
        for content in [b'not json', b'[]', b'{}', archive('../water_points.json', source(point())), archive('wrong.json', source(point())), source(point(lat=float('nan'))), source(point(id=True)), source(point(kind='unknown')), source(point(created='bad'))]:
            with self.subTest(content=content[:25]):
                monkeypatch.setattr(maps, 'DATA_FILE', tmp_path / 'water-map.json.gz')
                maps.import_upload(source(point()))
                original = maps.DATA_FILE.read_bytes()
                with self.assertRaises(ValueError):
                    maps.import_upload(content)
                assert maps.DATA_FILE.read_bytes() == original

    def test_duplicate_members_and_declared_oversize_are_rejected(self):
        monkeypatch = self
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, 'w') as z:
            z.writestr('a/water_points.json', source(point()))
            z.writestr('b/water_points.json', source(point()))
        with self.assertRaisesRegex(ValueError, 'ровно один'):
            maps.parse_upload(stream.getvalue())
        monkeypatch.setattr(maps, 'MAX_BYTES', 10)
        with self.assertRaisesRegex(ValueError, '100 МБ'):
            maps.parse_upload(source(point()))

    def test_routes_require_module_and_import_requires_manager(self):
        tmp_path = self.tmp_path
        monkeypatch = self
        monkeypatch.setattr(maps, 'DATA_FILE', tmp_path / 'water-map.json.gz')
        maps.import_upload(source(point()))
        for user in (None, {'modules': []}):
            client = client_for(user)
            for path in ('/mingkh/water-map', '/mingkh/api/water-map', '/mingkh/api/water-map/status'):
                assert client.get(path).status_code == 403
            assert client.post('/mingkh/api/water-map/import', content=source(point())).status_code == 403
        client = client_for({'modules': ['mingkh']})
        response = client.get('/mingkh/api/water-map')
        assert response.status_code == 200 and response.json()['meta']['count'] == 1
        assert response.headers['cache-control'] == 'private, no-store'
        assert response.headers['content-encoding'] == 'gzip'
        assert client.get('/mingkh/api/water-map', headers={'Accept-Encoding': 'identity'}).json()['meta']['count'] == 1
        with patch('services.auth.accounts.is_account_manager', return_value=False):
            assert client.post('/mingkh/api/water-map/import', content=source(point())).status_code == 403
            assert client.post('/mingkh/api/water-map/refresh').status_code == 403
        with patch('services.auth.accounts.is_account_manager', return_value=True):
            assert client.post('/mingkh/api/water-map/import', content=source(point(8))).json()['count'] == 1
            assert client.post('/mingkh/api/water-map/import', content=b'bad').status_code == 400
        assert client.get('/mingkh/api/water-map').json()['rows'][0][0] == 8

    def test_live_refresh_preserves_history_and_coordinates_and_uses_portal_filters(self):
        tmp_path = self.tmp_path
        monkeypatch = self
        monkeypatch.setattr(maps, 'DATA_FILE', tmp_path / 'water-map.json.gz')
        maps.import_upload(source(point(), point(6, created='2024-01-01 10:00:00')))
        calls = []

        class Portal:

            def fetch_all(self, filters):
                calls.append(filters)
                assert filters['filters.curators'] == maps.CURATOR
                assert filters['filters.statuses'] == maps.STATUSES
                if filters['filters.createdBefore'] < '2026-09-29':
                    return []
                return [{'cardId': 7, 'created': '29.09.2026 09:50:40', 'district': 'Тестовый округ', 'ecurFact': 'Восстановить работу внешней системы водоснабжения', 'address': 'Новый адрес', 'status': 'Решено'}, {'cardId': 8, 'created': '30.09.2026 09:50:40', 'district': 'Тестовый округ', 'ecurFact': 'Восстановить работу внешней системы водоснабжения', 'status': 'В работе'}]

            def request(self, method, path, params):
                assert params['id'] == 8
                return SimpleNamespace(text='var coords = [55.7, 37.6];')
        result = maps.refresh(Portal(), date(2026, 9, 30))
        stored = maps.load()
        assert result['count'] == 3 and (not result['archive'])
        assert result['archive_updated'] == '29.09.2026 15:05'
        assert result['refreshed_to'] == '2026-09-30'
        assert {r[0] for r in stored['rows']} == {6, 7, 8}
        assert stored['status'][next((r for r in stored['rows'] if r[0] == 7))[6]] == 'Решено'

    def test_failed_refresh_keeps_previous_archive(self):
        tmp_path = self.tmp_path
        monkeypatch = self
        monkeypatch.setattr(maps, 'DATA_FILE', tmp_path / 'water-map.json.gz')
        maps.import_upload(source(point()))
        original = maps.DATA_FILE.read_bytes()

        class BrokenPortal:

            def fetch_all(self, filters):
                raise ValueError('Unavailable')
        with self.assertRaises(ValueError):
            maps.refresh(BrokenPortal(), date(2026, 9, 30))
        assert maps.DATA_FILE.read_bytes() == original

    def test_busy_import_does_not_overwrite(self):
        tmp_path = self.tmp_path
        monkeypatch = self
        monkeypatch.setattr(maps, 'DATA_FILE', tmp_path / 'water-map.json.gz')
        with maps.LOCK:
            with self.assertRaises(maps.StoreBusy):
                maps.import_upload(source(point()))
        assert not maps.DATA_FILE.exists()

    def test_seed_initializes_once_and_never_overwrites_server_updates(self):
        self.setattr(maps, 'DATA_FILE', self.tmp_path/'water-map.json.gz')
        seed = self.tmp_path/'seed.json.gz'
        self.setattr(maps, 'SEED_FILE', seed)
        seed.write_bytes(gzip.compress(json.dumps(maps.parse_upload(source(point()))).encode()))
        self.assertTrue(maps.ensure_seed())
        self.assertEqual(maps.load()['rows'][0][0], 7)
        maps.import_upload(source(point(8)))
        self.assertFalse(maps.ensure_seed())
        self.assertEqual(maps.load()['rows'][0][0], 8)

    def test_refresh_limits_leave_archive_unchanged(self):
        self.setattr(maps, 'DATA_FILE', self.tmp_path/'water-map.json.gz')
        maps.import_upload(source(point()))
        before = maps.DATA_FILE.read_bytes()
        class UnusedPortal:
            def fetch_all(self, filters):
                raise AssertionError('Expired refresh must not query the portal')
        self.setattr(maps, 'MAX_REFRESH_SECONDS', -1)
        with self.assertRaisesRegex(ValueError, '20 минут'):
            maps.refresh(UnusedPortal(), date(2026, 9, 30))
        self.assertEqual(maps.DATA_FILE.read_bytes(), before)
