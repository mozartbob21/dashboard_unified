"""Run with python -m unittest test.test_access_control; uses an isolated SQLite DB."""
import importlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


class AccessControlTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import utils.db as db
        cls.temp = tempfile.TemporaryDirectory()
        cls.old_db, cls.old_ready = db.DB_FILE, db._SCHEMA_READY
        db.DB_FILE = Path(cls.temp.name) / "tests.sqlite"
        db._SCHEMA_READY = False
        import services.scheduler as scheduler
        with patch.object(scheduler, "start"):
            cls.module = importlib.import_module("app")
        cls.db, cls.scheduler = db, scheduler
        cls.secret = patch("services.auth.security.get_jwt_secret", return_value="isolated-test-secret-1234567890")
        cls.secret.start()

    @classmethod
    def tearDownClass(cls):
        cls.secret.stop()
        cls.db.DB_FILE, cls.db._SCHEMA_READY = cls.old_db, cls.old_ready
        cls.temp.cleanup()

    def setUp(self):
        from routers.auth import LOGIN_ATTEMPTS
        LOGIN_ATTEMPTS.clear()  # Each isolated test has its own login attempt window.
        from fastapi.testclient import TestClient
        from services.auth.accounts import initialize_access_control, provision_manager
        from services.auth.security import hash_password
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE account_control SET manager_user_id=NULL, grants_migrated=0")
            conn.execute("DELETE FROM users")
            conn.execute("DELETE FROM scheduler_jobs")
            conn.execute("DELETE FROM run_history")
            conn.execute("INSERT INTO users(username,password_hash,role,modules) VALUES (?,?,?,?)",
                         ("ordinary",hash_password("Test-password-123"),"Контроль данных",'["edo"]'))
            conn.execute("INSERT INTO users(username,password_hash,role,modules) VALUES (?,?,?,?)",
                         ("legacy",hash_password("Test-password-123"),"Пользователь",'[]'))
        self.scheduler._SEEDED = False
        initialize_access_control()
        self.credentials = provision_manager()
        self.client = TestClient(self.module.app)

    def tearDown(self):
        self.client.close()

    def login(self, username="ordinary", password="Test-password-123"):
        response = self.client.post('/login',data={"username":username,"password":password},follow_redirects=False)
        self.assertEqual(response.status_code,303)
        self.assertNotIn('error=',response.headers['location'])
        return response

    def manager_login(self):
        return self.login(self.credentials['username'],self.credentials['password'])

    def test_manager_only_even_for_admin_role(self):
        self.login()
        self.assertEqual(self.client.get('/users').status_code,403)
        self.assertEqual(self.client.get('/api/users').status_code,403)
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE users SET role='Администратор' WHERE username='ordinary'")
        self.assertEqual(self.client.get('/api/users').status_code,403)
        self.assertEqual(self.client.post('/api/users',json={"username":"blocked"}).status_code,403)
        self.assertEqual(self.manager_login().headers['location'],'/users')
        self.assertEqual(self.client.get('/users').status_code,200)
        listing=self.client.get('/api/users').json()
        self.assertNotIn('password_hash',str(listing))
        self.assertNotIn(self.credentials['password'],str(listing))

    def test_expired_session_is_not_a_form_error(self):
        from services.auth.security import create_access_token
        self.client.cookies.set('access_token',create_access_token({'sub':'user_manager'},expires_hours=-1))
        response=self.client.put('/api/users/3',json={'username':'ordinary'})
        self.assertEqual(response.status_code,401)
        self.assertEqual(response.json()['code'],'auth_required')
        self.assertIn('Сессия',response.json()['detail'])
        self.assertEqual(self.client.get('/api/users/notifications').status_code,401)

    def test_notifications_are_manager_only_and_can_be_read(self):
        self.login()
        self.assertEqual(self.client.get('/api/users/notifications').status_code,403)
        self.manager_login()
        payload={'username':'notified','password':'New-password-123','modules':[]}
        self.assertEqual(self.client.post('/api/users',json=payload).status_code,201)
        user=next(u for u in self.client.get('/api/users').json()['users'] if u['username']=='notified')
        payload.update(password='',modules=['edo'])
        self.assertEqual(self.client.put(f"/api/users/{user['id']}",json=payload).status_code,200)
        payload['modules']=['unknown']
        self.assertEqual(self.client.put(f"/api/users/{user['id']}",json=payload).status_code,400)
        data=self.client.get('/api/users/notifications').json()
        messages=' '.join(n['message'] for n in data['notifications'])
        self.assertIn('Создана учётная запись «notified»',messages)
        self.assertIn('изменён доступ',messages)
        self.assertIn('HTTP 400',messages)
        self.assertNotIn('New-password-123',messages)
        self.assertGreater(data['unread'],0)
        self.assertEqual(self.client.post('/api/users/notifications/read',json={'through_id':data['notifications'][0]['id']}).status_code,200)
        self.assertEqual(self.client.get('/api/users/notifications').json()['unread'],0)

    def test_validation_reports_field_and_blank_password_keeps_hash(self):
        self.manager_login()
        response=self.client.post('/api/users',json={'username':'ab'})
        self.assertEqual(response.status_code,422)
        self.assertIn('username',response.json()['detail'][0]['loc'])
        user=next(u for u in self.client.get('/api/users').json()['users'] if u['username']=='ordinary')
        self.assertEqual(self.client.put(f"/api/users/{user['id']}",json={'username':'ordinary','password':'','email':'example@yandex.ru','modules':['edo']}).status_code,200)
        self.login()

    def test_individual_grants_and_revocation_apply_to_existing_session(self):
        self.manager_login()
        payload={"username":"newperson","password":"New-password-123","modules":["zips"]}
        self.assertEqual(self.client.post('/api/users',json=payload).status_code,201)
        user=next(u for u in self.client.get('/api/users').json()['users'] if u['username']=='newperson')
        self.login('newperson','New-password-123')
        cookie=self.client.cookies.get('access_token')
        self.assertEqual(self.client.get('/zip_curator').status_code,200)
        self.assertEqual(self.client.get('/overdue').status_code,403)
        self.manager_login()
        payload.update(modules=['edo'],password='')
        self.assertEqual(self.client.put(f"/api/users/{user['id']}",json=payload).status_code,200)
        self.client.cookies.set('access_token',cookie,domain='testserver.local',path='/')
        self.assertEqual(self.client.get('/zip_curator').status_code,403)
        page=self.client.get('/').text
        self.assertIn('data-module="edo"',page)
        self.assertNotIn('data-module="zips"',page)

    def test_common_pages_and_aliases_are_filtered(self):
        self.login()
        self.assertEqual(self.client.get('/aichat').status_code,200)
        self.assertEqual(self.client.get('/scheduler').status_code,200)
        jobs=self.client.get('/api/scheduler/jobs').json()['jobs']
        self.assertEqual([j['module_id'] for j in jobs],['edo'])
        for path in ['/overdue/run-status','/api/cds/last-result','/zip_curator_original.html','/water-rm/has-key','/tools','/municipality-report.pdf','/data/dashboard.db','/data/cameras/missing.json']:
            self.assertEqual(self.client.get(path).status_code,403,path)
        with patch.object(self.scheduler,'_launch_module') as launch:
            self.assertEqual(self.client.post('/api/scheduler/jobs/cameras/run-now').status_code,403)
            launch.assert_not_called()
            self.assertEqual(self.client.post('/api/scheduler/jobs/edo/run-now').status_code,200)
            launch.assert_called_once_with('edo')
        for suffix in ['enable','disable','interval']:
            self.assertEqual(self.client.post('/api/scheduler/jobs/overdue/'+suffix,json={'interval_minutes':60}).status_code,403)

    def test_history_filter_precedes_limit_and_includes_camera_alias(self):
        with self.db.get_db_connection() as conn:
            for i,module in enumerate(['edo','overdue','camera_prescriptions']):
                conn.execute("INSERT INTO run_history (run_id,module_id,started_at,status) VALUES (?,?,?,'success')",
                             (str(i),module,f'2026-09-12T10:0{i}:00+00:00'))
        self.login()
        for path in ['/api/history/recent?limit=1','/api/history/all?limit=1']:
            self.assertEqual([r['module_id'] for r in self.client.get(path).json()],['edo'])
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[\"cameras\"]' WHERE username='ordinary'")
        self.assertEqual(self.client.get('/api/history/recent').json()[0]['module_id'],'camera_prescriptions')
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[]' WHERE username='ordinary'")
        self.assertEqual(self.client.get('/api/history/all').json(),[])
        self.assertEqual(self.client.get('/api/scheduler/jobs').json()['jobs'],[])

    def test_disable_user_and_validate_module_ids(self):
        self.manager_login()
        user=next(u for u in self.client.get('/api/users').json()['users'] if u['username']=='ordinary')
        payload={'username':'ordinary','modules':['not-a-module']}
        self.assertEqual(self.client.put(f"/api/users/{user['id']}",json=payload).status_code,400)
        payload.update(modules=[],is_active=False)
        self.assertEqual(self.client.put(f"/api/users/{user['id']}",json=payload).status_code,200)
        self.client.cookies.clear()
        response=self.client.post('/login',data={'username':'ordinary','password':'Test-password-123'},follow_redirects=False)
        self.assertIn('error=',response.headers['location'])

    def test_manager_cannot_be_disabled_and_provisioning_is_idempotent(self):
        from services.auth.accounts import provision_manager
        self.assertFalse(provision_manager()['created'])
        self.manager_login()
        user=next(u for u in self.client.get('/api/users').json()['users'] if u['is_manager'])
        self.assertEqual(self.client.put(f"/api/users/{user['id']}",json={'username':user['username'],'is_active':False}).status_code,400)
        self.assertEqual(self.client.put(f"/api/users/{user['id']}",json={'username':user['username'],'modules':['edo']}).status_code,400)

    def test_legacy_migration_is_once_and_registration_starts_without_grants(self):
        from services.auth.accounts import initialize_access_control
        from services.auth.registration import DEFAULT_MODULES
        from core.roles import ALL_MODULE_IDS
        with self.db.get_db_connection() as conn:
            row=conn.execute("SELECT modules FROM users WHERE username='legacy'").fetchone()
            self.assertEqual(json.loads(row['modules']),ALL_MODULE_IDS)
            conn.execute("UPDATE users SET modules='[\"edo\"]' WHERE username='legacy'")
        initialize_access_control()
        with self.db.get_db_connection() as conn:
            self.assertEqual(json.loads(conn.execute("SELECT modules FROM users WHERE username='legacy'").fetchone()[0]),['edo'])
        self.assertEqual(DEFAULT_MODULES,[])


if __name__ == '__main__':
    unittest.main()
