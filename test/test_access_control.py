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
        # Discovery may have imported auth against another DB before this fixture.
        from services.auth.registration import ensure_registration_tables
        ensure_registration_tables()
        cls.db, cls.scheduler = db, scheduler
        cls.secret = patch("services.auth.security.get_jwt_secret", return_value="isolated-test-secret-1234567890")
        cls.secret.start()
        cls.integration_key = patch('services.auth.integrations.KEY_FILE',Path(cls.temp.name)/'integrations.key')
        cls.integration_key.start()

    @classmethod
    def tearDownClass(cls):
        cls.secret.stop()
        cls.integration_key.stop()
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
            conn.execute("DELETE FROM integration_credentials")
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

    def test_login_form_preserves_origin_for_browser_submission(self):
        from fastapi.testclient import TestClient
        # Browser form POSTs send Origin: null under no-referrer, even to this app.
        # Exercise the policy and headers that browser navigation actually uses.
        for origin in ['http://127.0.0.1:8000', 'http://192.168.1.20:8000', 'https://office.test:8443']:
            with self.subTest(origin=origin), TestClient(self.module.app, base_url=origin) as client:
                page = client.get('/login')
                self.assertEqual(page.headers['Referrer-Policy'], 'same-origin')
                response = client.post('/login', data={
                    'username': 'ordinary', 'password': 'Test-password-123',
                }, headers={'Origin': origin, 'Sec-Fetch-Site': 'same-origin'}, follow_redirects=False)
                self.assertEqual(response.status_code, 303)
                self.assertEqual(response.headers['location'], '/')
                self.assertIn('access_token', client.cookies)
                self.assertEqual(client.get('/').status_code, 200)

    def test_rejected_login_shows_readable_form_without_authenticating(self):
        for headers in [
            {'Origin': 'https://example.invalid'},
            {'Origin': 'null', 'Sec-Fetch-Site': 'same-origin'},
            {'Origin': 'http://testserver:8001'},
            {'Origin': 'http://testserver', 'Sec-Fetch-Site': 'cross-site'},
        ]:
            with self.subTest(headers=headers), patch('routers.auth.authenticate_user') as authenticate:
                response = self.client.post('/login', data={
                    'username': 'ordinary', 'password': 'Test-password-123',
                }, headers=headers, follow_redirects=False)
                self.assertEqual(response.status_code, 403)
                self.assertEqual(response.headers['content-type'], 'text/html; charset=utf-8')
                self.assertIn('Введите логин и пароль ещё раз', response.text)
                self.assertIn('class="login-form"', response.text)
                self.assertNotIn('Test-password-123', response.text)
                self.assertNotIn('access_token', response.cookies)
                authenticate.assert_not_called()

    def test_invalid_password_still_returns_readable_login_error(self):
        response = self.client.post('/login', data={
            'username': 'ordinary', 'password': 'wrong-password',
        }, headers={'Origin': 'http://testserver', 'Sec-Fetch-Site': 'same-origin'})
        self.assertEqual(response.status_code, 200)
        self.assertIn('Неверный логин или пароль', response.text)
        self.assertNotIn('access_token', self.client.cookies)

    def test_home_and_favorites_use_real_local_account_identity(self):
        from services.auth.security import find_user_by_username, load_users
        from bs4 import BeautifulSoup
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules=? WHERE username='ordinary'",
                         (json.dumps(['edo', 'overdue', 'edds', 'zips']),))
            user_id = conn.execute("SELECT id FROM users WHERE username='ordinary'").fetchone()['id']
        self.login()
        # Exercise the real cookie -> SQLite -> user -> home chain, not a fake user dict.
        response = self.client.get('/', follow_redirects=False)
        self.assertEqual(response.status_code, 200)
        self.assertIn('text/html', response.headers['content-type'])
        self.assertIn('С возвращением!', response.text)
        self.assertEqual(find_user_by_username('ORDINARY')['id'], user_id)
        self.assertEqual(next(u for u in load_users() if u['username']=='ordinary')['id'], user_id)
        saved = self.client.put('/api/me/home-favorites', json={'favorites': ['zips', 'edo']})
        self.assertEqual(saved.status_code, 200)
        page = BeautifulSoup(self.client.get('/').text, 'html.parser')
        self.assertEqual(json.loads(page.find(id='homePreferences').string)['favorites'], ['zips', 'edo'])
        self.assertEqual(self.client.get('/api/me/settings').status_code, 200)
        self.client.cookies.clear()
        self.login('legacy')
        self.assertEqual(self.client.get('/api/me/home-favorites').json()['favorites'], [])

    def test_home_accepts_existing_username_token_and_still_rejects_disabled_account(self):
        from services.auth.security import create_access_token
        # Existing sessions carry sub/role only; no new token or login is required.
        token = create_access_token({'sub': 'ordinary', 'role': 'Контроль данных'})
        self.client.cookies.set('access_token', token)
        self.assertEqual(self.client.get('/', follow_redirects=False).status_code, 200)
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE users SET is_active=0 WHERE username='ordinary'")
        response = self.client.get('/', follow_redirects=False)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers['location'], '/login')
        self.client.cookies.clear()
        self.assertEqual(self.client.get('/', follow_redirects=False).status_code, 302)

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

    def test_archive_preserves_data_and_blocks_existing_session(self):
        self.login()
        cookie=self.client.cookies.get('access_token')
        self.manager_login()
        users=self.client.get('/api/users').json()['users']
        ordinary=next(u for u in users if u['username']=='ordinary')
        parent=next(u for u in users if u['is_manager'])
        self.assertEqual(self.client.post(f"/api/users/{parent['id']}/archive").status_code,400)
        self.assertEqual(self.client.post(f"/api/users/{ordinary['id']}/archive").status_code,200)
        archived=next(u for u in self.client.get('/api/users').json()['users'] if u['username']=='ordinary')
        self.assertTrue(archived['archived_at']);self.assertEqual(archived['modules'],['edo'])
        self.assertFalse(archived['is_active'])
        self.assertEqual(self.client.put(f"/api/users/{ordinary['id']}",json={'username':'ordinary','is_active':True}).status_code,400)
        self.client.cookies.set('access_token',cookie,domain='testserver.local',path='/')
        self.assertEqual(self.client.get('/edo',follow_redirects=False).status_code,302)
        self.manager_login()
        self.assertEqual(self.client.post(f"/api/users/{ordinary['id']}/restore").status_code,200)
        restored=next(u for u in self.client.get('/api/users').json()['users'] if u['username']=='ordinary')
        self.assertIsNone(restored['archived_at']);self.assertFalse(restored['is_active'])

    def test_integration_credentials_are_encrypted_and_manager_only(self):
        from services.auth.integrations import credentials
        self.login()
        self.assertEqual(self.client.get('/api/users/integrations/edds').status_code,403)
        self.assertEqual(self.client.put('/api/users/integrations/edds',json={'username':'u','password':'secret'}).status_code,403)
        self.manager_login()
        for service in ['edds','edds_arm','mingkh']:
            url='/api/users/integrations/'+service
            self.assertEqual(self.client.put(url,json={'username':'testlogin','password':'secret-test-only'}).status_code,200)
            data=self.client.get(url).json()
            self.assertTrue(data['configured']);self.assertNotIn('password',data)
            self.assertNotIn('secret-test-only',str(data))
            self.assertEqual(self.client.put(url,json={'username':'testlogin','password':''}).status_code,200)
            self.assertEqual(credentials(service)['password'],'secret-test-only')
        with self.db.get_db_connection() as conn:
            for row in conn.execute('SELECT encrypted_value FROM integration_credentials'):
                self.assertNotIn('secret-test-only',row[0]);self.assertNotIn('testlogin',row[0])

    def test_ai_chat_history_remains_shared_without_implicit_restricted_context(self):
        store=self.module.aichat_store
        chat_dir=Path(self.temp.name)/'shared-chat-test'
        with patch.object(store,'DATA_DIR',chat_dir), patch.object(store,'DIALOGS_FILE',chat_dir/'dialogs.json'):
            self.login()
            dialog=self.client.post('/aichat/api/dialogs',json={'title':'Shared test chat'}).json()
            self.manager_login()
            self.assertEqual(self.client.get('/aichat/api/dialogs/'+dialog['id']).json()['title'],'Shared test chat')
            with patch.object(self.module,'aichat_ask',return_value='Synthetic answer') as ai, patch.object(self.module,'build_platform_context') as context:
                reply=self.client.post('/aichat/api/send',data={'dialog_id':dialog['id'],'text':'Проверь воду в тестовом примере'})
                self.assertEqual(reply.status_code,200)
                context.assert_not_called()
                self.assertIn('Проверь воду',str(ai.call_args))
            self.login()
            ids=[d['id'] for d in self.client.get('/aichat/api/dialogs').json()['items']]
            self.assertIn(dialog['id'],ids)

    def test_mingkh_grant_and_missing_credentials(self):
        self.login()
        self.assertEqual(self.client.get('/mingkh').status_code,403)
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[\"mingkh\"]' WHERE username='ordinary'")
        self.assertEqual(self.client.get('/mingkh').status_code,200)
        result=self.client.get('/mingkh/api/dataset')
        self.assertEqual(result.status_code,503)
        self.assertIn('Администратору',result.json()['error'])
        self.assertIn('data-module="mingkh"',self.client.get('/').text)
        self.assertEqual(self.client.get('/api/users/integrations/mingkh').status_code,403)

    def test_cross_site_writes_and_secret_downloads_are_blocked(self):
        self.login()
        response=self.client.post('/api/home/preferences',json={},headers={'Origin':'https://example.invalid'})
        self.assertEqual(response.status_code,403)
        for path in ['/data/tools/accounts/any/jobs/any/document.pdf','/data/cds/browser-profile/Cookies','/data/cds/password.json']:
            self.assertEqual(self.client.get(path).status_code,403)
        csp=self.client.get('/aichat').headers['Content-Security-Policy']
        self.assertIn("connect-src 'self'",csp)
        self.assertEqual(self.client.get('/aichat').headers['Cache-Control'],'no-store')

    def test_edds_grant_page_and_refresh_without_credentials(self):
        self.login()
        for path in ['/edds','/edds/water-daily','/edds/status']:
            self.assertEqual(self.client.get(path).status_code,403)
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[\"edds\"]' WHERE username='ordinary'")
        page=self.client.get('/edds')
        self.assertEqual(page.status_code,200);self.assertIn('eddsRefresh',page.text)
        self.assertIn("const ZH_URL = EMBEDDED ? '/edds/water-daily'",page.text)
        self.assertIn('days',self.client.get('/edds/water-daily').json())
        self.assertEqual(self.client.post('/edds/refresh').status_code,400)
        home=self.client.get('/').text
        self.assertIn('data-module="edds"',home)
        self.assertNotIn('data-module="water_rm"',home)
        self.assertIn('data-module-view="list"',home)

    def test_delegated_manager_permissions_and_live_revocation(self):
        self.manager_login()
        listing=self.client.get('/api/users').json()
        self.assertTrue(listing['is_parent_manager'])
        ordinary=next(u for u in listing['users'] if u['username']=='ordinary')
        parent=next(u for u in listing['users'] if u['is_manager'])
        payload={'username':'ordinary','modules':['edo'],'can_manage_users':True}
        self.assertEqual(self.client.put(f"/api/users/{ordinary['id']}",json=payload).status_code,200)
        self.assertEqual(self.login().headers['location'],'/users')
        cookie=self.client.cookies.get('access_token')
        self.assertEqual(self.client.get('/users').status_code,200)
        self.assertFalse(self.client.get('/api/users').json()['is_parent_manager'])
        self.assertIn('href="/users"',self.client.get('/').text)
        self.assertEqual(self.client.get('/overdue').status_code,403)
        for target,username in [(parent['id'],parent['username']),(ordinary['id'],'ordinary')]:
            self.assertEqual(self.client.put(f'/api/users/{target}',json={'username':username,'modules':[]}).status_code,403)
        new={'username':'delegatechild','password':'Delegate-test-123','modules':['zips'],'can_manage_users':True}
        self.assertEqual(self.client.post('/api/users',json=new).status_code,403)
        new['can_manage_users']=False
        self.assertEqual(self.client.post('/api/users',json=new).status_code,201)
        child=next(u for u in self.client.get('/api/users').json()['users'] if u['username']=='delegatechild')
        new.update(password='',modules=['edo'])
        self.assertEqual(self.client.put(f"/api/users/{child['id']}",json=new).status_code,200)
        new['can_manage_users']=True
        self.assertEqual(self.client.put(f"/api/users/{child['id']}",json=new).status_code,403)
        self.manager_login()
        self.assertEqual(self.client.put(f"/api/users/{parent['id']}",json={'username':parent['username'],'can_manage_users':False}).status_code,400)
        payload['can_manage_users']=False
        self.assertEqual(self.client.put(f"/api/users/{ordinary['id']}",json=payload).status_code,200)
        self.client.cookies.set('access_token',cookie,domain='testserver.local',path='/')
        self.assertEqual(self.client.get('/api/users').status_code,403)
        self.assertEqual(self.client.get('/api/users/notifications').status_code,403)
        self.assertEqual(self.client.get('/edo').status_code,200)

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
