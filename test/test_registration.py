"""Registration works without SMTP and keeps module access under manager control."""
import importlib
import json
import sqlite3
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from unittest.mock import patch


class RegistrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import utils.db as db
        cls.temp = tempfile.TemporaryDirectory()
        cls.old_db, cls.old_ready = db.DB_FILE, db._SCHEMA_READY
        db.DB_FILE = Path(cls.temp.name) / "registration.sqlite"
        db._SCHEMA_READY = False
        import services.scheduler as scheduler
        with patch.object(scheduler, "start"):
            cls.module = importlib.import_module("app")
        from services.auth.registration import ensure_registration_tables
        ensure_registration_tables()
        cls.db = db
        cls.secret = patch("services.auth.security.get_jwt_secret", return_value="isolated-registration-secret-123456")
        cls.secret.start()

    @classmethod
    def tearDownClass(cls):
        cls.secret.stop()
        cls.db.DB_FILE, cls.db._SCHEMA_READY = cls.old_db, cls.old_ready
        cls.temp.cleanup()

    def setUp(self):
        from fastapi.testclient import TestClient
        with self.db.get_db_connection() as conn:
            conn.execute("UPDATE account_control SET manager_user_id=NULL, grants_migrated=1")
            conn.execute("DELETE FROM users")
            conn.execute("DELETE FROM user_settings")
            conn.execute("DELETE FROM account_notifications")
        self.client = TestClient(self.module.app)
        self.payload = {"username": "new_person", "email": "person@example.ru", "password": "Register-test-123"}

    def tearDown(self):
        self.client.close()

    def test_registration_immediately_creates_account_and_session_without_email(self):
        from services.auth.security import verify_password
        from services.auth.registration import get_settings
        payload = {**self.payload, "username": " new_person ", "email": " Person@Example.RU ",
                   "role": "Администратор", "modules": ["edds"], "can_manage_users": True}
        with patch("services.auth.mailer.send_verification_code") as send_mail:
            response = self.client.post("/api/register", json=payload)
            send_mail.assert_not_called()
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.json()["ok"])
        self.assertEqual(response.json()["redirect"], "/")
        self.assertNotIn("code", response.json())
        self.assertNotIn("debug_code", response.json())
        self.assertIn("access_token", self.client.cookies)
        self.assertIn("HttpOnly", response.headers["set-cookie"])
        self.assertIn("SameSite=strict", response.headers["set-cookie"])
        with self.db.get_db_connection() as conn:
            user = conn.execute("SELECT * FROM users WHERE username='new_person'").fetchone()
            self.assertEqual(user["email"], "person@example.ru")
            self.assertEqual(user["role"], "Пользователь")
            self.assertEqual(json.loads(user["modules"]), [])
            self.assertEqual(user["is_active"], 1)
            self.assertTrue(verify_password(self.payload["password"], user["password_hash"]))
            self.assertNotEqual(user["password_hash"], self.payload["password"])
            self.assertIsNotNone(conn.execute("SELECT 1 FROM user_settings WHERE user_id=?", (user["id"],)).fetchone())
            self.assertEqual(get_settings(user["id"]), {})
            notice = conn.execute("SELECT message FROM account_notifications").fetchone()[0]
            self.assertIn("new_person", notice)
        self.assertEqual(self.client.get("/").status_code, 200)
        self.assertEqual(self.client.get("/edds").status_code, 403)
        self.assertEqual(self.client.get("/api/users").status_code, 403)
        self.client.get("/logout", follow_redirects=False)
        login = self.client.post("/login", data={
            "username": self.payload["username"], "password": self.payload["password"],
        }, follow_redirects=False)
        self.assertEqual(login.status_code, 303)
        self.assertEqual(login.headers["location"], "/")

    def test_duplicate_username_and_email_are_rejected_without_changing_account(self):
        self.assertTrue(self.client.post("/api/register", json=self.payload).json()["ok"])
        self.client.cookies.clear()
        for payload, message in [
            ({**self.payload, "username": "NEW_PERSON", "email": "other@example.ru"}, "логин"),
            ({**self.payload, "username": "different", "email": "PERSON@EXAMPLE.RU"}, "e-mail"),
        ]:
            with self.subTest(payload=payload):
                response = self.client.post("/api/register", json=payload)
                self.assertEqual(response.status_code, 400)
                self.assertIn(message, response.json()["message"])
                self.assertNotIn("access_token", response.cookies)
        with self.db.get_db_connection() as conn:
            self.assertEqual(conn.execute("SELECT count(*) FROM users").fetchone()[0], 1)
            self.assertEqual(conn.execute("SELECT count(*) FROM account_notifications").fetchone()[0], 1)

    def test_invalid_data_does_not_create_account_or_session(self):
        for change in [
            {"username": "ab"}, {"username": "name with spaces"}, {"username": "x" * 33},
            {"email": "bad-email"}, {"email": ""}, {"password": "12345"},
            {"password": "я" * 37}, {"username": ["name"]}, {"password": 123456},
        ]:
            with self.subTest(change=change):
                response = self.client.post("/api/register", json={**self.payload, **change})
                self.assertEqual(response.status_code, 400)
                self.assertFalse(response.json()["ok"])
                self.assertNotIn("access_token", response.cookies)
        with self.db.get_db_connection() as conn:
            self.assertEqual(conn.execute("SELECT count(*) FROM users").fetchone()[0], 0)

    def test_form_has_one_step_and_verification_endpoints_are_removed(self):
        page = self.client.get("/register")
        self.assertEqual(page.status_code, 200)
        self.assertIn('id="registerForm"', page.text)
        self.assertIn('type="submit"', page.text)
        self.assertNotIn("regCode", page.text)
        self.assertNotIn("btnResend", page.text)
        for endpoint in ("verify", "resend"):
            response = self.client.post("/api/register/" + endpoint, json={"username": "new_person", "code": "123456"})
            self.assertEqual(response.status_code, 404)
            self.assertNotIn("access_token", response.cookies)

    def test_cross_site_registration_is_still_rejected(self):
        response = self.client.post("/api/register", json=self.payload, headers={"Origin": "https://other.invalid"})
        self.assertEqual(response.status_code, 403)
        with self.db.get_db_connection() as conn:
            self.assertEqual(conn.execute("SELECT count(*) FROM users").fetchone()[0], 0)

    def test_concurrent_registrations_cannot_share_email(self):
        from services.auth.registration import register_user
        with ThreadPoolExecutor(max_workers=2) as pool:
            futures = [pool.submit(register_user, name, "same@example.ru", "test-hash")
                       for name in ("first_person", "second_person")]
            results = [future.result() for future in futures]
        self.assertEqual(sum(ok for ok, _ in results), 1)
        with self.db.get_db_connection() as conn:
            self.assertEqual(conn.execute("SELECT count(*) FROM users").fetchone()[0], 1)

    def test_settings_failure_rolls_back_account_and_notification(self):
        from services.auth.registration import register_user
        with self.db.get_db_connection() as conn:
            conn.execute("""CREATE TRIGGER fail_settings BEFORE INSERT ON user_settings
                            BEGIN SELECT RAISE(ABORT, 'settings failure'); END""")
        try:
            with self.assertRaises(sqlite3.IntegrityError):
                register_user("new_person", "person@example.ru", "test-hash")
            with self.db.get_db_connection() as conn:
                self.assertEqual(conn.execute("SELECT count(*) FROM users").fetchone()[0], 0)
                self.assertEqual(conn.execute("SELECT count(*) FROM account_notifications").fetchone()[0], 0)
        finally:
            with self.db.get_db_connection() as conn:
                conn.execute("DROP TRIGGER fail_settings")
