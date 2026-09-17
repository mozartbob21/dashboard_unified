import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch

import requests
from fastapi import FastAPI
from fastapi.testclient import TestClient

from routers import edds
from services.edds import arm, runner


class ArmTests(unittest.TestCase):
    def test_csv_handles_quoted_fields_and_bom(self):
        rows = arm.parse_csv('\ufeffid_cds_claim;text_message\n1;"текст; с разделителем"\n', 'id_cds_claim')
        self.assertEqual(rows[1], ['1', 'текст; с разделителем'])

    def test_html_or_unknown_columns_cannot_be_treated_as_data(self):
        for text in ['<html><form>Login</form></html>', 'error;message\n1;denied', '']:
            with self.assertRaises(arm.ArmError):
                arm.parse_csv(text, 'id_cds_claim')

    def test_period_is_validated_before_connecting(self):
        with patch.object(arm, 'credentials') as credentials:
            for start, end in [(date(2026, 9, 17), date(2026, 9, 16)), (date(2024, 1, 1), date(2026, 1, 1))]:
                with self.assertRaises(arm.ArmError):
                    arm.fetch_report(start, end)
            credentials.assert_not_called()

    def test_arm_credentials_are_used_and_connection_closed(self):
        with patch.object(arm, 'credentials', return_value={'username': 'user', 'password': 'secret'}) as credentials, \
                patch.object(arm, 'ArmClient') as client:
            client.return_value.report.return_value = [['id_cds_claim'], ['1']]
            self.assertEqual(len(arm.fetch_report(date(2026, 9, 17), date(2026, 9, 17))), 2)
            credentials.assert_called_once_with('edds_arm')
            client.return_value.login.assert_called_once_with({'username': 'user', 'password': 'secret'})
            client.return_value.close.assert_called_once()

    def test_login_failure_releases_connection_and_lock(self):
        with patch.object(arm, 'credentials', return_value={'username': 'u', 'password': 'secret'}), \
                patch.object(arm, 'ArmClient') as client:
            client.return_value.login.side_effect = arm.ArmError('Failed', 403)
            with self.assertRaises(arm.ArmError):
                arm.fetch_report(date.today(), date.today())
            client.return_value.close.assert_called_once()
            self.assertFalse(arm.LOCK.locked())

    def test_login_uses_hidden_fields_and_post_without_exposing_secret_in_url(self):
        client = arm.ArmClient()
        self.addCleanup(client.close)
        page = '<form action="?act=login"><input type="hidden" name="csrf" value="token"><input name="user"><input type="password" name="pass"><button name="submit" value="1">Войти</button></form>'
        with patch.object(client, 'request', side_effect=[(page, arm.BASE_URL), ('<html>Report</html>', arm.BASE_URL)]) as request:
            client.login({'username': 'alice', 'password': 'private-password'})
            method, url, fields = request.call_args.args
            self.assertEqual(method, 'POST')
            self.assertNotIn('private-password', url)
            self.assertEqual(fields, {'csrf': 'token', 'submit': '1', 'user': 'alice', 'pass': 'private-password'})

    def test_wrong_password_returns_clear_error(self):
        client = arm.ArmClient()
        self.addCleanup(client.close)
        page = '<form><input name="user"><input type="password" name="pass"></form>'
        with patch.object(client, 'request', return_value=(page, arm.BASE_URL)):
            with self.assertRaises(arm.ArmError) as error:
                client.login({'username': 'u', 'password': 'secret'})
        self.assertEqual(error.exception.status, 403)
        self.assertNotIn('secret', str(error.exception))

    def test_redirect_cannot_send_credentials_to_another_host(self):
        client = arm.ArmClient()
        self.addCleanup(client.close)
        for target in ['http://zkh-kontur.mosreg.ru/new9/', 'https://evil.test/login', '//evil.test/login']:
            with self.assertRaises(arm.ArmError):
                client.safe_url(target)
        response = MagicMock(status_code=307, headers={'Location': 'https://evil.test/login'})
        with patch.object(client.session, 'request', return_value=response) as request:
            with self.assertRaises(arm.ArmError):
                client.request('POST', arm.BASE_URL, {'password': 'secret'})
            self.assertEqual(request.call_count, 1)

    def test_certificate_failure_has_safe_message(self):
        client = arm.ArmClient()
        self.addCleanup(client.close)
        with patch.object(client.session, 'request', side_effect=requests.exceptions.SSLError('sensitive transport details')):
            with self.assertRaises(arm.ArmError) as error:
                client.request('GET', arm.BASE_URL)
            self.assertIn('защищённое соединение', str(error.exception))
            self.assertNotIn('sensitive', str(error.exception))

    def test_no_seed_report_is_shown_as_live_data(self):
        with tempfile.TemporaryDirectory() as root, patch.object(runner, 'DATA', Path(root)), \
                patch.object(runner, 'status', return_value={'running': False}):
            result = runner.water_daily()
        self.assertFalse(result['available'])
        self.assertEqual(result['days'], {})


class RouteTests(unittest.TestCase):
    def setUp(self):
        self.app = FastAPI()
        self.app.include_router(edds.router)
        self.client = TestClient(self.app)

    def test_arm_data_requires_module_access(self):
        response = self.client.get('/edds/arm/report?from_date=2026-09-17&to_date=2026-09-17')
        self.assertEqual(response.status_code, 403)

    def test_authenticated_report_and_errors_are_not_cached(self):
        self.app.dependency_overrides[edds.require_edds] = lambda: None
        route = '/edds/arm/report?from_date=2026-09-17&to_date=2026-09-17'
        with patch.object(arm, 'fetch_report', return_value=[['id_cds_claim'], ['7']]):
            response = self.client.get(route)
            self.assertEqual(response.json()['grid'][1], ['7'])
            self.assertEqual(response.headers['cache-control'], 'no-store')
        with patch.object(arm, 'fetch_report', side_effect=arm.ArmError('Связь недоступна', 504)):
            response = self.client.get(route)
            self.assertEqual(response.status_code, 504)
            self.assertEqual(response.json()['detail'], 'Связь недоступна')
            self.assertEqual(response.headers['cache-control'], 'no-store')


if __name__ == '__main__':
    unittest.main()
