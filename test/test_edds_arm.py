import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

import requests
from fastapi import FastAPI
from fastapi.testclient import TestClient

from routers import edds
from services.edds import arm, runner


class ArmTests(unittest.TestCase):
    def setUp(self):
        environment = patch.dict(arm.os.environ, {"EDDS_ARM_TRANSPORT": "chrome"})
        environment.start()
        self.addCleanup(environment.stop)

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

    def test_arm_credentials_are_passed_to_shared_browser_service(self):
        from services.edds import chrome
        with patch.object(arm, 'credentials', return_value={'username': 'user', 'password': 'secret'}) as credentials, \
                patch.object(chrome, 'browser_service') as service:
            service.report.return_value = [['id_cds_claim'], ['1']]
            self.assertEqual(len(arm.fetch_report(date(2026, 9, 17), date(2026, 9, 17))), 2)
            credentials.assert_called_once_with('edds_arm')
            self.assertEqual(service.report.call_args.args, ({'username': 'user', 'password': 'secret'}, date(2026,9,17), date(2026,9,17), False))
            self.assertGreater(service.report.call_args.kwargs['deadline'], arm.time.monotonic())
            service.close.assert_not_called()

    def test_login_failure_releases_lock_for_retry(self):
        from services.edds import chrome
        with patch.object(arm, 'credentials', return_value={'username': 'u', 'password': 'secret'}), \
                patch.object(chrome, 'browser_service') as service:
            service.report.side_effect = arm.ArmError('Failed', 403)
            with self.assertRaises(arm.ArmError):
                arm.fetch_report(date.today(), date.today())
            self.assertFalse(arm.LOCK.locked())

    def test_large_period_is_loaded_in_contiguous_windows(self):
        from services.edds import chrome
        start, end = date(2026, 1, 1), date(2026, 3, 5)
        def report(_account, lo, hi, _coordinates, **_options):
            return [['id_cds_claim'], [lo.isoformat(), hi.isoformat()]]
        with patch.object(arm, 'credentials', return_value={'username': 'u', 'password': 'secret'}), \
                patch.object(chrome.browser_service, 'report', side_effect=report) as call:
            rows = arm.fetch_report(start, end)
        ranges = [(date.fromisoformat(row[0]), date.fromisoformat(row[1])) for row in rows[1:]]
        self.assertEqual(ranges[0][0], start)
        self.assertEqual(ranges[-1][1], end)
        self.assertTrue(all((hi-lo).days < 7 for lo,hi in ranges))
        for (_,previous_end),(following_start,_) in zip(ranges,ranges[1:]):
            self.assertEqual(previous_end+timedelta(days=1), following_start)
        self.assertEqual(sum((hi-lo).days+1 for lo,hi in ranges), (end-start).days+1)
        self.assertEqual(call.call_count, len(ranges))
        self.assertFalse(arm.LOCK.locked())

    def test_oversize_window_is_bisected_without_losing_rows(self):
        from services.edds import chrome
        def report(_account, lo, hi, _coordinates, **_options):
            if (hi - lo).days > 1:
                raise arm.ArmError('Отчёт слишком большой', 413)
            return [['id_cds_claim'], [lo.isoformat()], [hi.isoformat()]]
        with patch.object(arm, 'credentials', return_value={'username': 'u', 'password': 'secret'}), \
                patch.object(chrome.browser_service, 'report', side_effect=report):
            rows = arm.fetch_report(date(2026, 9, 1), date(2026, 9, 4))
        self.assertEqual(rows, [['id_cds_claim'], ['2026-09-01'], ['2026-09-02'],
                                ['2026-09-03'], ['2026-09-04']])

    def test_incompatible_report_headers_stop_merge(self):
        from services.edds import chrome
        def report(_account, lo, hi, _coordinates, **_options):
            return [[('id_cds_claim' if lo == date(2026,1,1) else 'other')], ['1']]
        with patch.object(arm, 'credentials', return_value={'username': 'u', 'password': 'secret'}), \
                patch.object(chrome.browser_service, 'report', side_effect=report):
            with self.assertRaisesRegex(arm.ArmError, 'Колонки отчёта'):
                arm.fetch_report(date(2026, 1, 1), date(2026, 2, 2))
        self.assertFalse(arm.LOCK.locked())

    def test_whole_period_deadline_stops_later_windows(self):
        from services.edds import chrome
        clock = {'now':0}
        def report(*args, **kwargs):
            self.assertLessEqual(kwargs['deadline'], 120)
            clock['now'] = 121
            return [['id_cds_claim'], ['1']]
        with patch.object(arm.time, 'monotonic', side_effect=lambda: clock['now']), \
             patch.object(arm, 'credentials', return_value={'username':'u','password':'p'}), \
             patch.object(chrome.browser_service, 'report', side_effect=report) as call:
            with self.assertRaises(arm.ArmError) as caught:
                arm.fetch_report(date(2026,1,1), date(2026,10,1))
        self.assertEqual(caught.exception.status, 504)
        self.assertEqual(call.call_count, 1)
        self.assertFalse(arm.LOCK.locked())

    def test_export_timeout_splits_but_login_timeout_does_not(self):
        from services.edds import chrome
        def report(_account, lo, hi, _coordinates, **options):
            if lo < hi:
                error = arm.ArmError('Export timed out',504)
                error.retry_smaller = True
                raise error
            return [['id_cds_claim'], [lo.isoformat()]]
        with patch.object(arm, 'credentials', return_value={'username':'u','password':'p'}), \
             patch.object(chrome.browser_service, 'report', side_effect=report):
            rows = arm.fetch_report(date(2026,1,1),date(2026,1,3))
        self.assertEqual(rows[1:],[['2026-01-01'],['2026-01-02'],['2026-01-03']])
        with patch.object(arm, 'credentials', return_value={'username':'u','password':'p'}), \
             patch.object(chrome.browser_service, 'report', side_effect=arm.ArmError('Login timed out',504)) as call:
            with self.assertRaises(arm.ArmError):
                arm.fetch_report(date(2026,1,1),date(2026,10,1))
            call.assert_called_once()

    def test_busy_lock_has_specific_retry_code(self):
        with patch.object(arm,'credentials',return_value={'username':'u','password':'p'}):
            arm.LOCK.acquire()
            try:
                with self.assertRaises(arm.ArmError) as caught:
                    arm.fetch_report(date(2026,1,1),date(2026,1,1))
            finally:
                arm.LOCK.release()
        self.assertEqual(caught.exception.status,409)
        self.assertEqual(caught.exception.code,'busy')

    def test_aggregate_size_is_bounded(self):
        from services.edds import chrome
        with patch.object(arm,'MAX_REPORT_BYTES',10), \
             patch.object(arm,'credentials',return_value={'username':'u','password':'p'}), \
             patch.object(chrome.browser_service,'report',return_value=[['id_cds_claim'],['1']]):
            with self.assertRaises(arm.ArmError) as caught:
                arm.fetch_report(date(2026,1,1),date(2026,1,1))
        self.assertEqual(caught.exception.status,413)
        self.assertFalse(arm.LOCK.locked())

    def test_collector_failure_is_classified_without_exposing_browser_trace(self):
        message = runner.failure_message('Executable doesn\'t exist at secret-user-path password=secret')
        self.assertIn('EDDS_CHROME_EXECUTABLE', message)
        self.assertNotIn('secret', message)
        self.assertIn('HTTP 503', runner.failure_message('Сервер вернул ошибку 503 на странице 2.'))

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

    def test_application_shutdown_closes_shared_browser(self):
        from services.edds import chrome
        with patch.object(chrome.browser_service, 'close') as close:
            with TestClient(self.app):
                close.assert_not_called()
            close.assert_called_once()

    def test_status_requires_module_access_before_reading_configuration(self):
        with patch.object(arm, 'transport') as transport:
            response = self.client.get('/edds/status')
        self.assertEqual(response.status_code, 403)
        transport.assert_not_called()

    def test_status_reports_bad_transport_as_handled_configuration_error(self):
        self.app.dependency_overrides[edds.require_edds] = lambda: None
        with patch.dict(arm.os.environ, {'EDDS_ARM_TRANSPORT': 'invalid-private-value'}):
            response = self.client.get('/edds/status')
        self.assertEqual(response.status_code, 503)
        self.assertIn('EDDS_ARM_TRANSPORT', response.json()['detail'])
        self.assertNotIn('invalid-private-value', response.text)
        self.assertEqual(response.headers['cache-control'], 'no-store')

    def test_status_accepts_server_gost_configuration(self):
        self.app.dependency_overrides[edds.require_edds] = lambda: None
        with patch.dict(arm.os.environ, {
            'EDDS_ARM_TRANSPORT': 'chrome',
            'EDDS_CHROME_EXECUTABLE': 'C:/Program Files/Chromium-Gost/chrome.exe',
            'EDDS_CHROME_HEADLESS': '0',
        }), patch.object(runner, 'status', return_value={'running': False}), \
                patch.object(edds, 'credentials', return_value={'username': 'test', 'password': 'secret'}):
            response = self.client.get('/edds/status')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['arm_transport'], 'chrome')
        self.assertTrue(response.json()['arm_configured'])
        self.assertNotIn('secret', response.text)

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


    def test_only_worker_busy_has_a_retryable_error_code(self):
        self.app.dependency_overrides[edds.require_edds] = lambda: None
        route='/edds/arm/report?from_date=2026-09-17&to_date=2026-09-17'
        for error,code in [(arm.ArmError('Worker still running',409,code='busy'),'busy'),
                           (arm.ArmError('Complete CAPTCHA on server',409),None)]:
            with patch.object(arm,'fetch_report',side_effect=error):
                response=self.client.get(route)
                self.assertEqual(response.status_code,409)
                self.assertEqual(response.json().get('code'),code)
                self.assertEqual(response.json()['detail'],str(error))


if __name__ == '__main__':
    unittest.main()
