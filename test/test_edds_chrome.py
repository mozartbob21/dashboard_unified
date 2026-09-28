import os
import tempfile
import time
import threading
from concurrent.futures import ThreadPoolExecutor
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from playwright.sync_api import Error as BrowserError
from services.edds import arm, chrome, collector


class ChromeTests(unittest.TestCase):
    def setUp(self):
        environment = patch.dict(os.environ, {'EDDS_CHROME_EXECUTABLE': __file__, 'EDDS_CHROME_HEADLESS': '1'})
        environment.start(); self.addCleanup(environment.stop)
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        patcher = patch.object(chrome, 'PROFILE', Path(self.temp.name) / 'profile')
        patcher.start(); self.addCleanup(patcher.stop)
        self.context = MagicMock()
        self.page = MagicMock()
        self.page.url = arm.BASE_URL
        self.context.pages = [self.page]
        patcher = patch.object(chrome, 'sync_playwright')
        self.engine = patcher.start().return_value.start.return_value
        self.addCleanup(patcher.stop)
        self.engine.chromium.launch_persistent_context.return_value = self.context
        self.client = chrome.ChromeArmClient()
        self.addCleanup(self.client.close)

    def test_installed_chrome_private_profile_and_strict_tls(self):
        args, options = self.engine.chromium.launch_persistent_context.call_args
        self.assertEqual(args, (str(chrome.PROFILE),))
        self.assertEqual(options['executable_path'], __file__)
        self.assertNotIn('channel', options)
        self.assertFalse(options['ignore_https_errors'])
        self.assertTrue(options['chromium_sandbox'])
        self.assertEqual(options['ignore_default_args'], ['--disable-extensions'])
        self.assertNotIn('args', options)

    def test_gost_executable_uses_chromium_api_with_visible_window(self):
        self.client.close()
        executable = 'C:/Program Files/Chromium-Gost/chrome.exe'
        with patch.dict(os.environ, {'EDDS_CHROME_EXECUTABLE': executable, 'EDDS_CHROME_HEADLESS': '0 '}), patch.object(chrome, 'browser_executable', return_value=executable):
            self.client = chrome.ChromeArmClient()
            self.addCleanup(self.client.close)
        args, options = self.engine.chromium.launch_persistent_context.call_args
        self.assertEqual(args, (str(chrome.PROFILE),))
        self.assertEqual(options['executable_path'], executable)
        self.assertNotIn('channel', options)
        self.assertTrue(options['chromium_sandbox'])
        self.assertFalse(options['ignore_https_errors'])
        self.assertFalse(options['headless'])

    def test_report_uses_archive_script_in_portal_tab(self):
        self.page.evaluate.return_value = {'grid': [['id_cds_claim', 'text_message'], ['7', 'Прорыв']]}
        self.assertEqual(self.client.report(date(2026,9,17),date(2026,9,17))[1], ['7','Прорыв'])
        script, payload = self.page.evaluate.call_args.args
        self.assertEqual(script, chrome.PORTAL_REPORT)
        self.assertEqual(payload['fromDate'], '2026-09-17')
        self.assertEqual(payload['toDate'], '2026-09-17')
        self.assertFalse(payload['coordinates'])
        self.page.goto.assert_called_once()
        self.context.request.post.assert_not_called()

    def test_navigation_always_reopens_portal_from_persisted_tab(self):
        self.client.open_portal()
        self.assertTrue(self.client.portal_open)
        self.page.goto.assert_called_once_with(arm.BASE_URL+'?act=cds_report_svod&id=3608', wait_until='domcontentloaded', timeout=45000)

    def test_active_session_is_reused_without_sending_password(self):
        self.page.wait_for_function.return_value.json_value.return_value = 'report'
        self.client.login({'username':'saved-admin','password':'saved-secret'})
        self.page.goto.assert_called_once()
        self.assertEqual(self.page.goto.call_args.args[0], chrome.REPORT_URL)
        self.page.locator.assert_not_called()
        self.page.evaluate.assert_not_called()

    def test_coordinates_use_same_browser_and_archive_script(self):
        self.page.evaluate.return_value = {'grid': [['Номер заявки', 'Широта', 'Долгота'], ['1', '55', '37']]}
        self.assertEqual(self.client.report(date.today(),date.today(),True)[1], ['1','55','37'])
        self.assertTrue(self.page.evaluate.call_args.args[1]['coordinates'])
        self.assertEqual(self.page.evaluate.call_args.args[0], chrome.PORTAL_REPORT)

    def test_foreign_url_is_rejected_before_password_is_sent(self):
        with self.assertRaises(arm.ArmError):
            self.client.request('POST','https://example.org/',{'password':'test-secret'})
        self.page.evaluate.assert_not_called()
        self.client.portal_open=True
        self.page.url='https://example.org/'
        with self.assertRaises(arm.ArmError):
            self.client.request('POST',arm.BASE_URL,{'password':'test-secret'})
        self.page.evaluate.assert_not_called()

    def test_transport_errors_are_bounded_and_safe(self):
        for result, status in [({'error':'timeout'},504),({'error':'size'},413),({'error':'network'},502),({'status':403},403)]:
            self.page.evaluate.return_value=result
            with self.assertRaises(arm.ArmError) as error:
                self.client.request('GET',arm.BASE_URL)
            self.assertEqual(error.exception.status,status)
        self.page.evaluate.side_effect=BrowserError('net::ERR_CERT_AUTHORITY_INVALID password=test-secret')
        with self.assertRaises(arm.ArmError) as error:
            self.client.request('GET',arm.BASE_URL)
        self.assertIn('сертификат',str(error.exception))
        self.assertNotIn('test-secret',str(error.exception))

    def test_no_request_after_deadline(self):
        self.client.portal_open=True
        self.client.deadline=time.monotonic()-1
        with self.assertRaises(arm.ArmError) as error:
            self.client.request('GET',arm.BASE_URL)
        self.assertEqual(error.exception.status,504)
        self.page.evaluate.assert_not_called()

    def test_close_stops_driver_even_if_context_close_fails(self):
        self.context.close.side_effect=BrowserError('browser already closed')
        self.client.close()
        self.engine.stop.assert_called_once()

    def test_automatic_mode_and_explicit_override(self):
        for mode in ('auto', 'chrome', 'chromium-gost'):
            with patch.dict(os.environ, {'EDDS_ARM_TRANSPORT': mode}):
                self.assertEqual(arm.transport(), 'chrome')
        for mode in ('requests', 'invalid'):
            with patch.dict(os.environ, {'EDDS_ARM_TRANSPORT': mode}):
                with self.assertRaises(arm.ArmError): arm.transport()

    def test_configured_browser_modes_all_create_browser_client(self):
        for mode in ('auto', 'chrome', ' Chrome ', 'Chromium-Gost', 'chromium-gost'):
            with self.subTest(mode=mode), patch.dict(os.environ, {
                'EDDS_ARM_TRANSPORT': mode,
                'EDDS_CHROME_EXECUTABLE': 'C:/Program Files/Chromium-Gost/chrome.exe',
            }), \
                    patch.object(chrome, 'ChromeArmClient') as browser, patch.object(arm, 'ArmClient') as direct:
                self.assertEqual(arm.transport(), 'chrome')
                self.assertIs(arm.create_client(), browser.return_value)
                direct.assert_not_called()

    def test_gost_requires_real_explicit_executable_without_fallback(self):
        for executable in ('', '/missing/Chromium-GOST/chrome.exe'):
            with patch.dict(os.environ, {'EDDS_ARM_TRANSPORT': 'Chromium-Gost', 'EDDS_CHROME_EXECUTABLE': executable}):
                with self.assertRaisesRegex(arm.ArmError, 'EDDS_CHROME_EXECUTABLE'):
                    arm.create_client()

    def test_gost_directory_is_rejected_before_browser_start(self):
        with patch.dict(os.environ, {'EDDS_CHROME_EXECUTABLE': self.temp.name}), \
                patch.object(chrome, 'sync_playwright') as playwright:
            with self.assertRaisesRegex(arm.ArmError, 'указана папка') as error:
                chrome.ChromeArmClient()
        self.assertEqual(error.exception.status, 503)
        self.assertIn('Application/chrome.exe', str(error.exception))
        self.assertNotIn(self.temp.name, str(error.exception))
        playwright.assert_not_called()

    def test_account_change_discards_old_cookie_but_same_account_keeps_session(self):
        account = {'username': 'admin-one', 'password': 'secret'}
        with patch.object(self.client, 'login') as login:
            self.client.authenticate(account)
            self.client.authenticate(account)
            self.assertEqual(self.context.clear_cookies.call_count, 1)
            self.client.authenticate({'username': 'admin-two', 'password': 'other'})
            self.assertEqual(self.context.clear_cookies.call_count, 2)
            self.assertEqual(login.call_count, 3)
        marker = (chrome.PROFILE / 'neurona-account.sha256').read_text()
        self.assertNotIn('secret', marker)
        self.assertNotIn('admin-two', marker)

    def test_closed_browser_during_authentication_returns_safe_error(self):
        self.context.clear_cookies.side_effect = BrowserError('Browser closed with token=private')
        with self.assertRaises(arm.ArmError) as error:
            self.client.authenticate({'username': 'saved-admin', 'password': 'secret'})
        self.assertNotIn('private', str(error.exception))

    def test_complaints_collector_uses_chromium_api(self):
        chromium = MagicMock()
        with patch.object(collector, 'load_water', return_value=None), \
                patch.object(collector, 'pull_windows', return_value=[]), \
                patch.object(collector, 'log'), patch.object(collector, 'sync_playwright') as manager, \
                patch.object(collector, 'ensure_session', side_effect=RuntimeError('stop before network')):
            manager.return_value.__enter__.return_value = SimpleNamespace(chromium=chromium)
            with self.assertRaisesRegex(RuntimeError, 'stop before network'):
                collector.main()
        chromium.launch.assert_called_once_with(headless=True)
        chromium.launch.return_value.close.assert_called_once()

    def test_chrome_error_does_not_silently_fallback_to_requests(self):
        with patch.dict(os.environ,{'EDDS_ARM_TRANSPORT':'chrome'}), \
             patch.object(chrome,'ChromeArmClient',side_effect=arm.ArmError('Chrome failed')), \
             patch.object(arm,'ArmClient') as requests:
            with self.assertRaises(arm.ArmError):arm.create_client()
            requests.assert_not_called()


class BrowserServiceTests(unittest.TestCase):
    def setUp(self):
        self.threads = []
        self.client = MagicMock()
        self.client.page.is_closed.return_value = False
        self.client.authenticate.side_effect = lambda account: self.threads.append(threading.get_ident())
        self.client.report.side_effect = lambda *args: self.threads.append(threading.get_ident()) or [['id_cds_claim'], ['7']]
        self.client.close.side_effect = lambda: self.threads.append(threading.get_ident())
        self.factory = MagicMock(side_effect=lambda: self.threads.append(threading.get_ident()) or self.client)
        self.service = chrome.BrowserService(factory=self.factory)
        self.addCleanup(self.service.close)
        self.account = {'username': 'saved-admin', 'password': 'saved-secret'}

    def test_calls_from_multiple_threads_share_one_browser_and_one_playwright_thread(self):
        with ThreadPoolExecutor(max_workers=2) as callers:
            tasks = [callers.submit(self.service.report, self.account, date.today(), date.today(), flag) for flag in (False, True)]
            for task in tasks: self.assertEqual(task.result()[1], ['7'])
        self.factory.assert_called_once()
        self.client.close.assert_not_called()
        self.service.close()
        self.client.close.assert_called_once()
        self.assertEqual(len(set(self.threads)), 1)
        self.assertNotIn(threading.get_ident(), self.threads)
        self.assertEqual(self.client.authenticate.call_count, 2)
        self.client.authenticate.assert_called_with(self.account)

    def test_authentication_failure_keeps_window_for_retry(self):
        self.client.authenticate.side_effect = [arm.ArmError('Denied', 403), None]
        with self.assertRaises(arm.ArmError): self.service.report(self.account, date.today(), date.today())
        self.assertEqual(self.service.report(self.account, date.today(), date.today())[1], ['7'])
        self.factory.assert_called_once()
        self.client.close.assert_not_called()

    def test_closed_browser_is_reopened_on_next_request(self):
        self.service.report(self.account, date.today(), date.today())
        self.client.page.is_closed.return_value = True
        self.service.report(self.account, date.today(), date.today())
        self.assertEqual(self.factory.call_count, 2)
        self.client.close.assert_called_once()

    def test_changed_admin_credentials_force_new_login(self):
        self.service.report(self.account, date.today(), date.today())
        changed = {**self.account, 'password': 'new-secret'}
        self.service.report(changed, date.today(), date.today())
        self.client.context.clear_cookies.assert_called_once()
        self.client.authenticate.assert_called_with(changed)

    def test_session_expiring_during_report_retries_login_once(self):
        self.client.report.side_effect = [arm.ArmError('Expired', 403), [['id_cds_claim'], ['7']]]
        self.assertEqual(self.service.report(self.account, date.today(), date.today())[1], ['7'])
        self.assertEqual(self.client.authenticate.call_count, 2)
        self.client.authenticate.assert_called_with(self.account)
        self.factory.assert_called_once()

    def test_forbidden_report_cannot_loop_login_indefinitely(self):
        self.client.report.side_effect = arm.ArmError('Forbidden', 403)
        with self.assertRaises(arm.ArmError): self.service.report(self.account, date.today(), date.today())
        self.assertEqual(self.client.report.call_count, 2)
