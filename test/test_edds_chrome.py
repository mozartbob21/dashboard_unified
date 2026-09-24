import os
import tempfile
import time
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from playwright.sync_api import Error as BrowserError
from services.edds import arm, chrome, collector


class ChromeTests(unittest.TestCase):
    def setUp(self):
        environment = patch.dict(os.environ, {'EDDS_CHROME_EXECUTABLE': '', 'EDDS_CHROME_HEADLESS': '1'})
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
        self.assertEqual(options['channel'], 'chrome')
        self.assertFalse(options['ignore_https_errors'])
        self.assertTrue(options['chromium_sandbox'])
        self.assertEqual(options['ignore_default_args'], ['--disable-extensions'])
        self.assertNotIn('args', options)

    def test_gost_executable_uses_chromium_api_with_visible_window(self):
        self.client.close()
        executable = 'C:/Program Files/Chromium-Gost/chrome.exe'
        with patch.dict(os.environ, {'EDDS_CHROME_EXECUTABLE': executable, 'EDDS_CHROME_HEADLESS': '0 '}):
            self.client = chrome.ChromeArmClient()
            self.addCleanup(self.client.close)
        args, options = self.engine.chromium.launch_persistent_context.call_args
        self.assertEqual(args, (str(chrome.PROFILE),))
        self.assertEqual(options['executable_path'], executable)
        self.assertNotIn('channel', options)
        self.assertTrue(options['chromium_sandbox'])
        self.assertFalse(options['ignore_https_errors'])
        self.assertFalse(options['headless'])

    def test_report_uses_browser_tab_not_python_or_api_request_context(self):
        self.page.evaluate.return_value = {'status':200,'url':arm.BASE_URL,'text':'id_cds_claim;text_message\n7;Прорыв\n'}
        self.assertEqual(self.client.report(date(2026,9,17),date(2026,9,17))[1], ['7','Прорыв'])
        script, payload = self.page.evaluate.call_args.args
        self.assertIn("mode: 'same-origin'", script)
        self.assertEqual(payload['method'], 'POST')
        self.assertEqual(payload['data']['saveToCSV'], '1')
        self.assertEqual(payload['data']['date_ot'], '17.09.26')
        self.page.goto.assert_called_once()
        self.context.request.post.assert_not_called()

    def test_navigation_always_reopens_portal_from_persisted_tab(self):
        self.client.open_portal()
        self.assertTrue(self.client.portal_open)
        self.page.goto.assert_called_once_with(arm.BASE_URL+'?act=cds_report_svod&id=3608', wait_until='domcontentloaded', timeout=45000)

    def test_login_uses_browser_fetch_and_never_places_secret_in_url(self):
        form='<form method="post"><input name="user"><input name="pass" type="password"></form>'
        self.page.evaluate.side_effect = [
            {'status':200,'url':arm.BASE_URL,'text':form},
            {'status':200,'url':arm.BASE_URL,'text':'<form><input name="date_ot"><input name="date_do"></form>'}]
        self.client.login({'username':'test','password':'test-secret'})
        payload=self.page.evaluate.call_args.args[1]
        self.assertEqual(payload['method'],'POST')
        self.assertEqual(payload['data']['pass'],'test-secret')
        self.assertNotIn('test-secret',payload['url'])

    def test_coordinates_follow_same_origin_csv_link(self):
        self.page.evaluate.side_effect = [
            {'status':200,'url':arm.BASE_URL,'text':'<html><a href="export.csv">csv</a></html>'},
            {'status':200,'url':arm.BASE_URL+'export.csv','text':'id_cds_claim;lat_;lon_\n1;55;37'}]
        self.assertEqual(self.client.report(date.today(),date.today(),True)[1],['1','55','37'])
        self.assertEqual(self.page.evaluate.call_args.args[1]['method'],'GET')

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
        with patch.dict(os.environ,{'EDDS_ARM_TRANSPORT':'auto'}):
            with patch.object(arm.sys,'platform','win32'):
                self.assertEqual(arm.transport(),'chrome')
            with patch.object(arm.sys,'platform','linux'):
                self.assertEqual(arm.transport(),'requests')
        with patch.dict(os.environ,{'EDDS_ARM_TRANSPORT':'invalid'}):
            with self.assertRaises(arm.ArmError):arm.transport()

    def test_configured_browser_modes_all_create_browser_client(self):
        for mode in ('auto', 'chrome', ' Chrome ', 'Chromium-Gost', 'chromium-gost'):
            with self.subTest(mode=mode), patch.dict(os.environ, {
                'EDDS_ARM_TRANSPORT': mode,
                'EDDS_CHROME_EXECUTABLE': 'C:/Program Files/Chromium-Gost/chrome.exe',
            }), patch.object(arm.sys, 'platform', 'win32'), \
                    patch.object(chrome, 'ChromeArmClient') as browser, patch.object(arm, 'ArmClient') as direct:
                self.assertEqual(arm.transport(), 'chrome')
                self.assertIs(arm.create_client(), browser.return_value)
                direct.assert_not_called()

    def test_gost_alias_requires_explicit_executable(self):
        with patch.dict(os.environ, {'EDDS_ARM_TRANSPORT': 'Chromium-Gost', 'EDDS_CHROME_EXECUTABLE': ''}):
            with self.assertRaisesRegex(arm.ArmError, 'EDDS_CHROME_EXECUTABLE'):
                arm.create_client()

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
