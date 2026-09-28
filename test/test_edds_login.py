"""DOM-login regressions without a browser, network or stored credentials."""
import time
import tempfile
from pathlib import Path
from datetime import date
import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from playwright.sync_api import Error as BrowserError, TimeoutError as BrowserTimeout
from services.edds import arm, chrome


class FakeLocator:
    def __init__(self, harness, name, *, tag='INPUT', attrs=None, text='', present=True):
        self.harness = harness
        self.name = name
        self.tag = tag
        self.attrs = attrs or {}
        self.text = text
        self.present = present
        self.error = None
        self.on_click = None
        self.target = {'action': None, 'method': None, 'target': None}

    @property
    def first(self):
        return self

    def count(self):
        return int(self.present)

    def locator(self, selector):
        if self.name == 'password' and selector == 'xpath=ancestor::form[1]':
            return self.harness.form
        if self.name == 'form':
            if selector.startswith('button'):
                return SimpleNamespace(all=lambda: self.harness.buttons)
            self.harness.username_selectors.append(selector)
            return self.harness.username
        raise AssertionError('Unexpected locator: ' + selector)

    def evaluate(self, script):
        if script == '(e) => e.tagName':
            return self.tag
        return dict(self.target)

    def inner_text(self):
        return self.text

    def get_attribute(self, key):
        return self.attrs.get(key)

    def fill(self, value, *, timeout):
        self.harness.events.append(('fill', self.name, value))
        self.harness.timeouts.append(timeout)
        if self.error:
            raise self.error

    def press(self, key, *, timeout):
        self.harness.events.append(('press', self.name, key))
        self.harness.timeouts.append(timeout)

    def click(self, *, timeout, no_wait_after):
        self.harness.events.append(('click', self.name))
        self.harness.timeouts.append(timeout)
        if self.error:
            raise self.error
        if self.on_click:
            self.on_click()


class LoginTests(unittest.TestCase):
    def setUp(self):
        self.account = {'username': 'synthetic-admin', 'password': 'synthetic-private-password'}
        self.events = []
        self.timeouts = []
        self.username_selectors = []
        self.password = FakeLocator(self, 'password', attrs={'name': 'pass', 'id': 'password'})
        self.username = FakeLocator(self, 'username', attrs={'name': 'login'})
        self.form = FakeLocator(self, 'form', tag='FORM')
        self.form.target = {'action': arm.BASE_URL + '?act=login', 'method': 'post', 'target': ''}
        # ESIA deliberately comes first, as in the reported portal layout.
        self.esia = FakeLocator(self, 'esia', tag='BUTTON', attrs={'id': 'esia-auth-button'}, text='Войти')
        self.submit = FakeLocator(self, 'submit', tag='BUTTON', attrs={'type': 'submit'}, text='Войти')
        self.buttons = [self.esia, self.submit]
        self.page = MagicMock()
        self.page.url = chrome.REPORT_URL
        self.page.locator.side_effect = self.locator
        self.page.evaluate.side_effect = self.evaluate
        self.page.goto.side_effect = self.goto
        self.client = object.__new__(chrome.ChromeArmClient)
        self.client.page = self.page
        self.client.context = MagicMock()
        self.client.playwright = None
        self.client.portal_open = True
        self.client.deadline = time.monotonic() + 150
        self.states = MagicMock(side_effect=['login', 'report', 'report'])
        patcher = patch.object(self.client, '_login_state', self.states)
        patcher.start()
        self.addCleanup(patcher.stop)

    def locator(self, selector):
        self.assertEqual(selector, 'input[type="password"]:visible')
        return self.password

    def evaluate(self, script):
        if script == chrome.LOGIN_GUARD:
            self.events.append(('guard',))
            return None
        if script == chrome.LOGIN_STATE:
            return 'login'
        raise AssertionError('Unexpected evaluation')

    def goto(self, url, **options):
        self.assertEqual(url, chrome.REPORT_URL)
        self.page.url = url
        self.events.append(('navigate', url))
        self.timeouts.append(options['timeout'])

    def assert_no_fill(self):
        self.assertFalse(any(event[0] == 'fill' for event in self.events))

    def assert_safe_error(self, error, status):
        self.assertEqual(error.exception.status, status)
        self.assertNotIn(self.account['password'], str(error.exception))
        self.assertNotIn('Call log', str(error.exception))

    def test_live_pass_field_input_blur_submit_and_fresh_success_proof(self):
        self.client.login(self.account)
        self.assertEqual(self.page.goto.call_count, 2)
        self.assertEqual(self.events, [
            ('navigate', chrome.REPORT_URL), ('guard',),
            ('fill', 'username', self.account['username']),
            ('fill', 'password', self.account['password']),
            ('press', 'password', 'Tab'), ('click', 'submit'),
            ('navigate', chrome.REPORT_URL),
        ])
        self.assertTrue(all(0 < value <= 30000 for value in self.timeouts))
        self.client.context.request.post.assert_not_called()
        self.assertIn("form-action 'self'", chrome.LOGIN_GUARD)
        self.assertIn("connect-src 'self'", chrome.LOGIN_GUARD)

    def test_existing_session_is_checked_by_fresh_navigation_without_fill(self):
        self.states.side_effect = ['report']
        self.client.login(self.account)
        self.page.goto.assert_called_once()
        self.page.locator.assert_not_called()
        self.assert_no_fill()

    def test_foreign_form_action_is_rejected_before_password_fill(self):
        self.form.target['action'] = 'https://outside.invalid/login'
        with self.assertRaises(arm.ArmError) as error:
            self.client.login(self.account)
        self.assert_safe_error(error, 502)
        self.assert_no_fill()

    def test_effective_submit_overrides_are_validated_before_fill(self):
        for override, value in [('action', 'https://outside.invalid/login'), ('method', 'get'), ('target', '_blank')]:
            with self.subTest(override=override):
                self.submit.target = {'action': None, 'method': None, 'target': None}
                self.submit.target[override] = value
                self.states.side_effect = ['login']
                self.events.clear()
                with self.assertRaises(arm.ArmError):
                    self.client.login(self.account)
                self.assert_no_fill()

    def test_get_form_never_places_password_in_url(self):
        self.form.target['method'] = 'get'
        with self.assertRaises(arm.ArmError) as error:
            self.client.login(self.account)
        self.assert_safe_error(error, 409)
        self.assert_no_fill()

    def test_disabled_submit_timeout_is_bounded_sanitized_and_not_retried(self):
        self.submit.error = BrowserTimeout('Call log: click disabled, password=' + self.account['password'])
        with self.assertRaises(arm.ArmError) as error:
            self.client.login(self.account)
        self.assert_safe_error(error, 504)
        self.assertIn('Кнопка', str(error.exception))
        self.assertEqual(sum(event[0] == 'click' for event in self.events), 1)

    def test_fill_error_never_returns_playwright_log_or_password(self):
        self.password.error = BrowserError('Call log: fill("' + self.account['password'] + '")')
        with self.assertRaises(arm.ArmError) as error:
            self.client.login(self.account)
        self.assert_safe_error(error, 502)
        self.assertFalse(any(event[0] == 'click' for event in self.events))

    def test_initial_captcha_or_verification_stops_before_credentials(self):
        for state in ['captcha', 'verification']:
            with self.subTest(state=state):
                self.states.side_effect = [state]
                self.events.clear()
                with self.assertRaises(arm.ArmError) as error:
                    self.client.login(self.account)
                self.assert_safe_error(error, 409)
                self.assert_no_fill()

    def test_challenge_or_blocked_redirect_after_submit_is_not_retried(self):
        for state in ['captcha', 'verification', 'blocked']:
            with self.subTest(state=state):
                self.states.side_effect = ['login', state]
                self.events.clear()
                with self.assertRaises(arm.ArmError) as error:
                    self.client.login(self.account)
                self.assert_safe_error(error, 409)
                self.assertEqual(sum(event[0] == 'click' for event in self.events), 1)

    def test_report_probe_proves_cookie_login_when_login_dom_remains(self):
        self.states.side_effect = ['login', BrowserTimeout('No navigation'), 'report']
        self.client.login(self.account)
        self.assertEqual(self.page.goto.call_count, 2)
        self.assertEqual(sum(event[0] == 'click' for event in self.events), 1)

    def test_fresh_report_requesting_login_means_authentication_failed(self):
        self.states.side_effect = ['login', 'other', 'login']
        with self.assertRaises(arm.ArmError) as error:
            self.client.login(self.account)
        self.assert_safe_error(error, 403)
        self.assertEqual(self.page.goto.call_count, 2)

    def test_wrong_password_has_no_retry_and_sanitized_error(self):
        self.states.side_effect = ['login', 'login_error']
        with self.assertRaises(arm.ArmError) as error:
            self.client.login(self.account)
        self.assert_safe_error(error, 403)
        self.assertEqual(sum(event[0] == 'click' for event in self.events), 1)

    def test_esia_only_form_is_never_clicked(self):
        self.buttons = [self.esia]
        with self.assertRaises(arm.ArmError):
            self.client.login(self.account)
        self.assert_no_fill()
        self.assertFalse(any(event[0] == 'click' for event in self.events))

    def test_esia_label_is_excluded_even_without_special_id(self):
        self.esia.attrs = {'type': 'submit'}
        self.esia.text = 'Войти через Госуслуги'
        self.client.login(self.account)
        self.assertEqual(sum(event == ('click', 'submit') for event in self.events), 1)
        self.assertNotIn(('click', 'esia'), self.events)

    def test_unexpected_origin_before_fill_never_receives_credentials(self):
        def redirected(url, **options):
            self.page.url = 'https://outside.invalid/login'
        self.page.goto.side_effect = redirected
        self.states.side_effect = ['login']
        with self.assertRaises(arm.ArmError):
            self.client.login(self.account)
        self.assert_no_fill()

    def test_old_alert_does_not_prevent_fresh_report_proof(self):
        # The guard masks existing alert nodes; unchanged AJAX login DOM can time out.
        self.states.side_effect = ['login_error', BrowserTimeout('Old alert remained'), 'report']
        self.client.login(self.account)
        self.assertEqual(self.page.goto.call_count, 2)
        self.assertEqual(sum(event[0] == 'click' for event in self.events), 1)

    def test_navigation_timeout_returns_safe_504(self):
        self.page.goto.side_effect = BrowserTimeout('Call log secret=' + self.account['password'])
        with self.assertRaises(arm.ArmError) as error:
            self.client.login(self.account)
        self.assert_safe_error(error, 504)
        self.assert_no_fill()

    def test_login_state_wait_is_bounded_and_carries_only_state_codes(self):
        with patch.object(chrome.time, 'monotonic', return_value=100):
            self.client.deadline = 140
            self.page.wait_for_function.return_value.json_value.return_value = 'report'
            result = chrome.ChromeArmClient._login_state(self.client, ['login', 'report'], 110, 15000)
        self.assertEqual(result, 'report')
        args, kwargs = self.page.wait_for_function.call_args
        self.assertEqual(kwargs['timeout'], 10000)
        self.assertEqual(kwargs['arg'], ['login', 'report'])
        self.assertNotIn(self.account['password'], args[0])

    def test_manual_verification_is_not_erased_on_authentication_retry(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(chrome, 'PROFILE', Path(folder)), \
                patch.object(self.client, 'login', side_effect=[arm.ArmError('Verification', 409), None]) as login:
            with self.assertRaises(arm.ArmError):
                self.client.authenticate(self.account)
            self.client.authenticate(self.account)
            self.client.context.clear_cookies.assert_called_once()
            self.assertEqual(login.call_count, 2)
            marker = (Path(folder) / 'neurona-account.sha256').read_text()
            self.assertNotIn(self.account['username'], marker)
            self.assertNotIn(self.account['password'], marker)

    def test_changed_credentials_challenge_keeps_manual_session_on_next_report(self):
        browser = MagicMock()
        browser.page.is_closed.return_value = False
        browser.report.return_value = [['id_cds_claim'], ['7']]
        browser.authenticate.side_effect = [None, arm.ArmError('Verification', 409), None]
        service = chrome.BrowserService(factory=lambda: browser)
        self.addCleanup(service.close)
        service.report(self.account, date.today(), date.today())
        changed = {**self.account, 'password': 'changed-synthetic-secret'}
        with self.assertRaises(arm.ArmError):
            service.report(changed, date.today(), date.today())
        self.assertEqual(service.report(changed, date.today(), date.today())[1], ['7'])
        browser.context.clear_cookies.assert_called_once()
        self.assertEqual(browser.authenticate.call_count, 3)
        browser.close.assert_not_called()

    def test_expired_deadline_stops_before_navigation(self):
        self.client.deadline = time.monotonic() - 1
        with self.assertRaises(arm.ArmError) as error:
            self.client.login(self.account)
        self.assert_safe_error(error, 504)
        self.page.goto.assert_not_called()


if __name__ == '__main__':
    unittest.main()
