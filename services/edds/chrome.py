"""ARM transport through the configured Chromium-GOST and its certificate stack.

No APIRequestContext: that would use a different network stack. Login uses the
live portal form; reports run as same-origin fetches inside the authenticated tab.
"""
import argparse
import hashlib
import json
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from playwright.sync_api import Error as BrowserError, TimeoutError as BrowserTimeout, sync_playwright

from services.edds.arm import ArmClient, ArmError, BASE_URL, MAX_BYTES

PROFILE = Path(__file__).resolve().parents[2] / '.private' / 'edds' / 'Chromium-Gost-profile'
PORTAL_REPORT = Path(__file__).with_name('portal.js').read_text(encoding='utf-8')
REPORT_URL = BASE_URL + '?act=cds_report_svod&id=3608'
LOGIN_BUTTON_NAME = re.compile(r'^\s*(?:Войти(?:\s+в\s+систему)?|Авторизоваться|Вход)\s*$', re.I)

# Return only a state code: never collect form values, page text or credentials.
LOGIN_STATE = r'''() => {
  const visible = e => !!e && !!e.getClientRects().length && getComputedStyle(e).visibility !== 'hidden';
  const has = selector => Array.from(document.querySelectorAll(selector)).some(visible);
  if (window.__neuronaAuthBlocked) return 'blocked';
  if (has('iframe[src*="captcha"], .g-recaptcha, .h-captcha, input[name*="captcha" i]')) return 'captcha';
  if (has('input[autocomplete="one-time-code"], input[name="otp"], input[name="totp"]')) return 'verification';
  if (has('input[type="password"]')) {
    const error = Array.from(document.querySelectorAll('[role="alert"], .alert-danger, .login-error, #login-error, .error-message'))
      .some(e => visible(e) && e.textContent.trim() &&
        (!window.__neuronaAuthPriorErrors || window.__neuronaAuthPriorErrors.get(e) !== e.textContent));
    return error ? 'login_error' : 'login';
  }
  const entry = Array.from(document.querySelectorAll('button, a, input[type="submit"], input[type="button"], [role="button"]'))
    .some(e => visible(e) && e.id !== 'esia-auth-button' &&
      /^(?:Войти(?:\s+в\s+систему)?|Авторизоваться|Вход)$/i.test(
        (e.getAttribute('aria-label') || (e.tagName === 'INPUT' ? e.value : e.textContent) || '').trim()));
  if (entry && !window.__neuronaAuthAttemptActive) return 'entry';
  if (document.querySelector('[name="date_ot"]') && document.querySelector('[name="date_do"]')) return 'report';
  // A hidden AJAX form may mean its POST is still in flight. Only a new,
  // fully loaded document can be treated as a landing page before report verification.
  return document.readyState === 'complete' && !window.__neuronaAuthAttemptActive ? 'other' : '';
}'''

# Chromium enforces this before a form POST or XHR can follow a foreign redirect.
# A route/url check after click alone is too late for a 307 carrying the password.
LOGIN_GUARD = r'''() => {
  window.__neuronaAuthAttemptActive = true;
  window.__neuronaAuthPriorErrors = new WeakMap(Array.from(document.querySelectorAll(
    '[role="alert"], .alert-danger, .login-error, #login-error, .error-message')).map(e => [e, e.textContent]));
  if (document.querySelector('meta[data-neurona-auth-guard]')) return;
  window.__neuronaAuthBlocked = false;
  document.addEventListener('securitypolicyviolation', event => {
    if (['form-action', 'connect-src'].includes(event.effectiveDirective)) window.__neuronaAuthBlocked = true;
  });
  const policy = document.createElement('meta');
  policy.httpEquiv = 'Content-Security-Policy';
  policy.content = "form-action 'self'; connect-src 'self'";
  policy.dataset.neuronaAuthGuard = 'true';
  document.head.appendChild(policy);
}'''


def browser_executable():
    executable = os.getenv('EDDS_CHROME_EXECUTABLE', '').strip()
    if executable and Path(executable).is_dir():
        raise ArmError('В EDDS_CHROME_EXECUTABLE указана папка. Укажите полный путь к файлу '
                       'браузера Chromium-GOST, например .../Application/chrome.exe '
                       '(файл .env рядом с app.py).', 503)
    if not executable or not Path(executable).is_file():
        raise ArmError('Укажите существующий EXE Chromium-GOST на компьютере-сервере '
                       'в EDDS_CHROME_EXECUTABLE (файл .env рядом с app.py).', 503)
    return executable

# same-origin mode rejects a cross-origin redirect BEFORE forwarding credentials.
# Bound the stream in Chrome so a huge report never crosses the automation pipe.
BROWSER_FETCH = r'''async ({url, method, data, timeout, limit}) => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    const response = await fetch(url, {
      method, mode: 'same-origin', credentials: 'same-origin', redirect: 'follow',
      cache: 'no-store', signal: controller.signal,
      ...(data ? {body: new URLSearchParams(data)} : {})
    });
    if (!response.ok) return {status: response.status};
    if (Number(response.headers.get('content-length')) > limit) return {error: 'size'};
    const reader = response.body.getReader();
    let size = 0; const chunks = [];
    while (true) {
      const {value, done} = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > limit) {await reader.cancel(); return {error: 'size'};}
      chunks.push(value);
    }
    const bytes = new Uint8Array(size); let offset = 0;
    for (const chunk of chunks) {bytes.set(chunk, offset); offset += chunk.byteLength;}
    let text;
    try {text = new TextDecoder('utf-8', {fatal: true}).decode(bytes);}
    catch (_) {text = new TextDecoder('windows-1251').decode(bytes);}
    return {status: response.status, text, url: response.url};
  } catch (error) {
    return {error: error.name === 'AbortError' ? 'timeout' : 'network'};
  } finally {clearTimeout(timer); controller.abort();}
}'''


def browser_error(error):
    """Classify without ever returning Playwright call logs, form data or cookies."""
    message = str(error).lower()
    if any(code in message for code in ('err_cert_', 'err_bad_ssl_client_auth_cert', 'err_ssl_client_auth')):
        return ArmError('Браузер ЕДДС на сервере не смог проверить сертификат или предъявить клиентский сертификат. '
                        'Откройте настройку ЕДДС на сервере под пользователем Windows, у которого работает портал; '
                        'проверьте доверие к УЦ и выбор клиентского сертификата.')
    if 'executable' in message and ('exist' in message or 'found' in message):
        return ArmError('На сервере не найден браузер ЕДДС. Проверьте путь к Chromium-GOST в EDDS_CHROME_EXECUTABLE.', 503)
    if any(word in message for word in ('singleton', 'profile in use', 'processsingleton')):
        return ArmError('Профиль браузера ЕДДС уже открыт. Закройте окно настройки и повторите загрузку.', 409)
    return ArmError('Браузер на сервере не смог открыть АРМ ЕДДС. Выполните настройку профиля ЕДДС '
                    'на офисном компьютере и запускайте «Нейрону» под той же учётной записью Windows.', 502)


class ChromeArmClient(ArmClient):
    def __init__(self, *, headless=None):
        self.deadline = time.monotonic() + 150
        self.playwright = self.context = self.page = None
        self.portal_open = False
        executable = browser_executable()
        try:
            PROFILE.mkdir(parents=True, exist_ok=True, mode=0o700)
        except OSError:
            raise ArmError('Нет доступа к рабочему профилю ЕДДС. Проверьте права на папку .private/edds на сервере.', 503) from None
        if headless is None:
            headless = os.getenv('EDDS_CHROME_HEADLESS', '0').strip().lower() not in {'0', 'false', 'no'}
        try:
            self.playwright = sync_playwright().start()
            options = {'headless': headless, 'ignore_https_errors': False,
                       'chromium_sandbox': True, 'service_workers': 'block',
                       'ignore_default_args': ['--disable-extensions'],
                       'timeout': 30000, 'accept_downloads': False}
            options['executable_path'] = executable
            self.context = self.playwright.chromium.launch_persistent_context(str(PROFILE), **options)
            self.context.set_default_timeout(45000)
            self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        except BrowserError as error:
            self.close()
            raise browser_error(error) from None

    def _login_timeout(self, deadline, maximum=15000):
        remaining = min(self.deadline, deadline) - time.monotonic()
        if remaining <= 0:
            raise ArmError('Истекло время входа в АРМ ЕДДС. Проверьте окно Chromium-GOST на сервере и повторите загрузку.', 504)
        return max(1, min(maximum, int(remaining * 1000)))

    def _login_state(self, states, deadline, maximum=15000):
        script = '(states) => { const state = (' + LOGIN_STATE + ')(); return states.includes(state) ? state : false; }'
        return self.page.wait_for_function(script, arg=states,
            timeout=self._login_timeout(deadline, maximum)).json_value()

    def _login_challenge(self, state):
        if state == 'blocked':
            raise ArmError('Форма входа АРМ попыталась отправить запрос на другой сайт. '
                           'Передача остановлена; нужна проверка интеграции.', 409)
        if state in {'captcha', 'verification'}:
            raise ArmError('АРМ ЕДДС требует капчу или код подтверждения. '
                           'Завершите проверку в Chromium-GOST на компьютере-сервере и повторите загрузку.', 409)

    def _check_login_target(self, form, submit):
        self.safe_url(self.page.url)
        # DOM properties resolve relative actions (including <base>) exactly as the browser does.
        target = form.evaluate('(form) => ({action: form.action, method: form.method, target: form.target})')
        override = submit.evaluate('''(button) => ({
          action: button.hasAttribute('formaction') ? button.formAction : null,
          method: button.hasAttribute('formmethod') ? button.formMethod : null,
          target: button.hasAttribute('formtarget') ? button.formTarget : null
        })''')
        self.safe_url(override['action'] or target['action'] or self.page.url)
        if (override['method'] or target['method'] or 'get').lower() != 'post':
            raise ArmError('Форма входа АРМ не использует POST. Отправка пароля в адресной строке запрещена; '
                           'нужна проверка формы портала.', 409)
        if (override['target'] if override['target'] is not None else target['target']) not in ('', '_self'):
            raise ArmError('Форма входа АРМ открывает другое окно. Нужна проверка интеграции.', 409)

    def _login_button(self, scope, *, entry=False):
        # Keep a semantic locator, not an index in a changing list of modal buttons.
        visible = scope.locator(':visible:not(#esia-auth-button)')
        button = scope.get_by_role('button', name=LOGIN_BUTTON_NAME).and_(visible)
        if entry:
            button = button.or_(scope.get_by_role('link', name=LOGIN_BUTTON_NAME).and_(visible))
        if button.count() != 1:
            label = 'Авторизоваться' if entry else 'Войти'
            raise ArmError(f'Не удалось однозначно найти кнопку «{label}» АРМ ЕДДС. '
                           'Нужна проверка страницы входа.')
        return button.first

    def _open_login_form(self, deadline):
        button = self._login_button(self.page, entry=True)
        self.safe_url(self.page.url)
        href = button.get_attribute('href', timeout=self._login_timeout(deadline))
        # Modal openers can use a fragment or a local JavaScript handler.
        # Recheck the actual page origin before entering any credentials.
        if href and not href.startswith(('#', 'javascript:')):
            self.safe_url(href)
        if button.get_attribute('target', timeout=self._login_timeout(deadline)) not in (None, '', '_self'):
            raise ArmError('Кнопка авторизации АРМ открывает другое окно. Нужна проверка интеграции.', 409)
        button.click(timeout=self._login_timeout(deadline), no_wait_after=True)

    def _type_login_field(self, field, value, deadline):
        # Some portal versions enable submit only on keyup/change, not on input.
        # Clear browser autofill, then type normally so all validators run.
        field.fill('', timeout=self._login_timeout(deadline))
        field.press_sequentially(value, delay=20, timeout=self._login_timeout(deadline))
        field.press('Tab', timeout=self._login_timeout(deadline))

    def login(self, account):
        deadline = min(self.deadline, time.monotonic() + 75)
        stage = 'open'
        try:
            # Cookies may have expired or been cleared for another admin account.
            # A previously rendered report is not proof that this session is still valid.
            self.open_portal(timeout=self._login_timeout(deadline, 30000))
            state = self._login_state(['report', 'entry', 'login', 'login_error', 'captcha', 'verification'], deadline)
            self.safe_url(self.page.url)
            self._login_challenge(state)
            if state == 'entry':
                stage = 'entry'
                self._open_login_form(deadline)
                state = self._login_state(['report', 'login', 'login_error', 'captcha', 'verification'], deadline)
                self.safe_url(self.page.url)
                self._login_challenge(state)
            if state == 'report':
                return
            stage = 'fields'
            password = self.page.locator('input[type="password"]:visible').first
            form = password.locator('xpath=ancestor::form[1]')
            if form.count() != 1:
                raise ArmError('Форма с полем пароля АРМ ЕДДС не найдена. Нужна проверка страницы входа.')
            username = form.locator('input[autocomplete="username"]:visible, input[name="login"]:visible, '
                                    'input[name="username"]:visible, input[name="user"]:visible, input[type="email"]:visible').first
            if not username.count():
                username = form.locator('input:is(:not([type]), [type="text"]):visible'
                    ':not([name*="captcha" i]):not([name*="otp" i]):not([name*="code" i])').first
            if not username.count():
                raise ArmError('Не найдено поле логина АРМ ЕДДС. Нужна проверка страницы входа.')
            submit = self._login_button(form)
            self._check_login_target(form, submit)
            self.page.evaluate(LOGIN_GUARD)
            self._type_login_field(username, account['username'], deadline)
            self.safe_url(self.page.url)
            self._type_login_field(password, account['password'], deadline)
            submit = self._login_button(form)
            self._check_login_target(form, submit)
            stage = 'submit'
            try:
                submit.click(timeout=self._login_timeout(deadline), no_wait_after=True)
            except BrowserTimeout:
                # No force/Enter fallback: a second submit could duplicate a pending login.
                if submit.count() and not submit.is_enabled(timeout=self._login_timeout(deadline, 1000)):
                    raise ArmError('Поля АРМ ЕДДС заполнены, но портал оставил кнопку «Войти» '
                                   'неактивной. Проверьте подсказки у полей в Chromium-GOST на сервере.', 504) from None
                raise
            stage = 'result'
            try:
                state = self._login_state(['report', 'other', 'login_error', 'captcha', 'verification', 'blocked'],
                                          deadline, 20000)
            except BrowserTimeout:
                # Some forms keep the login DOM after setting a session cookie.
                # Verify by reopening the report, without resubmitting credentials.
                state = self.page.evaluate(LOGIN_STATE)
            self.safe_url(self.page.url)
            self._login_challenge(state)
            if state == 'login_error':
                raise ArmError('АРМ ЕДДС не принял вход. Проверьте логин и пароль в записи «АРМ ЕДДС» '
                               'и сообщение в окне Chromium-GOST на сервере.', 403)
            self.open_portal(timeout=self._login_timeout(deadline, 20000))
            state = self._login_state(['report', 'entry', 'login', 'login_error', 'captcha', 'verification'], deadline)
            self.safe_url(self.page.url)
            self._login_challenge(state)
            if state != 'report':
                raise ArmError('После отправки формы АРМ ЕДДС снова запросил вход. Проверьте сохранённые '
                               'реквизиты «АРМ ЕДДС» и права аккаунта на сводный отчёт.', 403)
        except BrowserTimeout:
            messages = {
                'open': 'АРМ ЕДДС не показал форму входа или отчёт за отведённое время.',
                'entry': 'После нажатия «Авторизоваться» АРМ ЕДДС не показал поля входа.',
                'fields': 'Не удалось заполнить поля входа АРМ ЕДДС за отведённое время.',
                'submit': 'Не удалось нажать «Войти» АРМ ЕДДС: кнопка скрыта, перекрыта или страница ещё меняется.',
                'result': 'Кнопка «Войти» нажата, но АРМ ЕДДС не подтвердил вход за отведённое время.',
            }
            raise ArmError(messages[stage] + ' Проверьте окно Chromium-GOST на компьютере-сервере.', 504) from None
        except BrowserError as error:
            raise browser_error(error) from None

    def authenticate(self, account):
        # Do not reuse the previous administrator's portal session after an account change.
        marker = PROFILE / 'neurona-account.sha256'
        identity = hashlib.sha256(account['username'].encode('utf-8')).hexdigest()
        try:
            previous = marker.read_text(encoding='ascii', errors='replace') if marker.exists() else ''
            if previous != identity:
                self.context.clear_cookies()
                # Bind even a pending login to this account so a manually completed
                # CAPTCHA/second factor is not erased on the next attempt.
                marker.write_text(identity, encoding='ascii')
            self.login(account)
        except OSError:
            raise ArmError('Не удалось сохранить настройки рабочего профиля ЕДДС. '
                           'Проверьте права на .private/edds.', 503) from None
        except BrowserError as error:
            raise browser_error(error) from None

    def report(self, start, end, coordinates=False):
        if not self.portal_open:
            self.open_portal()
        self.safe_url(self.page.url)
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise ArmError('АРМ ЕДДС не успел сформировать отчёт. Сократите период.', 504)
        try:
            result = self.page.evaluate(PORTAL_REPORT, {
                'fromDate': start.isoformat(), 'toDate': end.isoformat(), 'coordinates': coordinates,
                'timeout': int(remaining * 1000), 'limit': MAX_BYTES,
            })
        except BrowserError as error:
            raise browser_error(error) from None
        if result.get('error'):
            raise ArmError(result['error'], result.get('status', 502))
        if not isinstance(result.get('grid'), list) or not result['grid']:
            raise ArmError('АРМ ЕДДС не вернул таблицу отчёта.')
        return result['grid']

    def close(self):
        # Called on browser restart or application shutdown, on the owning thread.
        try:
            if self.context:
                self.context.close()
        except BrowserError:
            pass
        finally:
            if self.playwright:
                try:
                    self.playwright.stop()
                except BrowserError:
                    pass
            self.context = self.playwright = self.page = None

    def open_portal(self, timeout=45000):
        try:
            self.page.goto(REPORT_URL, wait_until='domcontentloaded', timeout=timeout)
            self.safe_url(self.page.url)
            self.portal_open = True
        except BrowserTimeout:
            raise ArmError('Chromium-GOST не дождался страницы АРМ ЕДДС. Проверьте выбор сертификата '
                           'и открывшееся окно на сервере.', 504) from None
        except BrowserError as error:
            raise browser_error(error) from None

    def request(self, method, url, data=None):
        url = self.safe_url(url)
        if method not in {'GET', 'POST'}:
            raise ArmError('Неподдерживаемый запрос к АРМ ЕДДС.')
        if not self.portal_open:
            self.open_portal()
        # Reused profile tabs must never receive credentials while on another site.
        self.safe_url(self.page.url)
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise ArmError('АРМ ЕДДС не успел сформировать отчёт. Сократите период.', 504)
        try:
            result = self.page.evaluate(BROWSER_FETCH, {
                'url': url, 'method': method, 'data': data,
                'timeout': int(min(remaining, 60) * 1000), 'limit': MAX_BYTES,
            })
        except BrowserError as error:
            raise browser_error(error) from None
        if result.get('error') == 'size':
            raise ArmError('Отчёт слишком большой. Сократите период.', 413)
        if result.get('error') == 'timeout':
            raise ArmError('АРМ ЕДДС не ответил через браузер. Сократите период и повторите запрос.', 504)
        if result.get('error'):
            raise ArmError('Браузер открыл портал, но не смог получить отчёт. Проверьте доступ в рабочем '
                           'профиле ЕДДС на сервере, сертификаты и права на отчёт.')
        status = result.get('status', 502)
        if status in (401, 403):
            raise ArmError('АРМ ЕДДС отказал в доступе. Проверьте сохранённый логин, пароль и права на отчёты.', 403)
        if not 200 <= status < 300:
            raise ArmError(f'АРМ ЕДДС ответил кодом {status}. Повторите позднее.')
        return result['text'], self.safe_url(result['url'])


class BrowserService:
    """One long-lived GOST session, used only on its own Playwright worker thread."""
    def __init__(self, factory=None):
        self.factory = factory
        self._guard = threading.Lock()
        self._executor = None
        self._client = None
        self._account = None
        self._settings = None

    def report(self, account, start, end, coordinates=False):
        with self._guard:
            if self._executor is None:
                self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix='edds-gost')
            future = self._executor.submit(self._report, account, start, end, coordinates)
        return future.result()

    def _report(self, account, start, end, coordinates):
        settings = (os.getenv('EDDS_CHROME_EXECUTABLE', '').strip(), os.getenv('EDDS_CHROME_HEADLESS', '0').strip())
        if self._client is not None and (self._settings != settings or self._client.page.is_closed()):
            self._close_client()
        if self._client is None:
            self._client = (self.factory or ChromeArmClient)()
            self._settings = settings
        identity = hashlib.sha256(json.dumps(account, sort_keys=True).encode('utf-8')).digest()
        if self._account is not None and self._account != identity:
            try:
                self._client.context.clear_cookies()
            except BrowserError as error:
                self._close_client()
                raise browser_error(error) from None
        self._client.deadline = time.monotonic() + 150
        self._account = identity
        self._client.authenticate(account)
        try:
            return self._client.report(start, end, coordinates)
        except ArmError as error:
            if error.status != 403:
                raise
            # The portal session may expire between the login check and the export.
            self._client.authenticate(account)
            return self._client.report(start, end, coordinates)

    def _close_client(self):
        if self._client is not None:
            self._client.close()
        self._client = self._account = self._settings = None

    def close(self):
        with self._guard:
            executor, self._executor = self._executor, None
            if executor is not None:
                try:
                    executor.submit(self._close_client).result()
                finally:
                    executor.shutdown(wait=True)


browser_service = BrowserService()


def main():
    parser = argparse.ArgumentParser(description='Настройка рабочего профиля Chromium-GOST для ЕДДС на сервере')
    parser.add_argument('--setup', action='store_true', required=True)
    parser.parse_args()
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parents[2] / '.env')
    client = None
    try:
        client = ChromeArmClient(headless=False)
        try:
            client.open_portal()
        except ArmError as error:
            print(str(error))
        print('В открытом браузере проверьте вход в АРМ ЕДДС и доступ к отчёту по области.')
        print('Если нужен клиентский сертификат, выберите установленный сертификат вашей организации.')
        print('После проверки вернитесь сюда. Закрытый сертификат и пароль никуда копировать не нужно.')
        input('Нажмите Enter, чтобы закрыть профиль и завершить настройку: ')
    except (ArmError, KeyboardInterrupt, EOFError) as error:
        print(str(error) if isinstance(error, ArmError) else 'Настройка завершена.')
    finally:
        if client:
            client.close()


if __name__ == '__main__':
    main()
