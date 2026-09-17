"""ARM transport through installed Chrome, including its TLS/certificate stack.

No APIRequestContext: that would use a different network stack. All report and
login requests run as same-origin fetches inside the portal tab.
"""
import argparse
import os
import time
from pathlib import Path

from playwright.sync_api import Error as BrowserError, sync_playwright

from services.edds.arm import ArmClient, ArmError, BASE_URL, MAX_BYTES

PROFILE = Path(__file__).resolve().parents[2] / '.private' / 'edds' / 'chrome-profile'

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
        return ArmError('Chrome на сервере не смог проверить сертификат или предъявить клиентский сертификат. '
                        'Откройте настройку ЕДДС на сервере под пользователем Windows, у которого работает портал; '
                        'проверьте доверие к УЦ и выбор клиентского сертификата.')
    if 'executable' in message and ('exist' in message or 'found' in message):
        return ArmError('На сервере не найден Google Chrome. Установите Chrome или укажите EDDS_CHROME_EXECUTABLE.', 503)
    if any(word in message for word in ('singleton', 'profile in use', 'processsingleton')):
        return ArmError('Профиль Chrome ЕДДС уже открыт. Закройте окно настройки и повторите загрузку.', 409)
    return ArmError('Chrome на сервере не смог открыть АРМ ЕДДС. Выполните настройку профиля ЕДДС '
                    'на офисном компьютере и запускайте «Нейрону» под той же учётной записью Windows.', 502)


class ChromeArmClient(ArmClient):
    def __init__(self, *, headless=None):
        self.deadline = time.monotonic() + 150
        self.playwright = self.context = self.page = None
        self.portal_open = False
        try:
            PROFILE.mkdir(parents=True, exist_ok=True, mode=0o700)
        except OSError:
            raise ArmError('Нет доступа к рабочему профилю ЕДДС. Проверьте права на папку .private/edds на сервере.', 503) from None
        if headless is None:
            headless = os.getenv('EDDS_CHROME_HEADLESS', '1').lower() not in {'0', 'false', 'no'}
        try:
            self.playwright = sync_playwright().start()
            options = {'headless': headless, 'ignore_https_errors': False,
                       'chromium_sandbox': True, 'service_workers': 'block',
                       'ignore_default_args': ['--disable-extensions'],
                       'timeout': 30000, 'accept_downloads': False}
            executable = os.getenv('EDDS_CHROME_EXECUTABLE', '').strip()
            if executable:
                options['executable_path'] = executable
            else:
                options['channel'] = 'chrome'
            self.context = self.playwright.chromium.launch_persistent_context(str(PROFILE), **options)
            self.context.set_default_timeout(45000)
            self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        except BrowserError as error:
            self.close()
            raise browser_error(error) from None

    def close(self):
        # Always release the Chrome profile, even when a report or login failed.
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

    def open_portal(self):
        try:
            self.page.goto(BASE_URL + '?act=cds_report_svod&id=3608', wait_until='domcontentloaded', timeout=45000)
            self.safe_url(self.page.url)
            self.portal_open = True
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
            raise ArmError('АРМ ЕДДС не ответил через Chrome. Сократите период и повторите запрос.', 504)
        if result.get('error'):
            raise ArmError('Chrome открыл портал, но не смог получить отчёт. Проверьте доступ в рабочем '
                           'профиле ЕДДС на сервере, сертификаты и права на отчёт.')
        status = result.get('status', 502)
        if status in (401, 403):
            raise ArmError('АРМ ЕДДС отказал в доступе. Проверьте сохранённый логин, пароль и права на отчёты.', 403)
        if not 200 <= status < 300:
            raise ArmError(f'АРМ ЕДДС ответил кодом {status}. Повторите позднее.')
        return result['text'], self.safe_url(result['url'])


def main():
    parser = argparse.ArgumentParser(description='Настройка рабочего профиля Google Chrome для ЕДДС на сервере')
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
        print('В открытом Chrome проверьте вход в АРМ ЕДДС и доступ к отчёту по области.')
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
