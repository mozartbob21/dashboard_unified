"""Dobrodel's current admin login and report API, with a private cookie session."""
import os
import time
from urllib.parse import urljoin, urlsplit

import requests
from bs4 import BeautifulSoup

BASE = 'https://admin.vmeste.mosreg.ru'
REPORT_PATH = '/OperativeReportGenerating'
MAX_BYTES = 20 * 1024 * 1024
PAGE_SIZE = 500
ERROR_MESSAGES = {
    'credentials': 'Сохраните логин и пароль в записи «ЕДДС / Добродел» в настройках администратора.',
    'login': 'Добродел не подтвердил вход. Проверьте логин и пароль записи «ЕДДС / Добродел».',
    'access': 'Добродел не предоставил доступ к операционному отчёту МинЖКХ. Проверьте права аккаунта.',
    'timeout': 'Добродел не ответил вовремя. Повторите загрузку и проверьте доступ с компьютера-сервера.',
    'network': 'Компьютер-сервер не смог подключиться к admin.vmeste.mosreg.ru. Проверьте сеть и настройки прокси.',
    'tls': 'Сертификат Добродела не прошёл проверку на сервере. Проверьте доверенные сертификаты; при необходимости задайте DOBRODEL_CA_BUNDLE.',
    'format': 'Добродел вернул неожиданный ответ вместо данных. Проверьте доступ аккаунта к операционному отчёту.',
    'size': 'Отчёт Добродела превысил допустимый размер. Сократите период сбора.',
    'redirect': 'Добродел перенаправил запрос на другой адрес. Передача остановлена; требуется проверка интеграции.',
    'config': 'Не удалось прочитать настройки доступа к Доброделу. Заново сохраните логин и пароль в админ-панели.',
}


class DobrodelError(RuntimeError):
    def __init__(self, code, status=None):
        self.code = code if code in ERROR_MESSAGES or code == 'http' else 'format'
        self.status = status
        message = ERROR_MESSAGES.get(self.code)
        if self.code == 'http':
            message = f'Добродел вернул HTTP {status}. Повторите позже или проверьте доступность портала.'
        super().__init__(message)


class DobrodelClient:
    def __init__(self, username, password):
        self.username, self.password = username, password
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Neurona-Dobrodel/3.0'})
        self.verify = os.getenv('DOBRODEL_CA_BUNDLE') or True
        self.deadline = None

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.session.close()

    @staticmethod
    def safe_url(target):
        try:
            url = urljoin(BASE + '/', target)
            parsed = urlsplit(url)
            port = parsed.port
        except ValueError:
            raise DobrodelError('redirect') from None
        if (parsed.scheme != 'https' or parsed.hostname != 'admin.vmeste.mosreg.ru'
                or port not in (None, 443) or parsed.username or parsed.password):
            raise DobrodelError('redirect')
        return url

    def _timeouts(self):
        if self.deadline is None:
            return (15, 60)
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise DobrodelError('timeout')
        return (min(15, remaining), min(60, remaining))

    def request(self, method, target, *, params=None, data=None, headers=None):
        url = self.safe_url(target)
        if method not in {'GET', 'POST'} or (method != 'POST' and data is not None):
            raise DobrodelError('config')
        try:
            for _ in range(6):
                response = self.session.request(method, url, params=params, data=data,
                    headers=headers, timeout=self._timeouts(), allow_redirects=False,
                    verify=self.verify, stream=True)
                with response:
                    if response.status_code in (301, 302, 303, 307, 308):
                        url = self.safe_url(urljoin(response.url, response.headers.get('Location', '')))
                        params = None
                        if response.status_code == 303 or (response.status_code in (301, 302) and method == 'POST'):
                            method, data = 'GET', None
                        continue
                    if response.status_code in (401, 403):
                        raise DobrodelError('login', response.status_code)
                    if response.status_code != 200:
                        raise DobrodelError('http', response.status_code)
                    content = bytearray()
                    for part in response.iter_content(65536):
                        self._timeouts()
                        content.extend(part)
                        if len(content) > MAX_BYTES:
                            raise DobrodelError('size')
                    response._content = bytes(content)
                    response._content_consumed = True
                    return response
        except requests.exceptions.SSLError:
            raise DobrodelError('tls') from None
        except requests.exceptions.Timeout:
            raise DobrodelError('timeout') from None
        except (requests.exceptions.RequestException, OSError):
            raise DobrodelError('network') from None
        raise DobrodelError('redirect')

    def login(self):
        if not self.username or not self.password:
            raise DobrodelError('credentials')
        response = self.request('GET', '/login')
        doc = BeautifulSoup(response.text, 'html.parser')
        form = doc.select_one('#loginForm')
        if form is None or not form.select_one('[name="j_password"]'):
            raise DobrodelError('format')
        fields = {node['name']: node.get('value', '') for node in form.select('input[type="hidden"][name]')}
        fields.update(j_username=self.username, j_password=self.password, _spring_security_remember_me='on')
        # The live portal's /js/login.js uses AJAX POST login/admin. Its form
        # has neither action nor method; native form.submit() would issue GET.
        response = self.request('POST', '/login/admin', data=fields, headers={
            'Accept': 'application/json', 'X-Requested-With': 'XMLHttpRequest',
            'Referer': BASE + '/login', 'Origin': BASE})
        try:
            result = response.json()
        except ValueError:
            raise DobrodelError('format') from None
        if not isinstance(result, dict):
            raise DobrodelError('format')
        if result.get('error'):
            raise DobrodelError('login')
        # Verify permissions on a fixed local URL; never follow a JSON show URL.
        report = self.request('GET', REPORT_PATH)
        doc = BeautifulSoup(report.text, 'html.parser')
        if doc.select_one('input[type="password"]') or '/login' in urlsplit(report.url).path:
            raise DobrodelError('login')
        if not doc.select_one('#curatorSelect'):
            raise DobrodelError('access')

    def fetch_all(self, filters):
        records, previous, reauthenticated = [], None, False
        for page in range(600):
            params = {'orderBy': 'ID', 'page': page, 'size': PAGE_SIZE, **filters}
            try:
                response = self.request('GET', '/report/operative', params=params, headers={
                    'Accept': 'application/json', 'X-Requested-With': 'XMLHttpRequest',
                    'Referer': BASE + REPORT_PATH})
                if '/login' in urlsplit(response.url).path:
                    raise DobrodelError('login')
            except DobrodelError as error:
                if error.code != 'login' or reauthenticated:
                    raise
                self.login()
                reauthenticated = True
                response = self.request('GET', '/report/operative', params=params,
                                        headers={'Accept': 'application/json', 'X-Requested-With': 'XMLHttpRequest'})
            try:
                batch = response.json()
            except ValueError:
                raise DobrodelError('format') from None
            if not isinstance(batch, list) or any(
                not isinstance(row, dict) or len(str(row.get('cardId', ''))) > 32
                or not str(row.get('cardId', '')).isascii()
                or not str(row.get('cardId', '')).isdecimal() or int(row['cardId']) <= 0
                for row in batch
            ):
                raise DobrodelError('format')
            if not batch:
                return records
            marker = (batch[0].get('cardId'), batch[-1].get('cardId'))
            if marker == previous and marker != (None, None):
                raise DobrodelError('format')
            previous = marker
            records.extend(batch)
            if len(batch) < PAGE_SIZE:
                return records
        raise DobrodelError('size')
