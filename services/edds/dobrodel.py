"""Dobrodel's current admin login and report API, with a private cookie session."""
import os
import time
from datetime import date, timedelta
from urllib.parse import urljoin, urlsplit

import requests
from bs4 import BeautifulSoup
from urllib3.exceptions import ReadTimeoutError

BASE = 'https://admin.vmeste.mosreg.ru'
REPORT_PATH = '/OperativeReportGenerating'
MAX_BYTES = 20 * 1024 * 1024
PAGE_SIZE = 500
REPORT_TIMEOUT = 240
REPORT_BUDGET = 20 * 60
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
        self.on_progress = None

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

    def _timeouts(self, read_timeout=60):
        if self.deadline is None:
            return (15, read_timeout)
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise DobrodelError('timeout')
        return (min(15, remaining), min(read_timeout, remaining))

    def request(self, method, target, *, params=None, data=None, headers=None):
        url = self.safe_url(target)
        read_timeout = REPORT_TIMEOUT if urlsplit(url).path == '/report/operative' else 60
        if method not in {'GET', 'POST'} or (method != 'POST' and data is not None):
            raise DobrodelError('config')
        try:
            for _ in range(6):
                response = self.session.request(method, url, params=params, data=data,
                    headers=headers, timeout=self._timeouts(read_timeout), allow_redirects=False,
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
        except requests.exceptions.ConnectionError as error:
            # Requests wraps a timeout during iter_content as ConnectionError;
            # retain its meaning so a streamed report can use smaller intervals.
            causes = (error.__cause__, error.__context__, *error.args)
            code = 'timeout' if any(isinstance(cause, ReadTimeoutError) for cause in causes) else 'network'
            raise DobrodelError(code) from None
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

    def _progress(self, stage, filters, page=0, count=0):
        if self.on_progress:
            bounds = self._bounds(filters)
            self.on_progress({'stage': stage, 'page': page + 1, 'count': count,
                'from': bounds[0].isoformat() if bounds else '',
                'to': (bounds[1] - timedelta(days=1)).isoformat() if bounds else ''})

    @staticmethod
    def _bounds(filters):
        try:
            start = date.fromisoformat(filters['filters.createdAfter'])
            stop = date.fromisoformat(filters['filters.createdBefore'])
            return (start, stop) if start < stop else None
        except (KeyError, ValueError, TypeError):
            return None

    @staticmethod
    def _batch(response):
        if '/login' in urlsplit(response.url).path:
            raise DobrodelError('login')
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
        return batch

    def _page(self, filters, page, size):
        response = self.request('GET', '/report/operative',
            params={**filters, 'orderBy': 'ID', 'page': page, 'size': size}, headers={
                'Accept': 'application/json', 'X-Requested-With': 'XMLHttpRequest',
                'Referer': BASE + REPORT_PATH})
        return self._batch(response)

    def probe_report(self, filters):
        """Confirm the report API, not merely permission to view its HTML page."""
        return len(self._page(filters, 0, 1))

    def _pages(self, filters, size):
        records, previous, reauthenticated = [], None, False
        for page in range(600):
            self._progress('report', filters, page, len(records))
            try:
                batch = self._page(filters, page, size)
            except DobrodelError as error:
                if error.code != 'login' or reauthenticated:
                    raise
                self.login()
                reauthenticated = True
                batch = self._page(filters, page, size)
            if not batch:
                return records
            marker = (batch[0].get('cardId'), batch[-1].get('cardId'))
            if marker == previous and marker != (None, None):
                raise DobrodelError('format')
            previous = marker
            records.extend(batch)
            if len(batch) < size:
                return records
        raise DobrodelError('size')

    def fetch_all(self, filters):
        """Fetch complete intervals; never change page size halfway through one.

        Slow /operative queries are split by date. At the smallest interval a
        single retry starts at page zero with smaller pages. All attempts share
        the caller's deadline, including the separate water-map integration.
        """
        previous_deadline = self.deadline
        self.deadline = min(previous_deadline if previous_deadline is not None else float('inf'),
                            time.monotonic() + REPORT_BUDGET)

        def collect(current):
            self._timeouts()
            bounds = self._bounds(current)
            try:
                return self._pages(current, PAGE_SIZE)
            except DobrodelError as error:
                retryable = error.code in {'timeout', 'size'} or (
                    error.code == 'http' and error.status in {502, 503, 504})
                if not retryable or bounds is None:
                    raise
                self._timeouts()  # A per-request timeout may split; an expired job may not.
                start, stop = bounds
                if (stop - start).days > 1:
                    self._progress('split', current)
                    middle = start + (stop - start) // 2
                    return collect({**current, 'filters.createdBefore': middle.isoformat()}) + collect(
                        {**current, 'filters.createdAfter': middle.isoformat()})
                self._progress('retry', current)
                # Back off briefly without permitting retries to extend the job.
                time.sleep(min(2, max(0, self.deadline - time.monotonic())))
                self._timeouts()
                return self._pages(current, min(100, PAGE_SIZE))

        try:
            records = collect(dict(filters))
            # Date boundaries on older portal versions can be inclusive.
            # Keep a card only once when two successful subintervals overlap.
            return list({str(int(row['cardId'])): row for row in records}.values())
        finally:
            self.deadline = previous_deadline
