"""Read ARM reports using the saved integration account, entirely on the server."""
import csv
import io
import os
import threading
import time
from datetime import date
from urllib.parse import urljoin, urlsplit

import requests
from bs4 import BeautifulSoup

from services.auth.integrations import credentials

BASE_URL = 'https://zkh-kontur.mosreg.ru/new9/'
MAX_BYTES = 20 * 1024 * 1024
LOCK = threading.Lock()


class ArmError(Exception):
    def __init__(self, message, status=502):
        super().__init__(message)
        self.status = status


def validate_period(start: date, end: date):
    if start > end:
        raise ArmError('Начало периода должно быть не позднее окончания.', 400)
    if (end - start).days > 366:
        raise ArmError('Выберите период не больше года. Для первой загрузки достаточно недели.', 400)


def parse_csv(text, required):
    if '<html' in text[:1000].lower() or '<!doctype' in text[:1000].lower():
        raise ArmError('АРМ ЕДДС вернул страницу вместо отчёта. Проверьте права доступа к отчётам.')
    text = text.lstrip('\ufeff')
    first = text.splitlines()[0] if text.strip() else ''
    delimiter = ';' if first.count(';') >= first.count(',') else ','
    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    header = next((i for i, row in enumerate(rows[:10]) if required in [c.strip() for c in row]), None)
    if header is None:
        raise ArmError('Формат отчёта АРМ ЕДДС не распознан: отсутствует колонка «' + required + '».')
    return rows[header:]


class ArmClient:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers['User-Agent'] = 'Neurona-EDDS/3.0'
        self.verify = os.getenv('EDDS_ARM_CA_BUNDLE') or True
        self.deadline = time.monotonic() + 150

    def close(self):
        self.session.close()

    def safe_url(self, target):
        url = urljoin(BASE_URL, target)
        parsed = urlsplit(url)
        if (parsed.scheme != 'https' or parsed.hostname != urlsplit(BASE_URL).hostname
                or parsed.port not in (None, 443) or parsed.username or parsed.password):
            raise ArmError('АРМ ЕДДС перенаправил запрос на другой адрес. Требуется проверка настройки портала.')
        return url

    def request(self, method, url, data=None):
        url = self.safe_url(url)
        for _ in range(6):
            remaining = self.deadline - time.monotonic()
            if remaining <= 0:
                raise ArmError('АРМ ЕДДС не успел сформировать отчёт. Сократите период и повторите запрос.', 504)
            try:
                response = self.session.request(method, url, data=data, allow_redirects=False,
                                                timeout=(min(10, remaining), min(45, remaining)),
                                                verify=self.verify, stream=True)
                with response:
                    if response.status_code in (301, 302, 303, 307, 308):
                        url = self.safe_url(urljoin(url, response.headers.get('Location', '')))
                        if response.status_code == 303 or response.status_code in (301, 302) and method == 'POST':
                            method, data = 'GET', None
                        continue
                    if response.status_code in (401, 403):
                        raise ArmError('АРМ ЕДДС отказал в доступе. Проверьте логин, пароль и права на отчёты.', 403)
                    if not response.ok:
                        raise ArmError(f'АРМ ЕДДС ответил кодом {response.status_code}. Повторите позднее.')
                    chunks, size = [], 0
                    for chunk in response.iter_content(65536):
                        size += len(chunk)
                        if size > MAX_BYTES or time.monotonic() > self.deadline:
                            raise ArmError('Отчёт слишком большой или загружается слишком долго. Сократите период.', 413)
                        chunks.append(chunk)
                    raw = b''.join(chunks)
                    try:
                        text = raw.decode('utf-8-sig')
                    except UnicodeDecodeError:
                        text = raw.decode('cp1251')
                    return text, url
            except requests.exceptions.SSLError as error:
                if 'CERTIFICATE_VERIFY_FAILED' in str(error).upper():
                    raise ArmError('Сертификат АРМ ЕДДС не прошёл проверку на сервере «Нейроны». '
                                   'Проверьте цепочку сертификатов; доверенный файл центра сертификации '
                                   'можно указать в EDDS_ARM_CA_BUNDLE.') from None
                raise ArmError('Не удалось установить защищённое соединение с АРМ ЕДДС: '
                               'соединение прервано до входа. Проверьте доступ к порталу с компьютера-сервера.') from None
            except requests.exceptions.Timeout:
                raise ArmError('АРМ ЕДДС не ответил с компьютера-сервера. Проверьте доступ к порталу и сократите период.', 504) from None
            except (requests.exceptions.RequestException, OSError):
                raise ArmError('Компьютер-сервер не смог подключиться к АРМ ЕДДС. Проверьте сеть и настройки сертификатов.') from None
        raise ArmError('Слишком много перенаправлений при входе в АРМ ЕДДС.')

    def login(self, account):
        text, url = self.request('GET', BASE_URL + '?act=cds_report_svod&id=3608')
        doc = BeautifulSoup(text, 'html.parser')
        password = doc.select_one('input[type="password"]')
        if not password:
            # An authenticated report is the only acceptable passwordless response.
            if doc.select_one('[name="date_ot"]') and doc.select_one('[name="date_do"]'):
                return
            raise ArmError('Форма входа АРМ ЕДДС не найдена. Возможно, изменился адрес или способ входа.')
        form = password.find_parent('form')
        if not form or not password.get('name'):
            raise ArmError('Форма входа АРМ ЕДДС изменилась. Нужна проверка интеграции.')
        username = form.select_one('input[autocomplete="username"], input[type="email"]')
        if not username:
            username = next((node for node in form.select('input[name]')
                             if node.get('type', 'text').lower() == 'text'
                             and not any(word in node['name'].lower() for word in ('captcha', 'code', 'otp'))), None)
        if not username or not username.get('name'):
            raise ArmError('Поле логина АРМ ЕДДС не найдено. Нужна проверка интеграции.')
        if form.select_one('input[name*="captcha"], iframe[src*="captcha"], .g-recaptcha'):
            raise ArmError('Портал требует капчу. Автоматический вход с сервера сейчас недоступен.', 409)
        fields = {node['name']: node.get('value', '') for node in form.select('input[type="hidden"][name]')}
        submit = form.select_one('button[name], input[type="submit"][name]')
        if submit:
            fields[submit['name']] = submit.get('value', '')
        fields[username['name']] = account['username']
        fields[password['name']] = account['password']
        # Always POST credentials; never put them in URLs or logs.
        text, _ = self.request('POST', urljoin(url, form.get('action') or url), fields)
        if BeautifulSoup(text, 'html.parser').select_one('input[type="password"]'):
            raise ArmError('Вход в АРМ ЕДДС не выполнен. Проверьте сохранённый логин и пароль; возможна дополнительная проверка.', 403)

    def report(self, start, end, coordinates=False):
        if coordinates:
            url = BASE_URL + '?act=cds_claim_report&id=3654'
            fields = {'start': start.strftime('%d.%m.%Y'), 'end': end.strftime('%d.%m.%Y'), 'submit': '1',
                      'fields[id_cds_claim]': 'on', 'fields[lat_]': 'on', 'fields[lon_]': 'on'}
        else:
            url = BASE_URL + '?act=cds_report_svod&id=3608'
            fields = {'act': 'cds_report_svod', 'id': '3608', 'date_ot': start.strftime('%d.%m.%y'),
                      'date_do': end.strftime('%d.%m.%y'), 'id_tu_mun_raion': '0', 'saveToCSV': '1'}
        text, current = self.request('POST', url, fields)
        if coordinates and ('<html' in text[:1000].lower() or '<!doctype' in text[:1000].lower()):
            doc = BeautifulSoup(text, 'html.parser')
            link = next((a['href'] for a in doc.select('a[href]')
                         if urlsplit(a['href']).path.lower().endswith('.csv')), None)
            if link:
                text, _ = self.request('GET', urljoin(current, link))
        return parse_csv(text, 'id_cds_claim')


def fetch_report(start, end, coordinates=False):
    validate_period(start, end)
    account = credentials('edds_arm')
    if not account:
        raise ArmError('Сохраните логин и пароль для «АРМ ЕДДС» в разделе «Пользователи → Логины и пароли».', 400)
    if not LOCK.acquire(blocking=False):
        raise ArmError('Запрос к АРМ ЕДДС уже выполняется. Дождитесь его завершения.', 409)
    client = None
    try:
        client = ArmClient()
        client.login(account)
        return client.report(start, end, coordinates)
    finally:
        if client:
            client.close()
        LOCK.release()
