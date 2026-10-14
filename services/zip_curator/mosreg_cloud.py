"""Read-only ownCloud client for the configured Mosreg ZIP stock folder.

ownCloud's Meta API resolves the UI fileid to the logged-in user's path. A
Depth:0 PROPFIND verifies that identity before any recursive listing. There is
no account-root fallback, redirect following, or download-URL delegation.
The caller owns the supplied requests.Session and closes it after the sync.
"""
import json
import math
import re
import time
from collections import deque
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import parse_qs, quote, unquote, urlsplit
from xml.etree import ElementTree as ET

import requests

from services.zip_curator.yandex_disk import (
    DiskSyncError, MAX_FILE_BYTES, SYNC_SECONDS, RETRY_STATUS, validate_workbook,
)


ORIGIN = 'https://cloud.mosreg.ru'
DEFAULT_SOURCE_URL = ORIGIN + '/apps/files/?dir=/&fileid=10968596'
ATTEMPTS = 3
MAX_XML_BYTES = 8 * 1024 * 1024
MAX_ENTRIES = 5000
MAX_DIRECTORIES = 500
MAX_DEPTH = 12
DAV = '{DAV:}'
OC = '{http://owncloud.org/ns}'
_CONTROL = re.compile(r'[\x00-\x1f\x7f]')
_ENCODED = re.compile(r'%[0-9a-f]{2}', re.I)
_WORKBOOK = re.compile(r'\.xlsx?$', re.I)
_PROPERTIES = ('resourcetype', 'displayname', 'getcontentlength', 'getetag',
               'creationdate', 'getlastmodified')


class CloudSyncError(DiskSyncError):
    """Safe to display: never contains a remote response, URL or credential."""


class _HTTPError(CloudSyncError):
    def __init__(self, status):
        self.status = status
        if status in {401, 403}:
            message = ('Облако отклонило доступ. Проверьте сохранённые логин, пароль '
                       'и доступ учётной записи к папке. При двухфакторной защите '
                       'может потребоваться пароль приложения.')
        elif 300 <= status < 400:
            message = 'Облако перенаправило запрос. Проверьте доступ WebDAV; данные входа не пересылались.'
        elif status == 412:
            message = 'Файл в облаке изменился во время обновления. Повторите обновление.'
        elif status == 404:
            message = 'Папка или файл не найдены в облаке для этой учётной записи.'
        else:
            message = 'Облако не выполнило запрос. Повторите обновление позже.'
        super().__init__(message)


def normalize_source_url(value):
    """Validate without network access; the numeric fileid determines identity."""
    if not isinstance(value, str) or len(value) > 4096 or _CONTROL.search(value):
        raise CloudSyncError('Укажите корректную ссылку на папку облака Московской области.')
    try:
        parsed = urlsplit(value.strip())
        params = parse_qs(parsed.query, keep_blank_values=True, max_num_fields=20)
        ids = params.get('fileid', [])
        if (parsed.scheme != 'https' or parsed.hostname != 'cloud.mosreg.ru'
                or parsed.port not in (None, 443) or parsed.username is not None
                or parsed.password is not None or parsed.fragment
                or parsed.path.rstrip('/') not in ('/apps/files', '/index.php/apps/files')
                or len(ids) != 1 or not re.fullmatch(r'[1-9][0-9]{0,18}', ids[0])):
            raise ValueError()
    except (ValueError, TypeError):
        raise CloudSyncError('Нужна ссылка HTTPS на папку облака Московской области с числовым fileid.') from None
    return ORIGIN + '/apps/files/?dir=/&fileid=' + ids[0]


def _segments(value):
    if (not isinstance(value, str) or len(value) > 4096 or _CONTROL.search(value)
            or '\\' in value or _ENCODED.search(value)):
        raise CloudSyncError('Облако вернуло недопустимый путь.')
    parts = value.strip('/').split('/')
    if any(part in ('', '.', '..') for part in parts):
        raise CloudSyncError('Облако вернуло недопустимый путь.')
    return tuple(parts)


def normalize_folder_path(value):
    """Optional decoded path inside the account, never an account-root fallback."""
    if value in (None, ''):
        return ''
    if (not isinstance(value, str) or len(value) > 2048 or value.startswith('//')
            or '://' in value):
        raise CloudSyncError('Укажите путь к конкретной папке внутри облака, без адреса сайта.')
    try:
        parts = _segments(value)
    except CloudSyncError:
        raise CloudSyncError('Путь должен указывать на конкретную папку без переходов к родительским каталогам.') from None
    return '/' + '/'.join(parts)


def _encoded_path(parts):
    return '/' + '/'.join(quote(part, safe='') for part in parts)


def _href_parts(href):
    if not isinstance(href, str) or not href or _CONTROL.search(href):
        raise CloudSyncError('Облако вернуло недопустимую ссылку на файл.')
    try:
        parsed = urlsplit(href)
        if (parsed.query or parsed.fragment or parsed.username is not None
                or parsed.password is not None or parsed.port not in (None, 443)
                or (parsed.netloc and (parsed.hostname != 'cloud.mosreg.ru' or parsed.scheme != 'https'))
                or (parsed.scheme and (parsed.scheme != 'https' or not parsed.netloc))
                or not parsed.path.startswith('/') or parsed.path.startswith('//')
                or re.search(r'%(?:2f|5c|00)', parsed.path, re.I)
                or re.search(r'%(?![0-9a-f]{2})', parsed.path, re.I)):
            raise ValueError()
        return _segments(unquote(parsed.path, encoding='utf-8', errors='strict'))
    except (ValueError, UnicodeError):
        raise CloudSyncError('Облако вернуло недопустимую ссылку на файл.') from None


def _xml(payload):
    try:
        text = payload.decode('utf-8-sig')
        if '\x00' in text or re.search(r'<!\s*(?:DOCTYPE|ENTITY)\b', text, re.I):
            raise ValueError()
        root = ET.fromstring(text)
        if root.tag != DAV + 'multistatus':
            raise ValueError()
        responses = root.findall(DAV + 'response')
        if len(responses) > MAX_ENTRIES + 1:
            raise CloudSyncError('В облачной папке слишком много записей для одного обновления.')
        return responses
    except (UnicodeError, ValueError, ET.ParseError):
        raise CloudSyncError('Облако вернуло некорректный список файлов. Локальные данные сохранены.') from None


def _properties(response):
    direct_status = response.findtext(DAV + 'status')
    if direct_status and not re.fullmatch(r'HTTP/\S+ 2\d\d(?: .*|)', direct_status.strip()):
        raise CloudSyncError('Получен неполный список облачной папки. Локальные данные сохранены.')
    properties = {}
    for part in response.findall(DAV + 'propstat'):
        status = part.findtext(DAV + 'status') or ''
        if re.fullmatch(r'HTTP/\S+ 2\d\d(?: .*|)', status.strip()):
            prop = part.find(DAV + 'prop')
            if prop is not None:
                for child in prop:
                    if child.tag in properties:
                        raise CloudSyncError('Облако вернуло неоднозначные свойства файла.')
                    properties[child.tag] = child
        elif not re.fullmatch(r'HTTP/\S+ 404(?: .*|)', status.strip()):
            raise CloudSyncError('Получен неполный список облачной папки. Локальные данные сохранены.')
    return properties


def _text(properties, name):
    value = properties.get(name)
    return (value.text or '').strip() if value is not None else ''


def _date(value):
    if not value:
        return None
    try:
        try:
            result = datetime.fromisoformat(value.replace('Z', '+00:00'))
        except ValueError:
            result = parsedate_to_datetime(value)
        if result.tzinfo is None:
            return None
        return result.astimezone(timezone.utc).isoformat(timespec='seconds')
    except (ValueError, TypeError, OverflowError):
        return None


def _propfind_body(*, meta=False):
    root = ET.Element(DAV + 'propfind')
    prop = ET.SubElement(root, DAV + 'prop')
    if meta:
        ET.SubElement(prop, OC + 'meta-path-for-user')
    else:
        for name in _PROPERTIES:
            ET.SubElement(prop, DAV + name)
        ET.SubElement(prop, OC + 'fileid')
        ET.SubElement(prop, OC + 'checksums')
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


class CloudClient:
    def __init__(self, session, source_url, timeout=45, *, username, password, folder_path=None):
        self.source_url = normalize_source_url(source_url)
        self.fileid = parse_qs(urlsplit(self.source_url).query)['fileid'][0]
        self.folder_path = normalize_folder_path(folder_path)
        if (not isinstance(username, str) or not username.strip() or len(username) > 255
                or ':' in username or _CONTROL.search(username)
                or not isinstance(password, str) or not password or len(password) > 4096
                or _CONTROL.search(password)):
            raise CloudSyncError('Сохраните логин и пароль доступа к облаку в панели администратора.')
        self._auth = (username.strip(), password)
        self.session = session
        self.session.trust_env = False
        try:
            timeout = float(timeout)
            self.timeout = min(90.0, max(1.0, timeout)) if math.isfinite(timeout) else 45.0
        except (TypeError, ValueError):
            self.timeout = 45.0
        # A cooperative sync budget, checked between requests/received blocks.
        # Requests' socket timeout limits inactivity, not total transfer time.
        self.deadline = time.monotonic() + SYNC_SECONDS
        self._user = None
        self._root = None
        self._listed = {}

    def _timeout(self):
        left = self.deadline - time.monotonic()
        if left <= 0:
            raise CloudSyncError('Время обновления облачной папки истекло. Повторите обновление.')
        return min(10.0, left), min(self.timeout, left)

    def _pause(self, attempt):
        time.sleep(min(0.5 * 2 ** attempt, max(0, self.deadline - time.monotonic())))

    def _request(self, method, path, *, body=None, headers=None, limit=MAX_XML_BYTES, expected=207):
        # Call sites supply only constants or paths encoded from validated parts.
        if not path.startswith('/') or path.startswith('//'):
            raise CloudSyncError('Недопустимый путь запроса к облаку.')
        request_headers = {'Accept': 'application/xml', 'Content-Type': 'application/xml; charset=utf-8',
                           **(headers or {})}
        for attempt in range(ATTEMPTS):
            try:
                with self.session.request(method, ORIGIN + path, data=body, headers=request_headers,
                                          auth=self._auth, timeout=self._timeout(), stream=True,
                                          allow_redirects=False, verify=True) as response:
                    if response.status_code != expected:
                        if response.status_code in RETRY_STATUS and attempt < ATTEMPTS - 1:
                            self._pause(attempt)
                            continue
                        raise _HTTPError(response.status_code)
                    length = response.headers.get('Content-Length')
                    if length is not None:
                        try:
                            if int(length) < 0 or int(length) > limit:
                                raise CloudSyncError('Ответ облака превышает допустимый размер.')
                        except ValueError:
                            raise CloudSyncError('Облако вернуло некорректный размер ответа.') from None
                    content = bytearray()
                    for chunk in response.iter_content(chunk_size=65536):
                        self._timeout()
                        content.extend(chunk)
                        if len(content) > limit:
                            raise CloudSyncError('Ответ облака превышает допустимый размер.')
                    return bytes(content)
            except requests.exceptions.SSLError:
                raise CloudSyncError('Не удалось проверить сертификат облака. Проверьте доверенные сертификаты сервера.') from None
            except (requests.RequestException, UnicodeError):
                if attempt == ATTEMPTS - 1:
                    raise CloudSyncError('Соединение с облаком прервалось или истекло время ожидания. Выполнено три попытки.') from None
                self._pause(attempt)

    def _current_user(self):
        if self._user is not None:
            return
        raw = self._request('GET', '/ocs/v1.php/cloud/user?format=json', expected=200, limit=65536,
                            headers={'OCS-APIRequest': 'true', 'Accept': 'application/json'})
        try:
            value = json.loads(raw)['ocs']
            if str(value['meta']['statuscode']) not in {'100', '200'}:
                raise ValueError()
            uid = value['data']['id']
            if not isinstance(uid, str) or '/' in uid:
                raise ValueError()
            parts = _segments(uid)
            if len(parts) != 1 or len(uid) > 255:
                raise ValueError()
            self._user = parts[0]
        except (ValueError, TypeError, KeyError, CloudSyncError):
            raise CloudSyncError('Не удалось определить учётную запись облака. Проверьте доступ через WebDAV.') from None

    def _dav_parts(self, relative):
        return ('remote.php', 'dav', 'files', self._user, *relative)

    def _entry(self, response):
        parts = _href_parts(response.findtext(DAV + 'href'))
        prefix = self._dav_parts(())
        if parts[:len(prefix)] != prefix or len(parts) <= len(prefix):
            raise CloudSyncError('Облако вернуло файл за пределами выбранной учётной записи.')
        relative = parts[len(prefix):]
        props = _properties(response)
        resource = props.get(DAV + 'resourcetype')
        if resource is None:
            raise CloudSyncError('Получен неполный список облачной папки. Локальные данные сохранены.')
        kind = 'dir' if resource.find(DAV + 'collection') is not None else 'file'
        size = None
        if kind == 'file':
            raw_size = _text(props, DAV + 'getcontentlength')
            if not re.fullmatch(r'\d{1,18}', raw_size):
                raise CloudSyncError('Облако не указало допустимый размер файла.')
            size = int(raw_size)
        etag = _text(props, DAV + 'getetag')
        if len(etag) > 256 or _CONTROL.search(etag):
            raise CloudSyncError('Облако вернуло некорректную версию файла.')
        item = {'type': kind, 'name': relative[-1], 'path': '/' + '/'.join(relative),
                'size': size, 'etag': etag, 'created': _date(_text(props, DAV + 'creationdate')),
                'modified': _date(_text(props, DAV + 'getlastmodified')),
                'fileid': _text(props, OC + 'fileid')}
        checksums = props.get(OC + 'checksums')
        if checksums is not None:
            for child in checksums.findall(OC + 'checksum'):
                match = re.fullmatch(r'MD5:([0-9a-f]{32})', (child.text or '').strip(), re.I)
                if match:
                    item['md5'] = match.group(1).lower()
        return relative, item

    def _folder_entries(self, relative, depth):
        raw = self._request('PROPFIND', _encoded_path(self._dav_parts(relative)) + '/',
                            body=_propfind_body(), headers={'Depth': str(depth)})
        result = [self._entry(response) for response in _xml(raw)]
        seen = set()
        for path, _ in result:
            if path in seen or (path != relative and (depth == 0 or path[:-1] != relative)):
                raise CloudSyncError('Облако вернуло неоднозначный или неполный список папки.')
            seen.add(path)
        own = [item for path, item in result if path == relative]
        if len(own) != 1 or own[0]['type'] != 'dir':
            raise CloudSyncError('Облако не подтвердило выбранную папку. Локальные данные сохранены.')
        return result

    def _resolve(self):
        self._current_user()
        relative = None
        if self.folder_path:
            relative = _segments(self.folder_path)
        else:
            meta_path = '/remote.php/dav/meta/' + self.fileid
            try:
                raw = self._request('PROPFIND', meta_path, body=_propfind_body(meta=True), headers={'Depth': '0'})
            except _HTTPError as exc:
                if exc.status in {400, 404, 405, 501}:
                    raise CloudSyncError('Не удалось найти папку по fileid. Укажите её точный путь в настройках облака; корень аккаунта не используется.') from None
                raise
            responses = _xml(raw)
            if len(responses) != 1 or _href_parts(responses[0].findtext(DAV + 'href')) != ('remote.php', 'dav', 'meta', self.fileid):
                raise CloudSyncError('Облако не подтвердило папку по fileid.')
            path = _text(_properties(responses[0]), OC + 'meta-path-for-user')
            if not path:
                raise CloudSyncError('Облако не вернуло путь папки. Укажите точный путь в настройках облака.')
            relative = _segments(normalize_folder_path(path))
        properties = self._folder_entries(relative, 0)[0][1]
        if properties['fileid'] != self.fileid:
            raise CloudSyncError('Указанный путь не соответствует fileid выбранной папки. Проверьте настройки облака.')
        self._root = relative

    def list_files(self):
        self._listed = {}
        self._resolve()
        queue = deque([(self._root, 0)])
        seen, directories, files = {self._root}, 0, []
        while queue:
            parent, depth = queue.popleft()
            directories += 1
            if directories > MAX_DIRECTORIES:
                raise CloudSyncError('В облачной папке слишком много вложенных каталогов.')
            for path, item in self._folder_entries(parent, 1):
                if path == parent:
                    if parent == self._root and item['fileid'] != self.fileid:
                        raise CloudSyncError('Выбранная облачная папка изменилась во время обновления.')
                    continue
                if path in seen:
                    raise CloudSyncError('Список облачной папки содержит повторяющиеся записи.')
                seen.add(path)
                if len(seen) > MAX_ENTRIES:
                    raise CloudSyncError('В облачной папке слишком много записей для одного обновления.')
                if item['type'] == 'dir':
                    if depth >= MAX_DEPTH:
                        raise CloudSyncError('Вложенность облачной папки превышает допустимую глубину.')
                    queue.append((path, depth + 1))
                elif _WORKBOOK.search(item['name']):
                    files.append(item)
        self._listed = {item['path']: dict(item) for item in files}
        return files

    def download(self, item, read_rows):
        # Only an entry from the completed listing may be downloaded; mutable
        # caller fields cannot inject another URL, path, size or checksum.
        path = item.get('path') if isinstance(item, dict) else None
        listed = self._listed.get(path) if isinstance(path, str) else None
        if listed is None or self._root is None:
            raise CloudSyncError('Файл отсутствует в подтверждённом списке облачной папки.')
        if listed['size'] > MAX_FILE_BYTES:
            raise CloudSyncError('Файл превышает допустимый размер 50 МБ.')
        relative = _segments(path)
        if relative[:len(self._root)] != self._root or len(relative) <= len(self._root):
            raise CloudSyncError('Файл находится за пределами выбранной облачной папки.')
        headers = {'Accept': 'application/octet-stream'}
        if listed['etag']:
            headers['If-Match'] = listed['etag']
        content = self._request('GET', _encoded_path(self._dav_parts(relative)),
                                headers=headers, limit=MAX_FILE_BYTES, expected=200)
        try:
            validate_workbook(content, listed, read_rows)
        except DiskSyncError as exc:
            raise CloudSyncError(str(exc)) from None
        return content
