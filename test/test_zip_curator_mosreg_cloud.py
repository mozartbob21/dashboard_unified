"""Offline ownCloud contracts, bounds and credential/path containment."""
import hashlib
import io
import json
import zipfile
from urllib.parse import quote, urlsplit
from xml.etree import ElementTree as ET

import pytest
import requests

from services.zip_curator import mosreg_cloud as cloud


UID = 'actual-user-id'
ROOT = '/Запасы ЗИП'
SOURCE = cloud.DEFAULT_SOURCE_URL
META = '/remote.php/dav/meta/10968596'
USER = '/ocs/v1.php/cloud/user?format=json'
LOGIN = 'login@example.test'
PASSWORD = 'synthetic-private-password'


def dav(path):
    return '/remote.php/dav/files/' + UID + quote(path, safe='/')


def response_xml(entries):
    root = ET.Element(cloud.DAV + 'multistatus')
    for entry in entries:
        response = ET.SubElement(root, cloud.DAV + 'response')
        ET.SubElement(response, cloud.DAV + 'href').text = entry.get('href', dav(entry['path']))
        propstat = ET.SubElement(response, cloud.DAV + 'propstat')
        props = ET.SubElement(propstat, cloud.DAV + 'prop')
        resource = ET.SubElement(props, cloud.DAV + 'resourcetype')
        if entry.get('dir'):
            ET.SubElement(resource, cloud.DAV + 'collection')
        else:
            ET.SubElement(props, cloud.DAV + 'getcontentlength').text = str(entry.get('size', 0))
        ET.SubElement(props, cloud.OC + 'fileid').text = entry.get('id', '10968596')
        ET.SubElement(props, cloud.DAV + 'getetag').text = entry.get('etag', '"opaque-version"')
        ET.SubElement(props, cloud.DAV + 'creationdate').text = '2026-09-20T09:00:00Z'
        ET.SubElement(props, cloud.DAV + 'getlastmodified').text = 'Tue, 13 Oct 2026 10:11:12 GMT'
        if entry.get('md5'):
            checksums = ET.SubElement(props, cloud.OC + 'checksums')
            ET.SubElement(checksums, cloud.OC + 'checksum').text = 'MD5:' + entry['md5']
        ET.SubElement(propstat, cloud.DAV + 'status').text = entry.get('status', 'HTTP/1.1 200 OK')
    return ET.tostring(root, encoding='utf-8')


def meta_xml(path=ROOT, *, href=META + '/'):
    root = ET.Element(cloud.DAV + 'multistatus')
    response = ET.SubElement(root, cloud.DAV + 'response')
    ET.SubElement(response, cloud.DAV + 'href').text = href
    propstat = ET.SubElement(response, cloud.DAV + 'propstat')
    props = ET.SubElement(propstat, cloud.DAV + 'prop')
    ET.SubElement(props, cloud.OC + 'meta-path-for-user').text = path
    ET.SubElement(propstat, cloud.DAV + 'status').text = 'HTTP/1.1 200 OK'
    return ET.tostring(root, encoding='utf-8')


class Response:
    def __init__(self, content=b'', status=207, headers=None):
        self.content = content
        self.status_code = status
        self.headers = {'Content-Length': str(len(content)), **(headers or {})}
        self.closed = False

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.closed = True

    def iter_content(self, chunk_size):
        for offset in range(0, len(self.content), chunk_size):
            yield self.content[offset:offset + chunk_size]


class Session:
    def __init__(self):
        self.calls, self.routes = [], {}
        self.trust_env = True

    def add(self, method, path, responses, depth=None):
        self.routes[(method, cloud.ORIGIN + path, depth)] = list(responses)

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        assert url.startswith(cloud.ORIGIN + '/')
        assert kwargs['allow_redirects'] is False and kwargs['verify'] is True
        assert kwargs['auth'] == (LOGIN, PASSWORD)
        assert self.trust_env is False and kwargs['stream'] is True
        assert 0 < kwargs['timeout'][0] <= 10 and 0 < kwargs['timeout'][1] <= 90
        key = method, url, kwargs['headers'].get('Depth')
        assert self.routes.get(key), 'Unexpected offline request'
        value = self.routes[key].pop(0)
        if isinstance(value, Exception):
            raise value
        return value


@pytest.fixture
def session(monkeypatch):
    monkeypatch.setattr(cloud.CloudClient, '_pause', lambda *args: None)
    result = Session()
    result.add('GET', USER, [Response(json.dumps({'ocs': {'meta': {'statuscode': 100},
                                                         'data': {'id': UID}}}).encode(), status=200)])
    result.add('PROPFIND', META, [Response(meta_xml())], depth='0')
    result.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([{'path': ROOT, 'dir': True}]))], depth='0')
    return result


def client(session, **kwargs):
    return cloud.CloudClient(session, SOURCE, username=LOGIN, password=PASSWORD, **kwargs)


def workbook():
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, 'w') as archive:
        archive.writestr('[Content_Types].xml', '<Types/>')
        archive.writestr('xl/worksheets/sheet1.xml', '<worksheet/>')
    return stream.getvalue()


def test_resolves_real_owncloud_fileid_then_recursively_lists_and_downloads(session):
    data = workbook()
    folder = ROOT + '/Вода и тепло'
    first = ROOT + '/Реестр #1.xlsx'
    nested = folder + '/Реестр #1.xlsx'
    digest = hashlib.md5(data).hexdigest()
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': folder, 'dir': True, 'id': '2'},
        {'path': first, 'size': len(data), 'id': '3', 'md5': digest},
        {'path': ROOT + '/инструкция.txt', 'size': 2, 'id': '4'},
    ]))], depth='1')
    session.add('PROPFIND', dav(folder) + '/', [Response(response_xml([
        {'path': folder, 'dir': True, 'id': '2'}, {'path': nested, 'size': len(data), 'id': '5'},
    ]))], depth='1')
    downloaded = Response(data, status=200)
    session.add('GET', dav(first), [downloaded])
    api = client(session)
    items = api.list_files()
    assert {item['path'] for item in items} == {first, nested}
    assert all(item['type'] == 'file' and item['name'] == 'Реестр #1.xlsx' for item in items)
    item = next(item for item in items if item['path'] == first)
    assert item['created'] == '2026-09-20T09:00:00+00:00'
    assert item['modified'] == '2026-10-13T10:11:12+00:00'
    assert item['etag'] == '"opaque-version"' and item['md5'] == digest
    assert api.download(item, lambda payload: None) == data
    assert downloaded.closed
    assert session.calls[-1][2]['headers']['If-Match'] == '"opaque-version"'
    assert all(method != 'SEARCH' for method, _, _ in session.calls)
    meta_body = ET.fromstring(session.calls[1][2]['data'])
    assert meta_body.find('./' + cloud.DAV + 'prop/' + cloud.OC + 'meta-path-for-user') is not None
    assert not any(url == cloud.ORIGIN + '/remote.php/dav/files/' + UID + '/' for _, url, _ in session.calls)


def test_canonical_urls_and_optional_plain_folder_paths_are_network_free():
    assert cloud.normalize_source_url('https://cloud.mosreg.ru/index.php/apps/files/?fileid=10968596&dir=%2Fother') == SOURCE
    assert cloud.normalize_folder_path(None) == ''
    assert cloud.normalize_folder_path('Общие документы/Запасы/') == '/Общие документы/Запасы'


@pytest.mark.parametrize('url', [
    'http://cloud.mosreg.ru/apps/files/?fileid=10968596',
    'https://evil.test/apps/files/?fileid=10968596',
    'https://cloud.mosreg.ru.evil.test/apps/files/?fileid=10968596',
    'https://secret:password@cloud.mosreg.ru/apps/files/?fileid=10968596',
    'https://cloud.mosreg.ru:444/apps/files/?fileid=10968596',
    'https://cloud.mosreg.ru/apps/files/?dir=/',
    'https://cloud.mosreg.ru/apps/files/?fileid=abc',
    'https://cloud.mosreg.ru/apps/files/?fileid=10968596&fileid=1',
    'https://cloud.mosreg.ru/apps/files/?fileid=10968596#other',
    'https://cloud.mosreg.ru/login?fileid=10968596',
])
def test_source_url_rejects_foreign_hosts_ambiguous_identity_and_credentials(url):
    with pytest.raises(cloud.CloudSyncError) as caught:
        cloud.normalize_source_url(url)
    assert url not in str(caught.value) and 'password' not in str(caught.value)


@pytest.mark.parametrize('path', ['/', '..', '/a/../b', '/a//b', '//evil.test/path',
                                  'https://evil.test/path', '/a\\b', '/%2e%2e/a', '/a\x00b'])
def test_explicit_folder_rejects_root_traversal_and_encoded_ambiguity(path):
    with pytest.raises(cloud.CloudSyncError):
        cloud.normalize_folder_path(path)


def test_explicit_folder_is_verified_by_fileid_without_meta_request(session):
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([{'path': ROOT, 'dir': True}]))], depth='1')
    assert client(session, folder_path=ROOT).list_files() == []
    assert all(url != cloud.ORIGIN + META for _, url, _ in session.calls)


def test_explicit_folder_cannot_replace_requested_fileid_with_another_folder(session):
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([{'path': ROOT, 'dir': True, 'id': '999'}]))], depth='0')
    with pytest.raises(cloud.CloudSyncError, match='не соответствует fileid'):
        client(session, folder_path=ROOT).list_files()
    assert len(session.calls) == 2


@pytest.mark.parametrize('status', [404, 405, 501])
def test_unsupported_meta_api_never_falls_back_to_account_root(session, status):
    session.add('PROPFIND', META, [Response(b'PRIVATE_BACKEND_RESPONSE', status=status)], depth='0')
    with pytest.raises(cloud.CloudSyncError, match='точный путь'):
        client(session).list_files()
    assert len(session.calls) == 2


@pytest.mark.parametrize('path', ['/', '/a/../b', '//evil.test/a', '/%2e%2e/secret'])
def test_untrusted_meta_paths_do_not_select_another_root(session, path):
    session.add('PROPFIND', META, [Response(meta_xml(path))], depth='0')
    with pytest.raises(cloud.CloudSyncError):
        client(session).list_files()
    assert len(session.calls) == 2


@pytest.mark.parametrize('href', [
    'https://evil.test/remote.php/dav/files/actual-user-id/a.xlsx',
    '//evil.test/remote.php/dav/files/actual-user-id/a.xlsx',
    'http://cloud.mosreg.ru/remote.php/dav/files/actual-user-id/a.xlsx',
    '/remote.php/dav/files/another-user/a.xlsx',
    '/remote.php/dav/files/actual-user-id/Other/a.xlsx',
    '/remote.php/dav/files/actual-user-id/%2e%2e/a.xlsx',
    '/remote.php/dav/files/actual-user-id/%252e%252e/a.xlsx',
    '/remote.php/dav/files/actual-user-id/a%2fb.xlsx',
    '/remote.php/dav/files/actual-user-id/a%5cb.xlsx',
    '/remote.php/dav/files/actual-user-id/a.xlsx?token=private',
])
def test_listing_hrefs_cannot_escape_host_user_or_selected_folder(session, href):
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': ROOT + '/a.xlsx', 'href': href, 'size': 10},
    ]))], depth='1')
    with pytest.raises(cloud.CloudSyncError):
        client(session).list_files()
    assert len(session.calls) == 4


@pytest.mark.parametrize('status', [301, 302, 307, 308, 401, 403])
def test_auth_and_redirect_errors_are_safe_and_never_retry_or_follow(session, status):
    session.add('GET', USER, [Response(b'PRIVATE_RESPONSE ' + PASSWORD.encode(), status=status,
                                      headers={'Location': 'https://evil.test/'})])
    with pytest.raises(cloud.CloudSyncError) as caught:
        client(session).list_files()
    assert len(session.calls) == 1
    assert all(secret not in str(caught.value) for secret in (PASSWORD, LOGIN, 'evil.test', 'PRIVATE_RESPONSE'))


@pytest.mark.parametrize('payload', [
    b'<!DOCTYPE d:multistatus [<!ENTITY ex SYSTEM "file:///etc/passwd">]><d:multistatus xmlns:d="DAV:">&ex;</d:multistatus>',
    b'<!DOCTYPE d:multistatus [<!ENTITY ex "repeated">]><d:multistatus xmlns:d="DAV:">&ex;</d:multistatus>',
    b'<html><body>private sign-in page</body></html>',
    b'<d:multistatus xmlns:d="DAV:"><d:response>',
])
def test_xml_rejects_entities_html_and_incomplete_lists(session, payload):
    session.add('PROPFIND', META, [Response(payload)], depth='0')
    with pytest.raises(cloud.CloudSyncError):
        client(session).list_files()


def test_partial_or_missing_folder_response_cannot_be_reported_as_empty_success(session):
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([]))], depth='1')
    with pytest.raises(cloud.CloudSyncError, match='не подтвердило'):
        client(session).list_files()


def test_failed_child_properties_abort_complete_listing(session):
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': ROOT + '/a.xlsx', 'status': 'HTTP/1.1 403 Forbidden'},
    ]))], depth='1')
    with pytest.raises(cloud.CloudSyncError, match='неполный'):
        client(session).list_files()


def test_duplicate_entries_abort_listing(session):
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': ROOT + '/a.xlsx'}, {'path': ROOT + '/a.xlsx'},
    ]))], depth='1')
    with pytest.raises(cloud.CloudSyncError):
        client(session).list_files()


def test_missing_optional_properties_do_not_make_successful_listing_incomplete(session):
    xml = ET.fromstring(response_xml([{'path': ROOT, 'dir': True}]))
    response = xml.find(cloud.DAV + 'response')
    successful = response.find('./' + cloud.DAV + 'propstat/' + cloud.DAV + 'prop')
    creation = successful.find(cloud.DAV + 'creationdate')
    successful.remove(creation)
    missing = ET.SubElement(response, cloud.DAV + 'propstat')
    optional = ET.SubElement(missing, cloud.DAV + 'prop')
    ET.SubElement(optional, cloud.DAV + 'creationdate')
    ET.SubElement(optional, cloud.OC + 'checksums')
    ET.SubElement(missing, cloud.DAV + 'status').text = 'HTTP/1.1 404 Not Found'
    session.add('PROPFIND', dav(ROOT) + '/', [Response(ET.tostring(xml))], depth='1')
    assert client(session).list_files() == []


@pytest.mark.parametrize('uid', ['/other-user', 'other/user', '..', '%2e%2e', 'bad\\user', 42])
def test_canonical_user_identifier_cannot_inject_a_namespace(session, uid):
    body = json.dumps({'ocs': {'meta': {'statuscode': 100}, 'data': {'id': uid}}}).encode()
    session.add('GET', USER, [Response(body, status=200)])
    with pytest.raises(cloud.CloudSyncError, match='учётную запись'):
        client(session).list_files()
    assert len(session.calls) == 1


def test_absolute_same_origin_and_unicode_percent_names_are_supported(session):
    path = ROOT + '/Запас 100% готово.xlsx'
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': path, 'href': cloud.ORIGIN + dav(path)},
    ]))], depth='1')
    item, = client(session).list_files()
    assert item['path'] == path and item['name'] == 'Запас 100% готово.xlsx'


def test_depth_count_byte_and_time_limits_are_enforced(session, monkeypatch):
    monkeypatch.setattr(cloud, 'MAX_DEPTH', 0)
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': ROOT + '/child', 'dir': True},
    ]))], depth='1')
    with pytest.raises(cloud.CloudSyncError, match='глубину'):
        client(session).list_files()
    api = client(Session())
    api.deadline = 0
    with pytest.raises(cloud.CloudSyncError, match='истекло'):
        api.list_files()
    assert api.session.calls == []


def test_response_byte_limit_rejects_declared_and_streamed_oversize(session):
    session.add('GET', USER, [Response(b'{}', status=200, headers={'Content-Length': '65537'})])
    with pytest.raises(cloud.CloudSyncError, match='размер'):
        client(session).list_files()
    session.add('GET', USER, [Response(b' ' * 65537, status=200, headers={'Content-Length': '2'})])
    with pytest.raises(cloud.CloudSyncError, match='размер'):
        client(session).list_files()


def test_listing_entry_limit_never_returns_a_truncated_success(session, monkeypatch):
    monkeypatch.setattr(cloud, 'MAX_ENTRIES', 2)
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': ROOT + '/a.xlsx'}, {'path': ROOT + '/b.xlsx'},
    ]))], depth='1')
    with pytest.raises(cloud.CloudSyncError, match='слишком много'):
        client(session).list_files()


def test_directory_limit_stops_before_unbounded_recursive_requests(session, monkeypatch):
    monkeypatch.setattr(cloud, 'MAX_DIRECTORIES', 1)
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': ROOT + '/child', 'dir': True},
    ]))], depth='1')
    with pytest.raises(cloud.CloudSyncError, match='каталогов'):
        client(session).list_files()
    assert len(session.calls) == 4


def test_oversized_file_is_not_downloaded(session):
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': ROOT + '/huge.xlsx', 'size': cloud.MAX_FILE_BYTES + 1},
    ]))], depth='1')
    api = client(session)
    item, = api.list_files()
    before = len(session.calls)
    with pytest.raises(cloud.CloudSyncError, match='50 МБ'):
        api.download(item, lambda payload: None)
    assert len(session.calls) == before


def test_transient_network_errors_have_bounded_retries_without_secret_details(session):
    session.add('GET', USER, [requests.ConnectionError('https://secret:' + PASSWORD + '@example.test')] * 3)
    with pytest.raises(cloud.CloudSyncError, match='три попытки') as caught:
        client(session).list_files()
    assert len(session.calls) == 3 and PASSWORD not in str(caught.value)


def test_tls_errors_do_not_disable_verification_or_repeat_authentication(session):
    session.add('GET', USER, [requests.exceptions.SSLError('PRIVATE_CERTIFICATE_DETAILS')])
    with pytest.raises(cloud.CloudSyncError, match='сертификат') as caught:
        client(session).list_files()
    assert len(session.calls) == 1 and 'PRIVATE' not in str(caught.value)


def test_only_completed_listed_items_can_be_downloaded_and_bad_binary_is_rejected(session):
    data = workbook()
    path = ROOT + '/good.xlsx'
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': path, 'size': len(data)},
    ]))], depth='1')
    session.add('GET', dav(path), [Response(b'<html>not a workbook</html>', status=200)])
    api = client(session)
    with pytest.raises(cloud.CloudSyncError, match='подтверждённом'):
        api.download({'path': path}, lambda payload: None)
    assert session.calls == []
    item, = api.list_files()
    item['size'] = 1  # caller mutations cannot weaken expected size/checksum
    with pytest.raises(cloud.CloudSyncError, match='не полностью'):
        api.download(item, lambda payload: None)
    with pytest.raises(cloud.CloudSyncError, match='подтверждённом'):
        api.download({'path': '/Other/private.xlsx'}, lambda payload: None)


def test_etag_changed_during_download_does_not_overwrite_cached_version(session):
    path = ROOT + '/changing.xlsx'
    session.add('PROPFIND', dav(ROOT) + '/', [Response(response_xml([
        {'path': ROOT, 'dir': True}, {'path': path, 'size': 100},
    ]))], depth='1')
    session.add('GET', dav(path), [Response(status=412)])
    api = client(session)
    item, = api.list_files()
    with pytest.raises(cloud.CloudSyncError, match='изменился'):
        api.download(item, lambda payload: None)
