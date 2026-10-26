from pathlib import Path
from urllib.parse import parse_qs

import httpx
import pytest

from services.max_bot.presentation_source import (
    Attachment, DOCUMENT_PATH, ORIGIN, RedmineSource, SourceError, parse_attachments,
)


PAGE = '''<div id="content"><h2>Еженедельные презентации</h2>
<div class="attachments"><p><a href="/attachments/download/123/report.pptx" class="icon icon-attachment">Доклад &amp; итоги.pptx</a>
— Еженедельная сводка <span class="size">(2 MB)</span>
<a href="/users/9">Автор</a><span class="author">19.10.2026 16:30</span></p></div></div>'''
LOGIN = '<form><input name="authenticity_token" value="test-csrf"><input name="password" type="password"></form>'


def source(handler, **kwargs):
    return RedmineSource("unit-user", "unit-password", transport=httpx.MockTransport(handler), **kwargs)


def test_parse_attachment_metadata_and_entities():
    items = parse_attachments(PAGE)
    assert items == [Attachment("123", "Доклад & итоги.pptx", "Еженедельная сводка", "19.10.2026 16:30", "/attachments/download/123/report.pptx")]


def test_parse_empty_explicit_attachment_list_and_duplicate_id():
    assert parse_attachments('<div class="attachments"></div>') == []
    assert len(parse_attachments(PAGE + PAGE)) == 1
    with pytest.raises(SourceError, match="противоречивые"):
        parse_attachments(PAGE + PAGE.replace("Доклад &amp; итоги", "Другое имя"))


@pytest.mark.parametrize("page", ["<html>Недоступно</html>", "", "<html><h2>Проекты</h2></html>"])
def test_unrecognized_page_is_not_treated_as_empty_baseline(page):
    with pytest.raises(SourceError) as caught:
        parse_attachments(page)
    assert caught.value.code == "SOURCE_FORMAT_CHANGED"


def test_form_login_keeps_cookie_and_returns_attachment_list():
    requests = []

    def handle(request):
        requests.append(request)
        if request.url.path == "/login" and request.method == "GET":
            return httpx.Response(200, text=LOGIN, headers={"set-cookie": "session=challenge; Secure; Path=/"})
        if request.url.path == "/login" and request.method == "POST":
            assert "session=challenge" in request.headers["cookie"]
            data = parse_qs(request.content.decode())
            assert data["username"] == ["unit-user"]
            assert data["password"] == ["unit-password"]
            assert data["authenticity_token"] == ["test-csrf"]
            assert data["back_url"] == [ORIGIN + DOCUMENT_PATH]
            return httpx.Response(303, headers={"location": DOCUMENT_PATH, "set-cookie": "session=authorized; Secure; Path=/"})
        if request.url.path == DOCUMENT_PATH and "session=authorized" in request.headers.get("cookie", ""):
            return httpx.Response(200, text=PAGE)
        return httpx.Response(302, headers={"location": "/login"})

    with source(handle) as client:
        assert client.list_attachments()[0].id == "123"
    assert len([r for r in requests if r.method == "POST"]) == 1
    assert all(r.url.host == "digital.rm.mosreg.ru" for r in requests)


def test_wrong_credentials_returns_sanitized_error():
    with source(lambda request: httpx.Response(200, text=LOGIN)) as client:
        with pytest.raises(SourceError) as caught:
            client.list_attachments()
    assert caught.value.code == "SOURCE_AUTH_FAILED"
    assert "unit-password" not in str(caught.value)
    assert "unit-user" not in str(caught.value)


@pytest.mark.parametrize("location", [
    "https://outside.invalid/steal", "//outside.invalid/path", "http://digital.rm.mosreg.ru/documents/16",
    "https://user:secret@digital.rm.mosreg.ru/documents/16", "https://digital.rm.mosreg.ru:444/anything",
])
def test_cross_origin_and_insecure_redirects_are_never_followed(location):
    requests = []

    def handle(request):
        requests.append(request)
        return httpx.Response(307, headers={"location": location})

    with source(handle) as client:
        with pytest.raises(SourceError) as caught:
            client.list_attachments()
    assert caught.value.code == "SOURCE_UNSAFE_URL"
    assert len(requests) == 1
    assert "secret" not in str(caught.value)


def test_login_post_cannot_redirect_credentials_to_external_origin():
    methods = []

    def handle(request):
        methods.append(request.method)
        if request.method == "POST":
            return httpx.Response(307, headers={"location": "https://outside.invalid/steal"})
        return httpx.Response(200, text=LOGIN)

    with source(handle) as client:
        with pytest.raises(SourceError) as caught:
            client.list_attachments()
    assert caught.value.code == "SOURCE_UNSAFE_URL"
    assert methods == ["GET", "GET", "POST"]


@pytest.mark.parametrize("href", [
    "https://outside.invalid/attachments/download/123/report.pptx",
    "/attachments/download/123/%2e%2e%2fsecret",
    "/attachments/download/123/%5csecret", "/attachments/download/123/report.pptx?key=secret",
    "/attachments/download/123/../secret", "/attachments/download/123/%00file",
])
def test_bad_attachment_link_is_rejected(href):
    with pytest.raises(SourceError) as caught:
        parse_attachments(PAGE.replace("/attachments/download/123/report.pptx", href))
    assert caught.value.code == "SOURCE_UNSAFE_URL"


def test_download_streams_to_private_file_with_safe_basename(tmp_path):
    attachment = Attachment("123", "../a\\b:report.pptx", path="/attachments/download/123/report.pptx")
    with source(lambda request: httpx.Response(200, content=b"PK\x03\x04presentation", headers={"content-type": "application/octet-stream"})) as client:
        target = client.download(attachment, tmp_path / "cache")
    assert target.parent == tmp_path / "cache"
    assert target.read_bytes() == b"PK\x03\x04presentation"
    assert target.name == "123-_a_b_report.pptx"
    assert target.stat().st_mode & 0o077 == 0
    assert not list(target.parent.glob("*.part"))


@pytest.mark.parametrize("login_redirect", [True, False])
def test_independent_download_logs_in_without_prior_discovery(tmp_path, login_redirect):
    requests = []

    def handle(request):
        requests.append(request)
        if request.url.path == "/login" and request.method == "POST":
            assert parse_qs(request.content.decode())["password"] == ["unit-password"]
            return httpx.Response(303, headers={"location": DOCUMENT_PATH, "set-cookie": "session=authorized; Secure; Path=/"})
        if request.url.path == "/login":
            return httpx.Response(200, text=LOGIN)
        if request.url.path == DOCUMENT_PATH:
            return httpx.Response(200, text=PAGE)
        if "session=authorized" in request.headers.get("cookie", ""):
            return httpx.Response(200, content=b"PKfile", headers={"content-type": "application/octet-stream"})
        if login_redirect:
            return httpx.Response(302, headers={"location": "/login"})
        return httpx.Response(200, text=LOGIN, headers={"content-type": "text/html"})

    with source(handle) as client:
        target = client.download(Attachment("123", "report.pptx"), tmp_path)
    assert target.read_bytes() == b"PKfile"
    assert len([request for request in requests if request.method == "POST"]) == 1


def test_download_does_not_loop_when_source_forgets_login(tmp_path):
    posts = []

    def handle(request):
        if request.url.path == "/login" and request.method == "POST":
            posts.append(request)
            return httpx.Response(303, headers={"location": DOCUMENT_PATH})
        if request.url.path == DOCUMENT_PATH:
            return httpx.Response(200, text=PAGE)
        if request.url.path == "/login":
            return httpx.Response(200, text=LOGIN)
        return httpx.Response(302, headers={"location": "/login"})

    with source(handle) as client:
        with pytest.raises(SourceError) as caught:
            client.download(Attachment("123", "report.pptx"), tmp_path)
    assert caught.value.code == "SOURCE_AUTH_FAILED"
    assert len(posts) == 1
    assert not list(tmp_path.iterdir())


def test_download_permission_rejection_does_not_retry_credentials(tmp_path):
    requests = []

    def handle(request):
        requests.append(request)
        return httpx.Response(403)

    with source(handle) as client:
        with pytest.raises(SourceError) as caught:
            client.download(Attachment("123", "report.pptx"), tmp_path)
    assert caught.value.code == "SOURCE_AUTH_FAILED"
    assert len(requests) == 1


def test_long_unicode_filename_keeps_extension_and_filesystem_byte_limit(tmp_path):
    with source(lambda request: httpx.Response(200, content=b"PKfile")) as client:
        target = client.download(Attachment("123", "😀" * 200 + ".pptx"), tmp_path)
    assert target.name.endswith(".pptx")
    assert len(target.name.encode("utf-8")) <= 184
    assert target.read_bytes() == b"PKfile"


@pytest.mark.parametrize("content_type, body", [
    ("text/html; charset=utf-8", b"Not a file"),
    ("application/xhtml+xml", b"Not a file"),
    ("application/octet-stream", b" \n<!DOCTYPE html><html>Login</html>"),
    ("application/octet-stream", b"\xef\xbb\xbf<html>Login</html>"),
    ("application/octet-stream", b""),
])
def test_html_and_empty_downloads_are_rejected_and_partial_files_removed(tmp_path, content_type, body):
    with source(lambda request: httpx.Response(200, content=body, headers={"content-type": content_type})) as client:
        with pytest.raises(SourceError) as caught:
            client.download(Attachment("123", "report.pptx"), tmp_path)
    assert caught.value.code == "SOURCE_INVALID_FILE"
    assert not list(tmp_path.iterdir())


class Chunks(httpx.SyncByteStream):
    def __iter__(self):
        yield b"12345"
        yield b"67890"


@pytest.mark.parametrize("headers", [{"content-length": "99999"}, {}])
def test_download_size_bound_with_and_without_content_length(tmp_path, headers):
    with source(lambda request: httpx.Response(200, stream=Chunks(), headers=headers), max_bytes=8) as client:
        with pytest.raises(SourceError) as caught:
            client.download(Attachment("123", "report.pptx"), tmp_path)
    assert caught.value.code == "SOURCE_TOO_LARGE"
    assert not list(tmp_path.iterdir())


def test_html_response_size_is_bounded(monkeypatch):
    import services.max_bot.presentation_source as module
    monkeypatch.setattr(module, "HTML_MAX_BYTES", 8)
    with source(lambda request: httpx.Response(200, stream=Chunks())) as client:
        with pytest.raises(SourceError) as caught:
            client.list_attachments()
    assert caught.value.code == "SOURCE_TOO_LARGE"


def test_mid_download_network_failure_is_safe_and_removes_partial_file(tmp_path):
    class Fails(httpx.SyncByteStream):
        def __iter__(self):
            yield b"a" * 65536
            raise httpx.ReadError("secret URL and password must not be exposed")

    with source(lambda request: httpx.Response(200, stream=Fails())) as client:
        with pytest.raises(SourceError) as caught:
            client.download(Attachment("123", "report.pptx"), tmp_path)
    assert caught.value.code == "SOURCE_CONNECTION_FAILED"
    assert "secret" not in str(caught.value)
    assert not list(tmp_path.iterdir())


def test_file_id_and_link_id_must_match_before_request(tmp_path):
    requests = []

    def handle(request):
        requests.append(request)
        return httpx.Response(200, content=b"file")

    with source(handle) as client:
        with pytest.raises(SourceError):
            client.download(Attachment("123", "report.pptx", path="/attachments/download/999/report.pptx"), tmp_path)
    assert requests == []


def test_download_login_redirect_and_symlink_cache_are_rejected(tmp_path):
    def handle(request):
        if request.url.path == "/login":
            return httpx.Response(200, text=LOGIN)
        return httpx.Response(302, headers={"location": "/login"})

    with source(handle) as client:
        with pytest.raises(SourceError) as caught:
            client.download(Attachment("123", "report.pptx"), tmp_path)
        assert caught.value.code == "SOURCE_AUTH_FAILED"
        link = tmp_path / "link"
        link.symlink_to(tmp_path, target_is_directory=True)
        with pytest.raises(SourceError) as caught:
            client.download(Attachment("123", "report.pptx"), link)
        assert caught.value.code == "SOURCE_CACHE_FAILED"


def test_client_ignores_environment_proxies_and_enables_tls_verification(monkeypatch):
    import services.max_bot.presentation_source as module
    captured = {}

    class Client:
        def __init__(self, **kwargs):
            captured.update(kwargs)

        def close(self):
            pass

    monkeypatch.setattr(module.httpx, "Client", Client)
    with RedmineSource("user", "password"):
        pass
    assert captured["verify"] is True
    assert captured["trust_env"] is False
    assert captured["follow_redirects"] is False


@pytest.mark.parametrize("status", [401, 403, 404, 429, 500, 503])
def test_http_failures_return_only_sanitized_diagnostics(status):
    with source(lambda request: httpx.Response(status, text="secret-source-details")) as client:
        with pytest.raises(SourceError) as caught:
            client.list_attachments()
    assert caught.value.code == ("SOURCE_AUTH_FAILED" if status in {401, 403} else "SOURCE_HTTP_FAILED")
    assert "secret-source-details" not in str(caught.value)


def test_redirect_loop_is_bounded():
    calls = []

    def handle(request):
        calls.append(request)
        return httpx.Response(302, headers={"location": DOCUMENT_PATH})

    with source(handle) as client:
        with pytest.raises(SourceError) as caught:
            client.list_attachments()
    assert caught.value.code == "SOURCE_REDIRECT_FAILED"
    assert len(calls) == 6


def test_missing_credentials_do_not_make_network_requests():
    with pytest.raises(SourceError) as caught:
        RedmineSource("", "")
    assert caught.value.code == "SOURCE_NOT_CONFIGURED"
