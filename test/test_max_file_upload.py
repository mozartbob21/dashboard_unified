"""File delivery transport: synthetic files and mocked HTTPS, never live MAX."""
import io
import json
from pathlib import Path

import httpx
import pytest

from services.max_bot import client


BOT_TOKEN = "SYNTHETIC_BOT_TOKEN_NOT_REAL"
FILE_TOKEN = "SYNTHETIC_UPLOADED_FILE_TOKEN"
UPLOAD_URL = "https://fu.oneme.ru/api/upload.do?sig=SYNTHETIC_SIGNED_URL"


@pytest.fixture
def document(tmp_path):
    path = tmp_path / "Презентация.pptx"
    path.write_bytes(b"synthetic-presentation-content")
    return path


def api(responder):
    return client.MaxClient(BOT_TOKEN, transport=httpx.MockTransport(responder), sleeper=lambda _: None)


def test_upload_then_send_uses_private_recipient_and_signed_storage_without_credentials(document, monkeypatch):
    calls, options = [], []
    original = httpx.Client
    def factory(**kwargs):
        options.append(kwargs)
        return original(**kwargs)
    monkeypatch.setattr(client.httpx, "Client", factory)
    def responder(request):
        calls.append(request)
        if request.url.path == "/uploads":
            return httpx.Response(200, json={"url": UPLOAD_URL})
        if request.url.host == "fu.oneme.ru":
            return httpx.Response(200, json={"fileId": 101, "token": FILE_TOKEN})
        return httpx.Response(200, json={"message": {"body": {"mid": "synthetic"}}})
    with api(responder) as connection:
        # Even pre-existing cookies on the API client must not reach storage.
        connection._http.cookies.set("private", "synthetic-secret", domain="fu.oneme.ru")
        token = connection.upload_file(document)
        assert token == FILE_TOKEN
        assert connection.send_file(123, token, "Новая презентация") is True
    assert [r.url.host for r in calls] == ["platform-api2.max.ru", "fu.oneme.ru", "platform-api2.max.ru"]
    assert calls[0].method == "POST" and dict(calls[0].url.params) == {"type": "file"}
    assert calls[0].headers["Authorization"] == BOT_TOKEN
    assert "authorization" not in calls[1].headers and "cookie" not in calls[1].headers
    assert "multipart/form-data" in calls[1].headers["content-type"]
    assert b'name="data"' in calls[1].content and document.read_bytes() in calls[1].content
    assert str(document.parent).encode() not in calls[1].content
    assert calls[2].url.params["user_id"] == "123" and "chat_id" not in calls[2].url.params
    assert json.loads(calls[2].content) == {"text": "Новая презентация", "attachments": [{"type": "file", "payload": {"token": FILE_TOKEN}}]}
    assert len(options) == 2
    assert all(o["verify"] is True and o["trust_env"] is False and o["follow_redirects"] is False for o in options)
    assert "Authorization" not in options[1]["headers"]


@pytest.mark.parametrize("url", [
    "http://fu.oneme.ru/api/upload.do?sig=x", "https://fu.oneme.ru.evil.invalid/api/upload.do?sig=x",
    "https://evil.invalid/api/upload.do?sig=x", "https://127.0.0.1/api/upload.do?sig=x",
    "https://[::1]/api/upload.do?sig=x", "https://10.0.0.1/api/upload.do?sig=x",
    "https://user:secret@fu.oneme.ru/api/upload.do?sig=x", "https://fu.oneme.ru:8443/api/upload.do?sig=x",
    "https://fu.oneme.ru/api/upload.do?sig=x#fragment", "https://fu.oneme.ru/api/upload.do",
    "https://fu.oneme.ru/other?sig=x", "https://fu.oneme.ru/api/%75pload.do?sig=x",
    "https://fu.oneme.ru./api/upload.do?sig=x", "https://fu.oneme.ru\\@evil.invalid/api/upload.do?sig=x",
    "https://fu.oneme.ru/api/upload.do?sig=x\r\nX:secret", None, 123,
])
def test_untrusted_upload_url_never_receives_file_or_token(document, url):
    calls = []
    def responder(request):
        calls.append(request)
        return httpx.Response(200, json={"url": url})
    with api(responder) as connection, pytest.raises(client.MaxAPIError) as error:
        connection.upload_file(document)
    assert error.value.code == "MAX_UPLOAD_URL_DENIED" and not error.value.uncertain
    assert len(calls) == 1 and calls[0].url.path == "/uploads"
    assert "secret" not in str(error.value)


def test_upload_redirect_does_not_forward_file_or_authorization(document):
    calls = []
    def responder(request):
        calls.append(request)
        return httpx.Response(200, json={"url": UPLOAD_URL}) if len(calls) == 1 else httpx.Response(307, headers={"location": "https://evil.invalid/upload"})
    with api(responder) as connection, pytest.raises(client.MaxAPIError) as error:
        connection.upload_file(document)
    assert error.value.code == "MAX_UPLOAD_REDIRECT_DENIED" and len(calls) == 2
    assert not error.value.uncertain and not error.value.retryable


@pytest.mark.parametrize("size", [0, 51])
def test_size_check_precedes_any_network(document, size, monkeypatch):
    monkeypatch.setattr(client, "MAX_UPLOAD_BYTES", 50)
    document.write_bytes(b"x" * size)
    calls = []
    with api(lambda request: calls.append(request)) as connection, pytest.raises(client.MaxAPIError) as error:
        connection.upload_file(document)
    assert error.value.code == "MAX_FILE_SIZE_INVALID" and not calls


def test_missing_and_symlink_files_do_not_trigger_network(document, tmp_path):
    symlink = tmp_path / "link.pptx"
    symlink.symlink_to(document)
    for path in (tmp_path / "missing.pptx", symlink):
        calls = []
        with api(lambda request: calls.append(request)) as connection, pytest.raises(client.MaxAPIError) as error:
            connection.upload_file(path)
        assert error.value.code == "MAX_FILE_UNAVAILABLE" and not calls


def test_file_stream_is_chunked_without_reading_entire_file(document):
    document.write_bytes(b"x" * (65536 * 5 + 1))
    chunks = []
    class StreamingTransport(httpx.BaseTransport):
        def handle_request(self, request):
            if request.url.path == "/uploads":
                return httpx.Response(200, json={"url": UPLOAD_URL})
            chunks.extend(request.stream)
            return httpx.Response(200, json={"token": FILE_TOKEN})
    with client.MaxClient(BOT_TOKEN, transport=StreamingTransport()) as connection:
        assert connection.upload_file(document) == FILE_TOKEN
    assert max(map(len, chunks)) <= 65536
    assert sum(chunk == b"x" * 65536 for chunk in chunks) == 5


def test_bounded_reader_does_not_send_growth_and_detects_truncation():
    bounded = client._BoundedFile(io.BytesIO(b"original-and-later-growth"), 8)
    assert bounded.read(0) == b""
    assert bounded.read(1000) == b"original" and bounded.read(1000) == b""
    bounded = client._BoundedFile(io.BytesIO(b"short"), 8)
    assert bounded.read(1000) == b"short"
    with pytest.raises(OSError):
        bounded.read(1000)


def test_change_during_upload_is_not_returned_as_success(document):
    calls = []
    def responder(request):
        calls.append(request)
        if request.url.path == "/uploads":
            return httpx.Response(200, json={"url": UPLOAD_URL})
        document.write_bytes(document.read_bytes() + b"changed")
        return httpx.Response(200, json={"token": FILE_TOKEN})
    with api(responder) as connection, pytest.raises(client.MaxAPIError) as error:
        connection.upload_file(document)
    assert error.value.code == "MAX_FILE_UNAVAILABLE" and len(calls) == 2


def test_upload_reuses_configured_tls_context_without_relaxing_verification(document, monkeypatch):
    original = httpx.Client
    options = []
    def factory(**kwargs):
        options.append(kwargs)
        return original(**kwargs)
    monkeypatch.setattr(client.httpx, "Client", factory)
    pem = Path(__file__).resolve().parents[1] / "deploy/certs/russian-trusted-root-ca.pem"
    def responder(request):
        return httpx.Response(200, json={"url": UPLOAD_URL} if request.url.path == "/uploads" else {"token": FILE_TOKEN})
    with client.MaxClient(BOT_TOKEN, ca_bundle=pem, transport=httpx.MockTransport(responder)) as connection:
        connection.upload_file(document)
    assert options[0]["verify"] is options[1]["verify"]
    assert options[1]["verify"].check_hostname


@pytest.mark.parametrize("response", [
    lambda: httpx.Response(200, json={}),
    lambda: httpx.Response(200, json={"token": "has spaces"}),
    lambda: httpx.Response(200, content=b'{"token":"one","token":"two"}'),
    lambda: httpx.Response(200, content=b"PRIVATE_NOT_JSON"),
    lambda: httpx.Response(200, content=b'{"token":"x"}', headers={"content-length": "999999999"}),
])
def test_invalid_upload_response_is_sanitized_and_retryable_without_delivery(document, response, caplog):
    calls = []
    def responder(request):
        calls.append(request)
        return httpx.Response(200, json={"url": UPLOAD_URL}) if len(calls) == 1 else response()
    with api(responder) as connection, pytest.raises(client.MaxAPIError) as error:
        connection.upload_file(document)
    assert error.value.code == "MAX_UPLOAD_RESPONSE_INVALID" and error.value.retryable and not error.value.uncertain
    assert "PRIVATE" not in str(error.value) + caplog.text and BOT_TOKEN not in caplog.text
    assert len(calls) == 2


def test_upload_timeout_is_retryable_but_message_timeout_is_uncertain(document, caplog):
    calls = []
    def responder(request):
        calls.append(request)
        if request.url.path == "/uploads":
            return httpx.Response(200, json={"url": UPLOAD_URL})
        raise httpx.ReadTimeout("PRIVATE " + BOT_TOKEN, request=request)
    with api(responder) as connection:
        with pytest.raises(client.MaxAPIError) as error:
            connection.upload_file(document)
        assert error.value.retryable and not error.value.uncertain
        with pytest.raises(client.MaxAPIError) as error:
            connection.send_file(123, FILE_TOKEN, "Файл")
        assert error.value.uncertain and not error.value.retryable
    assert len(calls) == 3 and "PRIVATE" not in caplog.text and BOT_TOKEN not in caplog.text


@pytest.mark.parametrize("status", [400, 409, 422])
def test_only_explicit_attachment_not_ready_rejection_is_safe_to_retry(status):
    calls = []
    def responder(request):
        calls.append(request)
        return httpx.Response(status, json={"code": "attachment.not.ready", "message": "PRIVATE"})
    with api(responder) as connection, pytest.raises(client.MaxAPIError) as error:
        connection.send_file(123, FILE_TOKEN, "Файл")
    assert error.value.code == "MAX_ATTACHMENT_NOT_READY" and error.value.retryable and not error.value.uncertain
    assert len(calls) == 1 and "PRIVATE" not in str(error.value)


def test_not_ready_does_not_change_generic_text_send_behavior():
    with api(lambda _: httpx.Response(400, json={"code": "attachment.not.ready"})) as connection, pytest.raises(client.MaxAPIError) as error:
        connection.send(123, {"text": "Текст"})
    assert error.value.code == "MAX_REQUEST_REJECTED" and not error.value.retryable


@pytest.mark.parametrize("status", [500, 502, 503])
def test_server_error_on_send_is_always_uncertain_even_with_error_code(status):
    with api(lambda _: httpx.Response(status, json={"code": "attachment.not.ready"})) as connection, pytest.raises(client.MaxAPIError) as error:
        connection.send_file(123, FILE_TOKEN, "Файл")
    assert error.value.uncertain and not error.value.retryable


@pytest.mark.parametrize("payload", [
    {"url": "https://private.invalid"}, {"token": FILE_TOKEN, "url": "https://private.invalid"},
    {"token": None}, {"token": "x\r\nInjected:secret"}, {"token": "x" * 8193},
])
def test_file_payload_cannot_send_urls_or_invalid_tokens(payload):
    calls = []
    with api(lambda request: calls.append(request)) as connection, pytest.raises(client.MaxAPIError):
        connection.send(123, {"text": "Файл", "attachments": [{"type": "file", "payload": payload}]})
    assert not calls


def test_upload_allocation_cannot_be_used_for_other_media_or_body():
    calls = []
    with api(lambda request: calls.append(request)) as connection:
        for params, body in [({"type": "video"}, None), ({"type": "file"}, {"url": "https://private.invalid"}), (None, None)]:
            with pytest.raises(client.MaxAPIError):
                connection._request("POST", "/uploads", params=params, body=body)
    assert not calls
