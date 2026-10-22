"""Offline regression coverage for safe MAX connection diagnostics.

Only synthetic values and httpx.MockTransport are used. No .env, live MAX,
accounts database, stock snapshots or contacts are read by these tests.
"""
import socket
import ssl

import httpx
import pytest

from services.max_bot import client, polling


TOKEN = "SYNTHETIC_MAX_TOKEN_NOT_REAL"
PRIVATE = "SYNTHETIC_PRIVATE_QUERY_AND_PHONE"
DETAIL = f"https://example.invalid/private?token={TOKEN} {PRIVATE}"


def wrapped(cause, *, context=False):
    error = httpx.ConnectError(DETAIL)
    if context:
        error.__context__ = cause
    else:
        error.__cause__ = cause
    return error


def certificate_error():
    return ssl.SSLCertVerificationError(1, DETAIL)


@pytest.mark.parametrize(("make_error", "code", "retryable"), [
    (certificate_error, "MAX_TLS_CERTIFICATE_FAILED", False),
    (lambda: wrapped(certificate_error()), "MAX_TLS_CERTIFICATE_FAILED", False),
    (lambda: wrapped(wrapped(certificate_error(), context=True)), "MAX_TLS_CERTIFICATE_FAILED", False),
    (lambda: httpx.ConnectError("[SSL: CERTIFICATE_VERIFY_FAILED] " + DETAIL), "MAX_TLS_CERTIFICATE_FAILED", False),
    (lambda: socket.gaierror(socket.EAI_NONAME, DETAIL), "MAX_DNS_FAILED", True),
    (lambda: wrapped(socket.gaierror(socket.EAI_AGAIN, DETAIL)), "MAX_DNS_FAILED", True),
    (lambda: httpx.ConnectTimeout(DETAIL), "MAX_CONNECT_TIMEOUT", True),
    (lambda: httpx.ReadTimeout(DETAIL), "MAX_READ_TIMEOUT", True),
    (lambda: wrapped(ssl.SSLError(1, DETAIL)), "MAX_TLS_FAILED", True),
    (lambda: ssl.SSLError(1, DETAIL), "MAX_TLS_FAILED", True),
    (lambda: httpx.ConnectError(DETAIL), "MAX_CONNECTION_FAILED", True),
    (lambda: httpx.RemoteProtocolError(DETAIL), "MAX_CONNECTION_FAILED", True),
])
def test_connection_errors_have_distinct_safe_codes(make_error, code, retryable):
    error = client.classify_network_error(make_error())
    assert isinstance(error, client.MaxAPIError)
    assert error.code == code
    assert error.retryable is retryable
    assert error.uncertain is False
    assert error.message and str(error) == error.message
    assert TOKEN not in str(error) and PRIVATE not in str(error)
    assert "example.invalid" not in str(error)


@pytest.mark.parametrize("make_error", [
    certificate_error,
    lambda: wrapped(socket.gaierror(socket.EAI_NONAME, DETAIL)),
    lambda: httpx.ConnectTimeout(DETAIL),
    lambda: httpx.ReadTimeout(DETAIL),
    lambda: wrapped(ssl.SSLError(1, DETAIL)),
    lambda: httpx.ConnectError(DETAIL),
])
def test_network_failures_during_writes_are_never_automatically_retried(make_error):
    error = client.classify_network_error(make_error(), write=True)
    assert error.retryable is False
    # Preserve the existing conservative delivery policy: never guess whether
    # an outbound operation was accepted, or silently send it a second time.
    assert error.uncertain is True
    assert TOKEN not in str(error) and PRIVATE not in str(error)


def test_implicit_context_and_exception_arguments_preserve_ssl_cause():
    nested = httpx.ConnectError(certificate_error())
    outer = wrapped(nested, context=True)
    error = client.classify_network_error(outer)
    assert error.code == "MAX_TLS_CERTIFICATE_FAILED"
    assert error.retryable is False


def test_specific_certificate_failure_takes_priority_over_outer_tls_error():
    outer = ssl.SSLError(1, DETAIL)
    outer.__cause__ = certificate_error()
    error = client.classify_network_error(outer)
    assert error.code == "MAX_TLS_CERTIFICATE_FAILED"
    assert error.retryable is False


def test_cycles_in_exception_chains_do_not_hang_or_expose_values():
    first = httpx.ConnectError(DETAIL)
    second = RuntimeError(DETAIL)
    first.__cause__ = second
    second.__context__ = first
    error = client.classify_network_error(first)
    assert error.code == "MAX_CONNECTION_FAILED"
    assert TOKEN not in str(error) and PRIVATE not in str(error)


def test_read_transport_uses_classifier_and_suppresses_raw_network_details(caplog):
    calls = []
    def respond(request):
        calls.append(request)
        raise wrapped(certificate_error())
    with client.MaxClient(TOKEN, transport=httpx.MockTransport(respond)) as api:
        with pytest.raises(client.MaxAPIError) as caught:
            api.get_me()
    error = caught.value
    assert error.code == "MAX_TLS_CERTIFICATE_FAILED"
    assert error.retryable is False and error.uncertain is False
    assert error.__suppress_context__ is True
    assert len(calls) == 1
    assert TOKEN not in caplog.text + str(error)
    assert PRIVATE not in caplog.text + str(error)


@pytest.mark.parametrize("write", [False, True])
def test_read_timeout_keeps_retry_policy_and_delivery_uncertainty(write):
    calls = []
    def respond(request):
        calls.append(request)
        raise httpx.ReadTimeout(DETAIL, request=request)
    with client.MaxClient(TOKEN, transport=httpx.MockTransport(respond)) as api:
        with pytest.raises(client.MaxAPIError) as caught:
            if write:
                api.send(123, {"text": PRIVATE})
            else:
                api.get_updates()
    error = caught.value
    assert error.code == "MAX_READ_TIMEOUT"
    assert error.retryable is (not write)
    assert error.uncertain is write
    assert len(calls) == 1
    assert TOKEN not in str(error) and PRIVATE not in str(error)


class MemoryStore:
    """Only the lease calls reached before a failed startup are needed."""
    def __init__(self):
        self.released = []
    def lease(self, owner, ttl):
        return True
    def release_lease(self, owner):
        self.released.append(owner)
    def audit(self, event, outcome):
        return True


def test_certificate_failure_stops_polling_without_exponential_retries(caplog):
    calls, sleeps, store = [], [], MemoryStore()
    def respond(request):
        calls.append(request)
        raise wrapped(certificate_error())
    with client.MaxClient(TOKEN, transport=httpx.MockTransport(respond)) as api:
        worker = polling.DevPoller(store, api, object(), owner="synthetic-owner", sleeper=sleeps.append)
        with pytest.raises(client.MaxAPIError) as caught:
            worker.run(max_iterations=4)
    assert caught.value.code == "MAX_TLS_CERTIFICATE_FAILED"
    assert len(calls) == 1 and sleeps == []
    assert store.released
    assert "retry in" not in caplog.text
    assert TOKEN not in caplog.text and PRIVATE not in caplog.text


def test_dns_failure_retries_with_safe_code_and_existing_bounded_backoff(caplog):
    calls, sleeps, store = [], [], MemoryStore()
    def respond(request):
        calls.append(request)
        raise wrapped(socket.gaierror(socket.EAI_AGAIN, DETAIL))
    with client.MaxClient(TOKEN, transport=httpx.MockTransport(respond)) as api:
        worker = polling.DevPoller(store, api, object(), owner="synthetic-owner", sleeper=sleeps.append)
        assert worker.run(max_iterations=3) == 3
    assert len(calls) == 3 and sleeps == [2, 4, 8]
    assert "MAX_DNS_FAILED" in caplog.text
    assert TOKEN not in caplog.text and PRIVATE not in caplog.text
