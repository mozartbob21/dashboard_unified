"""MAX transport/poller tests. Every network call uses httpx.MockTransport."""
import json
import sys
from types import SimpleNamespace

import httpx
import pytest

from services.max_bot import client, polling
from services.max_bot.store import Store, StoreError


TOKEN = "SYNTHETIC_MAX_TOKEN_NOT_REAL"


class Clock:
    def __init__(self):
        self.now = 0.0
        self.sleeps = []
    def __call__(self):
        return self.now
    def sleep(self, seconds):
        self.sleeps.append(seconds)
        self.now += seconds


def max_client(responder, clock=None):
    clock = clock or Clock()
    return client.MaxClient(TOKEN, transport=httpx.MockTransport(responder), clock=clock, sleeper=clock.sleep)


def test_fixed_origin_header_only_auth_and_safe_http_options(monkeypatch):
    original = httpx.Client
    options, calls = {}, []
    def factory(**kwargs):
        options.update(kwargs)
        return original(**kwargs)
    monkeypatch.setattr(client.httpx, "Client", factory)
    def responder(request):
        calls.append(request)
        return httpx.Response(200, json={"user_id": 55, "is_bot": True, "first_name": "not returned"})
    with max_client(responder) as api:
        assert api.get_me() == {"user_id": 55, "is_bot": True}
        assert TOKEN not in repr(api)
    assert options["verify"] is True
    assert options["trust_env"] is False and options["follow_redirects"] is False
    assert options["timeout"].read == 35
    assert str(calls[0].url) == "https://platform-api2.max.ru/me"
    assert calls[0].headers["Authorization"] == TOKEN
    assert calls[0].headers["Accept-Encoding"] == "identity"
    assert TOKEN not in str(calls[0].url)


@pytest.mark.parametrize("token", ["", None, "has space", "secret\r\nLeak: yes", "токен", "x" * 4097])
def test_invalid_token_never_makes_request(token):
    with pytest.raises(client.MaxAPIError) as error:
        client.MaxClient(token)
    assert error.value.code == "CONFIG_INVALID"
    assert token is None or not token or str(token) not in str(error.value)


def test_certificate_bundle_is_never_installed_or_verification_disabled(tmp_path):
    missing = tmp_path / "not-present.pem"
    with pytest.raises(client.MaxAPIError) as error:
        client.MaxClient(TOKEN, ca_bundle=missing)
    assert error.value.code == "TLS_CONFIG_INVALID"
    assert not missing.exists() and str(missing) not in str(error.value)


def test_redirects_never_forward_token_to_another_host():
    calls = []
    def responder(request):
        calls.append(request)
        return httpx.Response(302, headers={"location": "https://another.invalid/collect"})
    with max_client(responder) as api, pytest.raises(client.MaxAPIError) as error:
        api.get_me()
    assert error.value.code == "REDIRECT_DENIED" and len(calls) == 1
    assert calls[0].url.host == "platform-api2.max.ru"


@pytest.mark.parametrize(("method", "path"), [("POST", "/subscriptions"), ("DELETE", "/subscriptions"),
    ("PATCH", "/me"), ("GET", "https://another.invalid/steal"), ("GET", "/../private"), ("POST", "/chats/123/actions")])
def test_transport_cannot_mutate_bot_settings_or_be_used_as_generic_proxy(method, path):
    calls = []
    with max_client(lambda request: calls.append(request)) as api, pytest.raises(client.MaxAPIError):
        api._request(method, path)
    assert calls == []


@pytest.mark.parametrize(("status", "code", "retryable"), [(401, "MAX_AUTH_FAILED", False), (403, "MAX_AUTH_FAILED", False),
                                                          (429, "MAX_RATE_LIMIT", True), (500, "MAX_UNAVAILABLE", True),
                                                          (400, "MAX_REQUEST_REJECTED", False)])
def test_api_failure_never_exposes_response_or_token(status, code, retryable, caplog):
    secret = "PRIVATE_PHONE_AND_SEARCH_TEXT"
    with max_client(lambda _: httpx.Response(status, text=TOKEN + secret)) as api, pytest.raises(client.MaxAPIError) as error:
        api.get_me()
    assert error.value.code == code and error.value.retryable == retryable
    assert TOKEN not in str(error.value) + caplog.text and secret not in str(error.value) + caplog.text


@pytest.mark.parametrize("result", [{"is_bot": False, "user_id": 55}, {"is_bot": True, "user_id": "55"},
                                   {"is_bot": True, "user_id": True}, {"is_bot": True, "user_id": 0}, {}])
def test_get_me_requires_actual_bot_identity(result):
    with max_client(lambda _: httpx.Response(200, json=result)) as api, pytest.raises(client.MaxAPIError):
        api.get_me()


def test_no_webhook_mutation_when_subscription_exists():
    calls = []
    def responder(request):
        calls.append(request)
        return httpx.Response(200, json={"subscriptions": [{"url": "https://private.invalid/webhook"}]})
    with max_client(responder) as api, pytest.raises(client.MaxAPIError) as error:
        api.check_no_webhook()
    assert error.value.code == "WEBHOOK_ACTIVE"
    assert [(r.method, r.url.path) for r in calls] == [("GET", "/subscriptions")]
    assert "private.invalid" not in str(error.value)


def test_poll_requests_only_allowed_types_and_carries_persisted_marker():
    calls = []
    def responder(request):
        calls.append(request)
        return httpx.Response(200, json={"updates": [], "marker": 42})
    with max_client(responder) as api:
        assert api.get_updates(41) == {"updates": [], "marker": 42}
        api.get_updates()
    assert calls[0].url.params["types"] == "message_created,message_callback,bot_started,bot_stopped"
    assert calls[0].url.params["marker"] == "41"
    assert calls[0].url.params["timeout"] == "30" and calls[0].url.params["limit"] == "50"
    assert "marker" not in calls[1].url.params


@pytest.mark.parametrize("marker", [True, -1, "3", 2**63, {}])
def test_invalid_outgoing_marker_never_hits_network(marker):
    calls = []
    with max_client(lambda r: calls.append(r)) as api, pytest.raises(client.MaxAPIError):
        api.get_updates(marker)
    assert calls == []


@pytest.mark.parametrize("result", [{"updates": [], "marker": -1}, {"updates": [None], "marker": 2},
                                   {"updates": [{}], "marker": None}, {"updates": [], "marker": "3"},
                                   {"updates": [{}] * 51, "marker": 3}, {"updates": "bad", "marker": 3}])
def test_invalid_update_batch_is_not_acknowledged(result):
    with max_client(lambda _: httpx.Response(200, json=result)) as api, pytest.raises(client.MaxAPIError):
        api.get_updates()


def test_regressed_marker_stops_instead_of_rewinding():
    with max_client(lambda _: httpx.Response(200, json={"updates": [], "marker": 9})) as api, pytest.raises(client.MaxAPIError) as error:
        api.get_updates(10)
    assert error.value.code == "MAX_MARKER_REGRESSED" and not error.value.retryable


@pytest.mark.parametrize("raw", [b"not-json", b"[]", b'{"a":1,"a":2}', b'{"a":NaN}'])
def test_malformed_response_is_safe_error(raw):
    with max_client(lambda _: httpx.Response(200, content=raw)) as api, pytest.raises(client.MaxAPIError) as error:
        api.get_me()
    assert error.value.code == "MAX_RESPONSE_INVALID"


def test_declared_or_streamed_oversized_json_is_bounded(monkeypatch):
    monkeypatch.setattr(client, "MAX_RESPONSE_BYTES", 100)
    for headers in ({"content-length": "1000000"}, {}):
        with max_client(lambda _: httpx.Response(200, content=b"x" * 101, headers=headers)) as api, pytest.raises(client.MaxAPIError):
            api.get_me()


def test_send_and_callback_use_correct_private_endpoints_and_global_throttle():
    calls, clock = [], Clock()
    body = {"text": "Синтетический результат", "attachments": [{"type": "inline_keyboard", "payload": {
        "buttons": [[{"type": "callback", "text": "Подробнее", "payload": "opaque-id"}]]}}]}
    def responder(request):
        calls.append(request)
        return httpx.Response(200, json={"success": True} if request.url.path == "/answers" else {"message": {"body": {"mid": "m"}}})
    with max_client(responder, clock) as api:
        assert api.send(123, body)
        assert api.answer("callback_123", body)
        assert api.send(123, {"text": "Ещё сообщение"})
    assert [r.url.path for r in calls] == ["/messages", "/answers", "/messages"]
    assert all(r.url.params["disable_link_preview"] == "true" for r in calls)
    assert calls[0].url.params["user_id"] == "123" and "chat_id" not in calls[0].url.params
    assert calls[1].url.params["callback_id"] == "callback_123"
    assert json.loads(calls[1].content) == {"message": body}
    assert len(clock.sleeps) == 2 and all(d >= 0.55 for d in clock.sleeps)
    assert all("format" not in json.loads(r.content) for r in calls)


@pytest.mark.parametrize("body", [{"text": ""}, {"text": "x" * 4001}, {"text": "hi", "format": "html"},
                                 {"text": "hi", "attachments": [{"type": "share", "url": "https://other.invalid"}]},
                                 {"text": "hi", "notify": "false"}])
def test_outgoing_body_is_limited_and_cannot_inject_arbitrary_media_or_format(body):
    calls = []
    with max_client(lambda r: calls.append(r)) as api, pytest.raises(client.MaxAPIError):
        api.send(123, body)
    assert not calls


def test_outbound_timeout_is_uncertain_and_not_retried(caplog):
    calls = []
    def responder(request):
        calls.append(request)
        raise httpx.ReadTimeout("contains " + TOKEN + " PHONE", request=request)
    with max_client(responder) as api, pytest.raises(client.MaxAPIError) as error:
        api.send(123, {"text": "PRIVATE_SEARCH"})
    assert error.value.uncertain and not error.value.retryable and len(calls) == 1
    assert TOKEN not in str(error.value) + caplog.text and "PRIVATE_SEARCH" not in caplog.text


def test_read_timeout_is_retryable_but_sanitized():
    def responder(request):
        raise httpx.ReadTimeout(TOKEN, request=request)
    with max_client(responder) as api, pytest.raises(client.MaxAPIError) as error:
        api.get_updates()
    assert error.value.retryable and not error.value.uncertain


@pytest.fixture
def store(tmp_path):
    return Store(SimpleNamespace(db_path=tmp_path / "max-bot.sqlite", accounts_path=tmp_path / "accounts.sqlite",
                                 link_ttl=300, binding_ttl=2592000, audit_days=30))


class Handler:
    def __init__(self, callback=None):
        self.calls = []
        self.callback = callback
    @staticmethod
    def event_key(update):
        return update.get("event_id")
    def handle(self, update):
        self.calls.append(update["event_id"])
        if self.callback:
            return self.callback(update)
        return {"user_id": 123, "body": {"text": "Синтетический ответ"}}


class Endpoint:
    def __init__(self, batches=None, *, webhook=False, bot_id=55):
        self.batches = list(batches or [])
        self.webhook = webhook
        self.bot_id = bot_id
        self.calls = []
        self.send_timeout = False
        self.read_failures = 0
    def __call__(self, request):
        self.calls.append(request)
        path = request.url.path
        if path == "/subscriptions":
            return httpx.Response(200, json={"subscriptions": [{}] if self.webhook else []})
        if path == "/me":
            return httpx.Response(200, json={"is_bot": True, "user_id": self.bot_id})
        if path == "/updates":
            if self.read_failures:
                self.read_failures -= 1
                raise httpx.ReadTimeout("SYNTHETIC", request=request)
            return httpx.Response(200, json=self.batches.pop(0) if self.batches else {"updates": [], "marker": None})
        if self.send_timeout:
            raise httpx.ReadTimeout("SYNTHETIC", request=request)
        return httpx.Response(200, json={"success": True} if path == "/answers" else {"message": {"body": {"mid": "m"}}})


def event(key="event1", type="message_created"):
    return {"event_id": key, "update_type": type}


def test_poller_persists_marker_and_claims_before_handler_only_once(store):
    endpoint = Endpoint([{"updates": [event()], "marker": 10}, {"updates": [event()], "marker": 11}])
    handler, clock = Handler(), Clock()
    with max_client(endpoint, clock) as api:
        poller = polling.DevPoller(store, api, handler, clock=clock, sleeper=clock.sleep)
        assert poller.poll_once()["sent"] == 1
        assert store.get_marker() == 10
        assert poller.poll_once()["ignored"] == 1
        assert store.get_marker() == 11 and handler.calls == ["event1"]
        poller.close()
    requests = [r for r in endpoint.calls if r.url.path == "/updates"]
    assert "marker" not in requests[0].url.params and requests[1].url.params["marker"] == "10"
    assert sum(r.method == "POST" for r in endpoint.calls) == 1
    assert any(d >= 0.5 for d in clock.sleeps)


def test_poller_webhook_prevents_any_updates_and_preserves_other_lease(store):
    endpoint = Endpoint(webhook=True)
    with max_client(endpoint) as api:
        poller = polling.DevPoller(store, api, Handler(), owner="first")
        with pytest.raises(client.MaxAPIError) as error:
            poller.poll_once()
        assert error.value.code == "WEBHOOK_ACTIVE"
    assert [r.url.path for r in endpoint.calls] == ["/subscriptions"]
    assert store.lease("second")
    poller.close()  # A stale owner must not delete the second worker's lease.
    assert not store.lease("third")


def test_new_webhook_stops_already_running_poller_before_next_updates(store):
    endpoint = Endpoint([{"updates": [], "marker": 10}])
    with max_client(endpoint) as api:
        poller = polling.DevPoller(store, api, Handler())
        poller.poll_once()
        endpoint.webhook = True
        with pytest.raises(client.MaxAPIError) as error:
            poller.poll_once()
        assert error.value.code == "WEBHOOK_ACTIVE"
        poller.close()
    assert sum(r.url.path == "/updates" for r in endpoint.calls) == 1


def test_other_worker_lease_stops_before_network(store):
    assert store.lease("already-running", ttl=120)
    endpoint = Endpoint()
    with max_client(endpoint) as api, pytest.raises(polling.PollingError):
        polling.DevPoller(store, api, Handler(), owner="competitor").poll_once()
    assert endpoint.calls == []


def test_bot_identity_change_stops_before_cursor_or_bindings_reused(store):
    store.ensure_bot_id(99)
    store.set_marker(50)
    endpoint = Endpoint(bot_id=55)
    with max_client(endpoint) as api, pytest.raises(StoreError) as error:
        polling.DevPoller(store, api, Handler()).poll_once()
    assert error.value.code == "BOT_ID_CHANGED" and store.get_marker() == 50
    assert all(r.url.path != "/updates" for r in endpoint.calls)


def test_uncertain_send_is_not_replayed_after_batch_redelivery(store):
    endpoint = Endpoint([{"updates": [event()], "marker": 10}, {"updates": [event()], "marker": 11}])
    endpoint.send_timeout = True
    handler = Handler()
    with max_client(endpoint) as api:
        poller = polling.DevPoller(store, api, handler, sleeper=lambda _: None)
        assert poller.poll_once()["uncertain"] == 1
        assert poller.poll_once()["ignored"] == 1
        poller.close()
    assert handler.calls == ["event1"]
    assert sum(r.method == "POST" for r in endpoint.calls) == 1
    with store._connection() as connection:
        assert connection.execute("SELECT status FROM max_bot_events").fetchone()[0] == "outbound_uncertain"


def test_handler_failure_does_not_log_input_or_retry_side_effects(store, caplog):
    def fail(_):
        raise ValueError("PASSWORD_PHONE_PRIVATE_SEARCH")
    endpoint = Endpoint([{"updates": [event()], "marker": 10}])
    with max_client(endpoint) as api:
        poller = polling.DevPoller(store, api, Handler(fail))
        assert poller.poll_once()["failed"] == 1
        poller.close()
    assert "PASSWORD_PHONE_PRIVATE_SEARCH" not in caplog.text
    assert not any(r.method == "POST" for r in endpoint.calls)


def test_callback_updates_message_instead_of_posting_duplicate(store):
    endpoint = Endpoint([{"updates": [event(type="message_callback")], "marker": 10}])
    handler = Handler(lambda _: {"user_id": 123, "callback_id": "cb1", "body": {"text": "Синтетическое уточнение"}})
    with max_client(endpoint) as api:
        poller = polling.DevPoller(store, api, handler)
        assert poller.poll_once()["sent"] == 1
        poller.close()
    assert [r.url.path for r in endpoint.calls if r.method == "POST"] == ["/answers"]


def test_unsupported_event_is_ignored_without_handler_or_outbound(store):
    endpoint = Endpoint([{"updates": [event(type="user_added")], "marker": 10}])
    handler = Handler()
    with max_client(endpoint) as api:
        poller = polling.DevPoller(store, api, handler)
        assert poller.poll_once()["ignored"] == 1
        poller.close()
    assert not handler.calls and store.get_marker() == 10


def test_malformed_nested_event_does_not_block_next_update(store, caplog):
    class FragileHandler(Handler):
        @staticmethod
        def event_key(update):
            if update.get("event_id") == "malformed":
                raise ValueError("PRIVATE_RAW_UPDATE")
            return update.get("event_id")
    endpoint = Endpoint([{"updates": [event("malformed"), event("valid")], "marker": 10}])
    handler = FragileHandler()
    with max_client(endpoint) as api:
        poller = polling.DevPoller(store, api, handler)
        outcome = poller.poll_once()
        poller.close()
    assert outcome["ignored"] == 1 and outcome["sent"] == 1
    assert handler.calls == ["valid"] and store.get_marker() == 10
    assert "PRIVATE_RAW_UPDATE" not in caplog.text


def test_reconnect_has_bounded_backoff_and_releases_lease(store):
    endpoint, clock = Endpoint(), Clock()
    endpoint.read_failures = 6
    with max_client(endpoint, clock) as api:
        poller = polling.DevPoller(store, api, Handler(), clock=clock, sleeper=clock.sleep)
        assert poller.run(max_iterations=6) == 6
    assert clock.sleeps == [2, 4, 8, 16, 30, 30]
    assert store.lease("next-worker")


def test_cli_requires_explicit_action_and_never_echoes_accidental_secret(capsys):
    with pytest.raises(SystemExit) as error:
        polling.main([])
    assert error.value.code == 2
    with pytest.raises(SystemExit):
        polling.main(["--token", TOKEN])
    assert TOKEN not in capsys.readouterr().err


def test_offline_config_check_does_not_create_store_or_make_network(monkeypatch, tmp_path, capsys):
    seen = []
    def settings(env):
        seen.append(env)
        return SimpleNamespace(enabled=False, mode="development-polling", bot_token=env.get("MAX_ZIP_BOT_TOKEN", ""))
    monkeypatch.setitem(sys.modules, "services.max_bot.config", SimpleNamespace(Settings=SimpleNamespace(from_env=settings)))
    monkeypatch.setattr(client.httpx, "Client", lambda **_: pytest.fail("offline check made network client"))
    env_file = tmp_path / ".env"
    env_file.write_text("MAX_ZIP_BOT_TOKEN=" + TOKEN + "\n")
    monkeypatch.delenv("MAX_ZIP_BOT_TOKEN", raising=False)
    assert polling.main(["--check-config", "--env-file", str(env_file)]) == 0
    output = capsys.readouterr()
    assert TOKEN not in output.out + output.err and "Токен настроен" in output.out
    assert seen[0]["MAX_ZIP_BOT_TOKEN"] == TOKEN
    assert list(tmp_path.iterdir()) == [env_file]


def test_disabled_environment_cannot_start_network_even_with_run_flag(monkeypatch):
    settings = SimpleNamespace(enabled=False, mode="development-polling", bot_token=TOKEN)
    monkeypatch.setitem(sys.modules, "services.max_bot.config", SimpleNamespace(Settings=SimpleNamespace(from_env=lambda _: settings)))
    monkeypatch.setattr(client.httpx, "Client", lambda **_: pytest.fail("disabled bot made network client"))
    assert polling.main(["--run-development-only"]) == 2


def test_offline_enabled_but_missing_token_is_not_reported_ready(monkeypatch, capsys):
    class IncompleteSettings:
        enabled = True
        mode = "development-polling"
        bot_token = ""
        @staticmethod
        def validate_polling():
            raise ValueError("missing token")
    monkeypatch.setitem(sys.modules, "services.max_bot.config", SimpleNamespace(Settings=SimpleNamespace(from_env=lambda _: IncompleteSettings())))
    monkeypatch.setattr(client.httpx, "Client", lambda **_: pytest.fail("offline check made network client"))
    assert polling.main(["--check-config"]) == 2
    assert "Не удалось проверить конфигурацию" in capsys.readouterr().err


def test_env_process_values_take_precedence_without_interpolation(monkeypatch, tmp_path):
    seen = []
    def settings(env):
        seen.append(env)
        return SimpleNamespace(enabled=False, mode="development-polling", bot_token=env.get("MAX_ZIP_BOT_TOKEN", ""))
    monkeypatch.setitem(sys.modules, "services.max_bot.config", SimpleNamespace(Settings=SimpleNamespace(from_env=settings)))
    env_file = tmp_path / ".env"
    env_file.write_text("MAX_ZIP_BOT_TOKEN=FILE_TOKEN\nUNUSED_VALUE=${HOME}\n")
    monkeypatch.setenv("MAX_ZIP_BOT_TOKEN", "PROCESS_TOKEN")
    assert polling.main(["--check-config", "--env-file", str(env_file)]) == 0
    assert seen[0]["MAX_ZIP_BOT_TOKEN"] == "PROCESS_TOKEN"
    assert seen[0]["UNUSED_VALUE"] == "${HOME}"
