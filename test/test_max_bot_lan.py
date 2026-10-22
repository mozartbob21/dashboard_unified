"""Scoped LAN HTTP compatibility: synthetic accounts only, no outbound requests."""
from dataclasses import replace
from urllib.parse import urlsplit

import pytest
from fastapi.testclient import TestClient

from services.max_bot.config import Settings
from test.test_max_bot_link_api import api, COOKIE, CSRF


LAN = "http://10.113.6.20:8000"
PEER = ("10.113.6.74", 53000)


def enable_lan(api, *origins):
    api.config_box["value"] = replace(api.config, http_origins=tuple(origins or (LAN,)))


def client_for(api, origin=LAN, peer=PEER):
    # This project's Starlette TestClient cannot parse an IPv6 netloc itself.
    # An actual Host header still exercises the application's IPv6 URL handling.
    client = TestClient(api.app, base_url=LAN if origin.startswith("http://[") else origin,
                        client=peer, raise_server_exceptions=False)
    client.cookies.set("access_token", COOKIE)
    client.headers.update({"Host": urlsplit(origin).netloc, "Origin": origin, "X-CSRF-Token": CSRF})
    return client


def test_lan_http_is_opt_in_and_specific_origins_are_loaded():
    assert Settings.from_env({}).http_origins == ()
    assert Settings.from_env({"MAX_ZIP_BOT_HTTP_ORIGINS": " \t "}).http_origins == ()
    origins = (LAN, "http://192.168.1.20:8080", "http://[fd12:3456::20]:8000")
    settings = Settings.from_env({"MAX_ZIP_BOT_HTTP_ORIGINS": " , ".join(origins)})
    assert settings.http_origins == origins


@pytest.mark.parametrize("origin", [
    "https://10.113.6.20:8000", "ftp://10.113.6.20:8000", "//10.113.6.20:8000",
    "http://neurona.internal:8000", "http://localhost:8000", "http://127.0.0.1:8000",
    "http://8.8.8.8:8000", "http://192.0.2.20:8000", "http://198.51.100.20:8000",
    "http://203.0.113.20:8000", "http://169.254.10.20:8000", "http://100.64.1.20:8000",
    "http://0.0.0.0:8000", "http://255.255.255.255:8000", "http://224.0.0.1:8000",
    "http://[::]:8000", "http://[::1]:8000", "http://[fe80::20]:8000",
    "http://[2001:db8::20]:8000", "http://[::ffff:10.113.6.20]:8000",
    "http://[fd12:3456::20%eth0]:8000", "http://10.113.6.20:0",
    "http://10.113.6.20:65536", "http://10.113.6.20:wrong",
    "http://user:password@10.113.6.20:8000", "http://@10.113.6.20:8000",
    "http://10.113.6.20:8000/path", "http://10.113.6.20:8000?x=1",
    "http://10.113.6.20:8000#fragment", "http://10.113.6.20:8000\nHost: other",
    "http://10.113.6.20:8000\\@8.8.8.8", "http://012.113.6.20:8000",
])
def test_origin_configuration_rejects_ambiguous_or_non_private_targets(origin):
    with pytest.raises(ValueError):
        Settings.from_env({"MAX_ZIP_BOT_HTTP_ORIGINS": origin})


@pytest.mark.parametrize("path", ["/max-bot/api/status", "/max-bot/api/link", "/max-bot/api/confirm", "/max-bot/api/unlink"])
def test_default_http_api_is_still_denied_without_opening_store(api, path):
    with client_for(api) as client:
        response = client.get(path) if path.endswith("status") else client.post(path, json={})
    assert response.status_code == 403
    assert response.json()["detail"]["code"] == "MAX_BOT_HTTPS_REQUIRED"
    assert not api.config.db_path.exists()


def test_authorized_blocked_page_explains_transport_without_linking_secrets(api):
    with client_for(api) as client:
        response = client.get("/max-bot")
    assert response.status_code == 200
    assert "MAX" in response.text and "HTTPS" in response.text
    assert CSRF not in response.text
    assert "synthetic-token-only" not in response.text
    assert "no-store" in response.headers["cache-control"]
    assert not api.config.db_path.exists()


@pytest.mark.parametrize("origin", ["http://10.113.6.20:8001", "http://10.113.6.21:8000", "http://192.168.1.20:8000"])
def test_opt_in_does_not_open_other_private_hosts_or_ports(api, origin):
    enable_lan(api)
    with client_for(api, origin=origin) as client:
        response = client.post("/max-bot/api/link", json={})
    assert response.status_code == 403
    assert response.json()["detail"]["code"] == "MAX_BOT_HTTPS_REQUIRED"
    assert not api.config.db_path.exists()


@pytest.mark.parametrize("peer", [
    "8.8.8.8", "192.0.2.10", "198.51.100.10", "203.0.113.10", "169.254.1.10",
    "100.64.1.10", "0.0.0.0", "255.255.255.255", "224.0.0.1", "::",
    "fe80::10", "2001:db8::10", "::ffff:10.113.6.74", "fd12:3456::10%eth0", "::1%eth0", "client.invalid",
])
def test_private_host_does_not_authorize_public_or_unusable_peer(api, peer):
    enable_lan(api)
    with client_for(api, peer=(peer, 53000)) as client:
        response = client.post("/max-bot/api/link", json={})
    assert response.status_code == 403
    assert response.json()["detail"]["code"] == "MAX_BOT_HTTPS_REQUIRED"
    assert not api.config.db_path.exists()


@pytest.mark.parametrize("headers", [
    {"X-Forwarded-For": "10.113.6.74"},
    {"X-Forwarded-Proto": "https"},
    {"X-Forwarded-Host": "10.113.6.20:8000"},
    {"Forwarded": 'for="10.113.6.74";proto=https;host="10.113.6.20:8000"'},
    {"X-Forwarded-For": "10.113.6.74", "X-Forwarded-Proto": "https"},
])
def test_untrusted_forwarding_headers_cannot_grant_http_linking(api, headers):
    enable_lan(api)
    with client_for(api, peer=("8.8.8.8", 53000)) as client:
        response = client.post("/max-bot/api/link", json={}, headers=headers)
    assert response.status_code == 403
    assert response.json()["detail"]["code"] == "MAX_BOT_HTTPS_REQUIRED"
    assert not api.config.db_path.exists()


@pytest.mark.parametrize("origin,peer", [
    (LAN, PEER),
    ("http://172.20.10.20:8000", ("172.20.10.74", 53000)),
    ("http://192.168.1.20:8080", ("192.168.1.74", 53000)),
    ("http://[fd12:3456::20]:8000", ("fd12:3456::74", 53000)),
    (LAN, ("127.0.0.1", 53000)),
])
def test_exact_scoped_private_origin_can_use_authenticated_api(api, origin, peer):
    enable_lan(api, origin)
    with client_for(api, origin=origin, peer=peer) as client:
        page = client.get("/max-bot")
        assert page.status_code == 200 and CSRF in page.text
        response = client.post("/max-bot/api/link", json={})
    assert response.status_code == 200
    assert len(response.json()["code"]) == 12


@pytest.mark.parametrize("headers", [
    {"Origin": "http://10.113.6.20:8001"}, {"Origin": "http://other.invalid"},
    {"Origin": "null"}, {"Origin": ""}, {"X-CSRF-Token": ""}, {"X-CSRF-Token": "0" * 64},
])
def test_private_http_preserves_exact_origin_and_session_csrf(api, headers):
    enable_lan(api)
    with client_for(api) as client:
        for operation in ("link", "confirm", "unlink"):
            response = client.post("/max-bot/api/" + operation, json={}, headers=headers)
            assert response.status_code == 403
            assert response.json()["detail"]["code"] == "CSRF_FAILED"
    assert not api.config.db_path.exists()


@pytest.mark.parametrize("actor,status", [
    (None, 401), ({"id": 1, "modules": [], "is_active": True}, 403),
    ({"id": 1, "modules": ["zips"], "is_active": False}, 401),
    ({"id": 1, "modules": ["zips"], "kc_sub": "external"}, 403),
])
def test_private_http_preserves_session_and_current_module_access(api, actor, status):
    enable_lan(api)
    api.actor["user"] = actor
    with client_for(api) as client:
        assert client.get("/max-bot").status_code == status
        assert client.get("/max-bot/api/status").status_code == status
        for operation in ("link", "confirm", "unlink"):
            assert client.post("/max-bot/api/" + operation, json={}).status_code == status
    assert not api.config.db_path.exists()


def test_private_http_still_requires_platform_confirmation_and_supports_disconnect(api):
    enable_lan(api)
    with client_for(api) as client:
        code = client.post("/max-bot/api/link", json={}).json()["code"]
        request_id = api.store().redeem_link(code, 123456789, 987654321)["request_id"]
        assert client.get("/max-bot/api/status").json()["active"] is False
        result = client.post("/max-bot/api/confirm", json={"request_id": request_id})
        assert result.status_code == 200
        assert api.store().authorize(123456789, 987654321)["id"] == 1
        assert client.post("/max-bot/api/unlink", json={}).status_code == 200
        assert client.get("/max-bot/api/status").json()["active"] is False
