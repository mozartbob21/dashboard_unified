"""Authenticated linking E2E with temporary SQLite, synthetic sessions and no network."""
from dataclasses import replace
from datetime import datetime, timezone
import hashlib
import hmac
import re
import sqlite3
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.testclient import TestClient

from routers import max_bot
from services.max_bot.config import Settings
from services.max_bot.store import StoreError

ORIGIN = "https://neurona.test"
COOKIE = "synthetic-session-cookie"
CSRF = hmac.new(b"synthetic-test-only-secret", ("max-bot-csrf:" + COOKIE).encode(), hashlib.sha256).hexdigest()


@pytest.fixture
def api(tmp_path, monkeypatch):
    accounts = tmp_path / "accounts.sqlite"
    with sqlite3.connect(accounts) as conn:
        conn.execute("CREATE TABLE users(id INTEGER PRIMARY KEY,username TEXT,role TEXT,modules TEXT,is_active INTEGER,password_hash TEXT)")
        conn.execute("CREATE TABLE account_control(id INTEGER PRIMARY KEY, manager_user_id INTEGER)")
        conn.execute("INSERT INTO account_control VALUES(1,1)")
        conn.execute("CREATE TABLE account_managers(user_id INTEGER PRIMARY KEY)")
        conn.execute("INSERT INTO account_managers VALUES(2)")
        conn.executemany("INSERT INTO users VALUES (?,?,?,?,?,?)", [
            (1, "reader", "Пользователь", '["zips"]', 1, "synthetic-hash-one"),
            (2, "second", "Пользователь", '["zips"]', 1, "synthetic-hash-two"),
        ])
    config = Settings(enabled=True, mode="development-polling", bot_token="synthetic-token-only",
                      accounts_path=accounts, db_path=tmp_path / "bot.sqlite")
    actor = {"user": {"id": 1, "username": "reader", "role": "Пользователь", "modules": ["zips"], "is_active": True}}
    # Lazy auth imports cannot touch real user storage, secrets or .env.
    monkeypatch.setitem(sys.modules, "services.auth.security", SimpleNamespace(
        get_user_from_token=lambda _token: None,
        get_jwt_secret=lambda: "synthetic-test-only-secret",
    ))
    app = FastAPI()
    @app.middleware("http")
    async def synthetic_existing_session(request, call_next):
        request.state.user = actor["user"]
        return await call_next(request)
    app.include_router(max_bot.router)
    # Actual shared template is exercised when the new template extends it.
    repo = Path(__file__).resolve().parents[1]
    if not (repo / "templates/base_module.html").exists():
        import services
        repo = next(Path(path).parent for path in services.__path__ if (Path(path).parent / "templates/base_module.html").exists())
    app.mount("/static", StaticFiles(directory=str(repo / "static")), name="static")
    monkeypatch.setattr(max_bot, "templates", Jinja2Templates(directory=[str(max_bot.ROOT / "templates"), str(repo / "templates")]))
    current_config = {"value": config}
    app.dependency_overrides[max_bot.settings] = lambda: current_config["value"]
    max_bot._store.cache_clear()
    client = TestClient(app, base_url=ORIGIN, raise_server_exceptions=False)
    client.cookies.set("access_token", COOKIE)
    client.headers.update({"Origin": ORIGIN, "X-CSRF-Token": CSRF})
    yield SimpleNamespace(client=client, app=app, actor=actor, config=config, config_box=current_config,
                          store=lambda: max_bot._store(config))
    client.close()
    max_bot._store.cache_clear()


def post(api, operation, payload=None, **kwargs):
    return api.client.post("/max-bot/api/" + operation, json={} if payload is None else payload, **kwargs)


def pending(api):
    response = post(api, "link")
    assert response.status_code == 200
    return api.store().redeem_link(response.json()["code"], 123456789, 987654321)["request_id"]


def test_full_linking_journey_is_two_step_private_and_reversible(api):
    page = api.client.get("/max-bot")
    assert page.status_code == 200
    assert f'name="max-bot-csrf" content="{CSRF}"' in page.text
    assert "synthetic-token-only" not in page.text
    assert api.client.get("/max-bot/api/status").json()["active"] is False
    request_id = pending(api)
    with pytest.raises(StoreError):
        api.store().authorize(123456789, 987654321)
    response = api.client.get("/max-bot/api/status")
    assert response.status_code == 200
    state = response.json()
    assert state["pending"][0]["id"] == request_id
    assert state["pending"][0]["label"] == "MAX •••• 6789"
    assert "123456789" not in response.text and "987654321" not in response.text
    assert datetime.fromisoformat(state["pending"][0]["expires_at"]).tzinfo is not None
    assert post(api, "confirm", {"request_id": request_id}).status_code == 200
    assert api.store().authorize(123456789, 987654321)["id"] == 1
    assert api.client.get("/max-bot/api/status").json()["active"] is True
    assert post(api, "unlink").status_code == 200
    assert api.client.get("/max-bot/api/status").json()["active"] is False
    with pytest.raises(StoreError):
        api.store().authorize(123456789, 987654321)
    for reply in (page, response):
        assert "no-store" in reply.headers["cache-control"]
        assert reply.headers["referrer-policy"] == "no-referrer"


@pytest.mark.parametrize("actor,status", [
    (None, 401),
    ({"id": 1, "modules": [], "is_active": True, "role": "Администратор"}, 403),
    ({"id": 1, "modules": ["zips"], "is_active": False}, 401),
    ({"id": 1, "modules": ["zips"], "kc_sub": "external-identity"}, 403),
    ({"id": "1", "modules": ["zips"]}, 403),
    ({"id": True, "modules": ["zips"]}, 403),
])
def test_every_route_rejects_missing_revoked_or_nonlocal_identity(api, actor, status):
    api.actor["user"] = actor
    assert api.client.get("/max-bot").status_code == status
    assert api.client.get("/max-bot/api/status").status_code == status
    for operation in ("link", "confirm", "unlink"):
        assert post(api, operation).status_code == status
    assert not api.config.db_path.exists()


@pytest.mark.parametrize("operation,payload", [("link", {"user_id": 2}), ("unlink", {"user_id": 2}),
    ("confirm", {"user_id": 2, "request_id": "a" * 32})])
def test_client_cannot_select_account_id(api, operation, payload):
    assert post(api, operation, payload).status_code == 400


def test_pending_request_is_owned_by_authorized_account(api):
    request_id = pending(api)
    api.actor["user"] = {**api.actor["user"], "id": 2, "username": "second"}
    assert api.client.get("/max-bot/api/status").json()["pending"] == []
    assert post(api, "confirm", {"request_id": request_id}).status_code == 400
    api.actor["user"] = {**api.actor["user"], "id": 1, "username": "reader"}
    assert post(api, "confirm", {"request_id": request_id}).status_code == 200


@pytest.mark.parametrize("headers", [
    {"Origin": "https://other.invalid"}, {"Origin": "null"}, {"Origin": ""},
    {"X-CSRF-Token": ""}, {"X-CSRF-Token": "0" * 64},
])
def test_csrf_and_origin_are_required(api, headers):
    for operation in ("link", "confirm", "unlink"):
        assert post(api, operation, headers=headers).status_code == 403


def test_non_ascii_csrf_is_rejected_as_403_not_server_error(api):
    response = api.client.post("/max-bot/api/link", content="{}", headers={
        "Content-Type": "application/json", "X-CSRF-Token": b"\xff" * 64,
    })
    assert response.status_code == 403


def test_csrf_token_from_another_session_is_invalid(api):
    api.client.cookies.clear()
    api.client.cookies.set("access_token", "different-synthetic-cookie")
    assert post(api, "link").status_code == 403


@pytest.mark.parametrize("path", ["/max-bot/api/status", "/max-bot/api/link"])
def test_lan_http_cannot_reveal_or_exchange_linking_data(api, path):
    with TestClient(api.app, base_url="http://192.168.1.20:8000", client=("192.168.1.10", 53000)) as client:
        client.cookies.set("access_token", COOKIE)
        headers = {"Origin": "http://192.168.1.20:8000", "X-CSRF-Token": CSRF}
        response = client.post(path, json={}, headers=headers) if path.endswith("link") else client.get(path)
    assert response.status_code == 403
    assert "code" not in response.json().get("detail", {}) or response.json()["detail"]["code"] == "MAX_BOT_HTTPS_REQUIRED"


def test_loopback_http_is_allowed_but_host_header_does_not_make_remote_peer_local(api):
    with TestClient(api.app, base_url="http://127.0.0.1:8000", client=("127.0.0.1", 53000)) as local:
        local.cookies.set("access_token", COOKIE)
        response = local.post("/max-bot/api/link", json={}, headers={"Origin": "http://127.0.0.1:8000", "X-CSRF-Token": CSRF})
        assert response.status_code == 200
    with TestClient(api.app, base_url="http://127.0.0.1:8000", client=("192.168.1.10", 53000)) as remote:
        remote.cookies.set("access_token", COOKIE)
        assert remote.get("/max-bot/api/status").status_code == 403


@pytest.mark.parametrize("body,status", [("[]", 400), ("{broken", 400), ("x" * 1025, 413),
    ('{"request_id":"' + "a" * 32 + '","request_id":"' + "b" * 32 + '"}', 400)])
def test_bounded_strict_request_json(api, body, status):
    reply = api.client.post("/max-bot/api/confirm", content=body, headers={"Content-Type": "application/json"})
    assert reply.status_code == status


def test_form_content_and_unknown_business_routes_are_not_exposed(api):
    assert api.client.post("/max-bot/api/link", content="user_id=1").status_code == 415
    for path in ("/max-bot/api/accounts", "/max-bot/api/contacts", "/max-bot/api/scan", "/zip_curator/api/state", "/data/dashboard.db"):
        assert api.client.get(path).status_code == 404


def test_disabled_configuration_does_not_create_security_store(api):
    api.config_box["value"] = replace(api.config, enabled=False, bot_token="")
    assert api.client.get("/max-bot/api/status").json() == {"enabled": False, "active": False, "pending": []}
    assert post(api, "link").status_code == 503
    assert not api.config.db_path.exists()


@pytest.mark.parametrize("change", ["UPDATE users SET modules='[]' WHERE id=1", "UPDATE users SET is_active=0 WHERE id=1"])
def test_store_rechecks_fresh_account_even_if_request_state_is_stale(api, change):
    with sqlite3.connect(api.config.accounts_path) as conn:
        conn.execute(change)
    assert api.client.get("/max-bot/api/status").status_code == 403
    assert post(api, "link").status_code == 403


def test_code_is_single_response_not_repeated_in_status_or_technical_output(api, capsys):
    result = post(api, "link")
    assert result.status_code == 200
    code = result.json()["code"]
    assert re.fullmatch(r"[0-9A-HJKMNP-TV-Z]{12}", code)
    assert datetime.fromisoformat(result.json()["expires_at"]) > datetime.now(timezone.utc)
    assert code not in api.client.get("/max-bot/api/status").text
    logs = capsys.readouterr()
    assert code not in logs.out + logs.err
