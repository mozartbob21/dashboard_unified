"""Home navigation and tour entry use the real local-account eligibility gate.

Only the home function is compiled from app.py: application startup, live DBs,
scheduled jobs, secrets and network integrations are never imported or started.
"""
import ast
import asyncio
from html.parser import HTMLParser
from types import SimpleNamespace

import pytest
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from services.max_bot import config
from test.test_max_bot_link_api import api


LOCAL_READER = {"id": 1, "username": "synthetic-reader", "role": "Пользователь",
                "modules": ["zips"], "is_active": True}


def configure(monkeypatch, *, enabled="true", mode="development-polling"):
    # Use a complete synthetic environment, never real .env values or tokens.
    monkeypatch.setattr(config, "os", SimpleNamespace(environ={
        "MAX_ZIP_BOT_ENABLED": enabled, "MAX_ZIP_BOT_MODE": mode,
        "MAX_ZIP_BOT_TOKEN": "synthetic-navigation-token",
    }))


@pytest.mark.parametrize("user, expected", [
    (LOCAL_READER, True),
    ({**LOCAL_READER, "role": "Администратор"}, True),
    ({**LOCAL_READER, "modules": []}, False),
    ({**LOCAL_READER, "role": "Администратор", "username": "admin", "modules": []}, False),
    ({**LOCAL_READER, "modules": ["mingkh", "summarizer"]}, False),
    ({**LOCAL_READER, "is_active": False}, False),
    ({**LOCAL_READER, "id": True}, False),
    ({**LOCAL_READER, "id": "1"}, False),
    ({**LOCAL_READER, "id": None}, False),
    ({**LOCAL_READER, "kc_sub": "synthetic-keycloak-user"}, False),
    ({**LOCAL_READER, "role": "Администратор", "kc_sub": "synthetic-keycloak-admin"}, False),
    ({}, False),
    (None, False),
])
def test_entry_eligibility_uses_current_account_and_explicit_zips_grant(monkeypatch, user, expected):
    configure(monkeypatch)
    assert config.available_for_user(user) is expected


@pytest.mark.parametrize("enabled, mode", [
    ("false", "development-polling"),
    ("false", "disabled"),
    ("true", "invalid-mode"),
    ("true", "production"),
])
def test_disabled_or_malformed_optional_bot_configuration_hides_entry(monkeypatch, enabled, mode):
    configure(monkeypatch, enabled=enabled, mode=mode)
    assert config.available_for_user(LOCAL_READER) is False


class Links(HTMLParser):
    def __init__(self):
        super().__init__()
        self.hrefs = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            self.hrefs.append(dict(attrs).get("href"))


@pytest.fixture
def render_actual_home():
    repository = config.ROOT
    source = repository / "app.py"
    tree = ast.parse(source.read_text(encoding="utf-8-sig"), filename=str(source))
    function = next(node for node in tree.body
                    if isinstance(node, ast.AsyncFunctionDef) and node.name == "home")
    function.decorator_list = []
    module = ast.fix_missing_locations(ast.Module(body=[function], type_ignores=[]))

    application = FastAPI()
    application.mount("/static", StaticFiles(directory=str(repository / "static")), name="static")
    templates = Jinja2Templates(directory=str(repository / "templates"))
    namespace = {
        "Request": Request,
        "templates": templates,
        "effective_modules": lambda user: user.get("modules", []),
        "home_preferences": SimpleNamespace(preferences=lambda user: {}),
        "is_full_access": lambda user: user.get("role") == "Администратор",
        "is_account_manager": lambda user: user.get("role") == "Администратор",
        # Exercise the actual production eligibility function, not a constant.
        "max_bot_available_for_user": config.available_for_user,
    }
    exec(compile(module, str(source), "exec"), namespace)

    def render(user):
        request = Request({
            "type": "http", "method": "GET", "scheme": "https", "path": "/",
            "raw_path": b"/", "query_string": b"", "headers": [],
            "server": ("neurona.test", 443), "client": ("127.0.0.1", 53000),
            "root_path": "", "app": application, "router": application.router,
        })
        request.state.user = user
        response = asyncio.run(namespace["home"](request))
        parser = Links()
        parser.feed(response.body.decode("utf-8"))
        return response, parser.hrefs

    return render


@pytest.mark.parametrize("user, enabled, mode, expected", [
    (LOCAL_READER, "true", "development-polling", True),
    ({**LOCAL_READER, "role": "Администратор"}, "true", "development-polling", True),
    ({**LOCAL_READER, "modules": []}, "true", "development-polling", False),
    ({**LOCAL_READER, "role": "Администратор", "username": "admin", "modules": []}, "true", "development-polling", False),
    ({**LOCAL_READER, "is_active": False}, "true", "development-polling", False),
    ({**LOCAL_READER, "id": "1"}, "true", "development-polling", False),
    ({**LOCAL_READER, "id": True}, "true", "development-polling", False),
    ({**LOCAL_READER, "kc_sub": "synthetic-keycloak-user"}, "true", "development-polling", False),
    (LOCAL_READER, "false", "development-polling", False),
    (LOCAL_READER, "true", "invalid-mode", False),
])
def test_actual_home_and_template_show_both_links_only_when_eligible(
        monkeypatch, render_actual_home, user, enabled, mode, expected):
    configure(monkeypatch, enabled=enabled, mode=mode)
    response, links = render_actual_home(user)
    assert response.status_code == 200
    assert response.context["max_bot_available"] is expected
    assert ("/max-bot" in links) is expected
    assert ("/max-bot?tour=1" in links) is expected
    assert "synthetic-navigation-token" not in response.body.decode("utf-8")


@pytest.mark.parametrize("actor, status", [
    (None, 401),
    ({**LOCAL_READER, "modules": []}, 403),
    ({**LOCAL_READER, "role": "Администратор", "modules": []}, 403),
    ({**LOCAL_READER, "is_active": False}, 401),
    ({**LOCAL_READER, "id": True}, 403),
    ({**LOCAL_READER, "id": "1"}, 403),
    ({**LOCAL_READER, "kc_sub": "synthetic-keycloak-user"}, 403),
])
def test_tour_query_cannot_bypass_server_authorization_or_create_link_data(api, actor, status):
    api.actor["user"] = actor
    response = api.client.get("/max-bot?tour=1")
    assert response.status_code == status
    assert not api.config.db_path.exists()
    assert "synthetic-token-only" not in response.text


def test_authorized_tour_is_read_only_and_does_not_create_link_codes(api):
    response = api.client.get("/max-bot?tour=1")
    assert response.status_code == 200
    assert not api.config.db_path.exists()
    assert "synthetic-token-only" not in response.text
    assert "no-store" in response.headers["cache-control"]
