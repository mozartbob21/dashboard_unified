"""Manager-only pilot access with real temporary account tables and no network."""
import ast
import asyncio
import sqlite3
from dataclasses import replace
from types import SimpleNamespace
from pathlib import Path

import pytest
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from services.max_bot import config
from services.max_bot.store import StoreError
from test.test_max_bot_link_api import api, post
from test.test_max_bot_store import store, fails, pending, linked, change, NOW, MAX_ID, CHAT_ID
from test.test_max_bot_handler import handler, link, event, button, callback


def test_existing_admin_role_and_client_flag_do_not_grant_pilot_access(api, monkeypatch):
    with sqlite3.connect(api.config.accounts_path) as conn:
        conn.execute('UPDATE account_control SET manager_user_id=NULL')
        conn.execute("UPDATE users SET role='Администратор' WHERE id=1")
    api.actor['user'].update(role='Администратор', can_manage_users=True)
    monkeypatch.setattr(config.Settings, 'from_env', classmethod(lambda cls, mapping=None: api.config))
    assert not config.available_for_user(api.actor['user'])
    assert api.client.get('/max-bot?tour=1').status_code == 403
    assert api.client.get('/max-bot/api/status').status_code == 403
    for operation in ('link', 'confirm', 'unlink'):
        response = post(api, operation)
        assert response.status_code == 403
        assert response.json()['detail']['code'] == 'MANAGER_REQUIRED'
    assert not api.config.db_path.exists()


def test_revocation_blocks_stale_web_session_before_any_link_code_is_created(api):
    assert api.client.get('/max-bot').status_code == 200
    with sqlite3.connect(api.config.accounts_path) as conn:
        conn.execute('UPDATE account_control SET manager_user_id=NULL')
    assert api.client.get('/max-bot').status_code == 403
    assert post(api, 'link').status_code == 403
    assert not api.config.db_path.exists()


@pytest.mark.parametrize('user_id', [1, 2])
@pytest.mark.parametrize('stage', ['create', 'redeem', 'confirm', 'authorize'])
def test_parent_and_delegated_permissions_are_rechecked_at_every_step(store, user_id, stage):
    code = store.create_link(user_id, now=NOW)['code']
    request_id = None
    if stage in ('confirm', 'authorize'):
        request_id = store.redeem_link(code, MAX_ID, CHAT_ID, now=NOW)['request_id']
    if stage == 'authorize':
        store.confirm_link(user_id, request_id, now=NOW)
    change(store, 'UPDATE account_control SET manager_user_id=NULL')
    change(store, 'DELETE FROM account_managers')
    operations = {
        'create': lambda: store.create_link(user_id, now=NOW),
        'redeem': lambda: store.redeem_link(code, MAX_ID, CHAT_ID, now=NOW),
        'confirm': lambda: store.confirm_link(user_id, request_id, now=NOW),
        'authorize': lambda: store.authorize(MAX_ID, CHAT_ID, now=NOW),
    }
    fails('MANAGER_REQUIRED', operations[stage])


def test_revoked_manager_cannot_use_existing_search_or_contact_buttons(handler, monkeypatch):
    link(handler)
    token = button(handler.handle(event('муфта D100')))
    with sqlite3.connect(handler.store.accounts_path) as conn:
        conn.execute('UPDATE account_control SET manager_user_id=NULL')
    def forbidden_read(*args, **kwargs):
        pytest.fail('Inventory and contacts must not be read after manager revocation')
    monkeypatch.setattr(handler.stock, 'search', forbidden_read)
    monkeypatch.setattr(handler.stock, 'contacts', forbidden_read)
    for request in (event('муфта'), callback(token)):
        response = handler.handle(request)['body']['text']
        assert 'Управление пользователями' in response
        assert 'Тестовый инженер' not in response
    # A revoked account can still explicitly disconnect through its private MAX chat.
    assert handler.handle(event('/stop')) is not None
    fails('LINK_REQUIRED', handler.store.authorize, 123456789, 321)


def test_missing_manager_schema_fails_closed_without_changing_accounts(api, monkeypatch):
    with sqlite3.connect(api.config.accounts_path) as conn:
        conn.execute('DROP TABLE account_managers')
    previous = api.config.accounts_path.read_bytes()
    monkeypatch.setattr(config.Settings, 'from_env', classmethod(lambda cls, mapping=None: api.config))
    assert not config.available_for_user(api.actor['user'])
    assert api.client.get('/max-bot').status_code == 503
    fails('SYSTEM_UNAVAILABLE', api.store().create_link, 1)
    assert api.config.accounts_path.read_bytes() == previous


def test_zips_page_uses_same_manager_gate_as_main_header(api, monkeypatch):
    root = Path(__file__).resolve().parents[1]
    tree = ast.parse((root/'app.py').read_text(encoding='utf-8-sig'))
    function = next(node for node in tree.body if isinstance(node, ast.AsyncFunctionDef) and node.name == 'zips_page')
    function.decorator_list = []
    module = ast.fix_missing_locations(ast.Module(body=[function], type_ignores=[]))
    namespace = {'Request': Request, 'HTMLResponse': object,
                 'templates': Jinja2Templates(directory=str(root/'templates')),
                 'get_user_from_token': lambda token: api.actor['user'],
                 'max_bot_available_for_user': config.available_for_user}
    exec(compile(module, str(root/'app.py'), 'exec'), namespace)
    monkeypatch.setattr(config.Settings, 'from_env', classmethod(lambda cls, mapping=None: api.config))
    def render():
        request = Request({'type':'http', 'method':'GET', 'scheme':'https', 'path':'/zips',
                           'query_string':b'', 'headers':[], 'server':('neurona.test',443),
                           'client':('127.0.0.1',53000), 'app':api.app, 'router':api.app.router})
        return asyncio.run(namespace['zips_page'](request)).body.decode()
    assert 'href="/max-bot"' in render()
    with sqlite3.connect(api.config.accounts_path) as conn:
        conn.execute('UPDATE account_control SET manager_user_id=NULL')
    assert 'href="/max-bot"' not in render()
