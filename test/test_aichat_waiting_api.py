"""Offline contract tests for the real waiting-facts route without app startup."""
import ast
import asyncio
from copy import deepcopy
import os
from pathlib import Path
from unittest.mock import Mock

from fastapi import FastAPI, HTTPException, Request
from fastapi.testclient import TestClient
import pytest

from services.aichat import waiting_facts

PROJECT_ROOT = Path(os.environ.get('AICHAT_PROJECT_ROOT', Path(__file__).resolve().parents[1]))


@pytest.fixture
def api(monkeypatch):
    public = {
        'items': [
            {'topic': 'ИИ', 'text': 'Языковые модели находят закономерности в текстовых данных.'},
            {'topic': 'ЖКХ', 'text': 'Теплоизоляция помогает уменьшать потери тепла в зданиях.'},
        ],
        'generated_at': '2026-10-13T10:00:00+00:00',
        'preparing': False,
    }
    get_facts = Mock(side_effect=lambda **kwargs: deepcopy(public))
    monkeypatch.setattr(waiting_facts, 'get_facts', get_facts)
    # A test regression must not silently reach a real provider or cache worker.
    monkeypatch.setattr(waiting_facts, '_generate_items', Mock(side_effect=AssertionError('No model calls in API tests')))
    monkeypatch.setattr('socket.create_connection', Mock(side_effect=AssertionError('No network in API tests')))
    monkeypatch.setattr('socket.socket.connect', Mock(side_effect=AssertionError('No network in API tests')))
    identity = {'user': {'id': 7, 'username': 'private-account', 'email': 'private@example.invalid'}}
    app = FastAPI()

    @app.middleware('http')
    async def controlled_identity(request, call_next):
        if 'user' in identity:
            request.state.user = identity['user']
        return await call_next(request)

    source = (PROJECT_ROOT / 'app.py').read_text(encoding='utf-8-sig')
    route = next(node for node in ast.parse(source).body
                 if isinstance(node, ast.AsyncFunctionDef) and node.name == 'aichat_waiting_facts')
    # Keep the actual decorator, method, path, annotations and body.
    namespace = {'app': app, 'Request': Request, 'HTTPException': HTTPException, 'asyncio': asyncio}
    exec(compile(ast.Module(body=[route], type_ignores=[]), str(PROJECT_ROOT / 'app.py'), 'exec'), namespace)
    with TestClient(app) as client:
        yield client, identity, get_facts, public


@pytest.mark.parametrize('identity_value', [None, {}])
def test_missing_authentication_cannot_read_or_warm_facts(api, identity_value):
    client, identity, get_facts, _ = api
    identity['user'] = identity_value
    result = client.get('/aichat/api/waiting-facts')
    assert result.status_code == 401
    assert set(result.json()) == {'detail'}
    get_facts.assert_not_called()
    identity.pop('user')
    assert client.get('/aichat/api/waiting-facts?warm=false').status_code == 401
    get_facts.assert_not_called()


def test_authenticated_get_has_only_public_facts_and_is_not_cached(api):
    client, _, get_facts, public = api
    result = client.get('/aichat/api/waiting-facts?question=PRIVATE_QUESTION&dialog_id=PRIVATE_DIALOG')
    assert result.status_code == 200
    assert result.headers['cache-control'] == 'private, no-store'
    assert result.headers['content-type'] == 'application/json'
    assert result.json() == public
    assert set(result.json()) == {'items', 'generated_at', 'preparing'}
    assert all(set(item) == {'topic', 'text'} for item in result.json()['items'])
    get_facts.assert_called_once_with(warm=True)
    for private in ['private-account', 'private@example.invalid', 'PRIVATE_QUESTION', 'PRIVATE_DIALOG']:
        assert private not in result.text


def test_request_time_read_forwards_false_without_activating_generation(api):
    client, _, get_facts, public = api
    result = client.get('/aichat/api/waiting-facts?warm=false')
    assert result.status_code == 200
    assert result.json() == public
    assert result.headers['cache-control'] == 'private, no-store'
    get_facts.assert_called_once_with(warm=False)


def test_invalid_warm_flag_is_rejected_before_service_call(api):
    client, _, get_facts, _ = api
    result = client.get('/aichat/api/waiting-facts?warm=not-a-boolean')
    assert result.status_code == 422
    assert result.json()['detail'][0]['loc'] == ['query', 'warm']
    get_facts.assert_not_called()
