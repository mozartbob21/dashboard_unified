# test/test_api_ecur.py
"""Тесты API ЕЦУР. Внешние запросы мокируем на уровне Python-функций."""
from unittest.mock import patch, MagicMock
import sys

import pytest


@pytest.fixture
def ecur_idle():
    """Гарантирует, что ecur не 'running', и восстанавливает статус после."""
    from app import run_status
    old = dict(run_status["ecur"])
    run_status["ecur"]["running"] = False
    yield
    run_status["ecur"].update(old)


@pytest.fixture
def mock_ecur_client():
    """
    Создаёт мок-модуль services.ecur.client, чтобы избежать ImportError.
    Патчит sys.modules до импорта app.py эндпоинтов.
    """
    mock_module = MagicMock()
    mock_module.authenticate_user = MagicMock(return_value=(True, 42))
    mock_module.get_current_data = MagicMock(return_value={
        "is_authed": False,
        "rows": None,
        "meta": None,
    })
    mock_module.clear_session = MagicMock()
    mock_module.refresh_data = MagicMock(return_value=(True, 10))

    sys.modules["services.ecur.client"] = mock_module
    yield mock_module
    del sys.modules["services.ecur.client"]


def test_login_success(admin_client, ecur_idle, mock_ecur_client):
    mock_ecur_client.authenticate_user.return_value = (True, 42)

    resp = admin_client.post(
        "/ecur/api/login",
        json={"email": "user@example.com", "password": "secret"},
    )

    assert resp.status_code == 200
    data = resp.json()
    assert data["ok"] is True
    assert data["count"] == 42


def test_login_bad_credentials(admin_client, ecur_idle, mock_ecur_client):
    mock_ecur_client.authenticate_user.return_value = (False, "Неверный логин или пароль")

    resp = admin_client.post(
        "/ecur/api/login",
        json={"email": "user@example.com", "password": "wrong"},
    )

    assert resp.status_code == 400
    assert resp.json()["ok"] is False


def test_login_conflict_when_running(admin_client, ecur_idle, mock_ecur_client):
    from app import run_status
    run_status["ecur"]["running"] = True

    resp = admin_client.post(
        "/ecur/api/login",
        json={"email": "user@example.com", "password": "secret"},
    )

    assert resp.status_code == 409


def test_refresh_requires_session(admin_client, mock_ecur_client):
    mock_ecur_client.get_current_data.return_value = {
        "is_authed": False,
        "rows": None,
        "meta": None,
    }

    resp = admin_client.post("/ecur/api/refresh")

    assert resp.status_code == 401


def test_logout_resets_session(admin_client, mock_ecur_client):
    resp = admin_client.post("/ecur/api/logout")

    assert resp.status_code == 200
    assert resp.json()["ok"] is True
    mock_ecur_client.clear_session.assert_called_once()
    