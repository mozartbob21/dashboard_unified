# test/test_api_mgkh_rm.py
"""Тесты API ЖКХ Redmine. Subprocess и фоновые запуски мокируем."""
from unittest.mock import patch


def test_run_check_requires_auth(api_client):
    resp = api_client.post("/mgkh-rm/run-check", follow_redirects=False)
    assert resp.status_code in (302, 401)


def test_run_check_starts_service(admin_client):
    with patch("app.start_background_service", return_value={"ok": True}) as mock_start:
        resp = admin_client.post("/mgkh-rm/run-check")

    assert resp.status_code == 200
    assert resp.json()["ok"] is True
    mock_start.assert_called_once()

    args, _ = mock_start.call_args
    assert args[0] == "mgkh_rm"


def test_run_check_blocks_when_already_running(admin_client):
    from fastapi.responses import JSONResponse

    busy = JSONResponse(status_code=400, content={"ok": False, "message": "уже выполняется"})
    with patch("app.start_background_service", return_value=busy):
        resp = admin_client.post("/mgkh-rm/run-check")

    assert resp.status_code == 400


def test_run_status_returns_state(admin_client):
    resp = admin_client.get("/mgkh-rm/run-status")

    assert resp.status_code == 200
    data = resp.json()
    assert "running" in data
    assert "stage" in data
    