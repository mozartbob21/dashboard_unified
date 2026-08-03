# test/conftest.py
import pytest
import sqlite3
import sys
from pathlib import Path
from unittest.mock import patch

# Корень проекта в sys.path, чтобы работали импорты app и services
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

SCHEMA = """
CREATE TABLE IF NOT EXISTS scheduler_jobs (
    module_id TEXT PRIMARY KEY,
    module_label TEXT,
    module_icon TEXT,
    command TEXT,
    enabled INTEGER DEFAULT 0,
    interval_minutes INTEGER,
    last_run_at TEXT,
    next_run_at TEXT,
    last_status TEXT,
    last_error TEXT,
    created_at TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS run_history (
    run_id TEXT PRIMARY KEY,
    module_id TEXT,
    module_label TEXT,
    module_icon TEXT,
    user TEXT,
    started_at TEXT,
    finished_at TEXT,
    duration_seconds REAL,
    status TEXT,
    error_message TEXT
);
"""

FAKE_ADMIN = {
    "username": "admin",
    "role": "admin",
    "modules": [
        "edo", "overdue", "watercontrol", "utnkr", "cameras",
        "cds", "mgkh_rm", "ecur", "appeals", "municipality-report",
    ],
}


@pytest.fixture
def isolated_db(monkeypatch):
    """
    Чистая in-memory БД на ВЕСЬ тест (одно общее соединение).
    Включается явно: def test_xxx(isolated_db): ...
    """
    conn = sqlite3.connect(":memory:", check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)

    from contextlib import contextmanager

    @contextmanager
    def fake_get_db_connection():
        yield conn

    import utils.db
    monkeypatch.setattr(utils.db, "get_db_connection", fake_get_db_connection)

    yield conn
    conn.close()


@pytest.fixture
def api_client():
    """Неавторизованный FastAPI TestClient (реальная БД)."""
    from fastapi.testclient import TestClient
    from app import app
    return TestClient(app)


@pytest.fixture
def admin_client(api_client):
    """TestClient, авторизованный как админ (auth замокан)."""
    with patch("app.get_user_from_token", return_value=FAKE_ADMIN), \
         patch("services.auth.security.get_user_from_token", return_value=FAKE_ADMIN), \
         patch("app.has_module_access", return_value=True):
        api_client.cookies.set("access_token", "fake-admin-token")
        yield api_client
        