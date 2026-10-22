"""Actual API/storage tests with temporary SQLite only; no app startup or network."""
from concurrent.futures import ThreadPoolExecutor
import copy

from fastapi import FastAPI
from fastapi.testclient import TestClient
import pytest

from routers import release_notes as routes
from services import release_notes as releases
from services.release_catalog import RELEASES as DEPLOYED_RELEASES
from utils import db


RELEASE_ID = "2026-10-15-v3"
CATALOG = (
    {"id": "2026-10-15-test", "version": "15.10.2026", "title": "Обновление сервиса",
     "published_at": "2026-10-15T00:00:00+03:00", "entries": (
         {"id": "everyone", "audience": "all", "title": "Общее улучшение", "text": "Доступно всем"},
         {"id": "zip", "audience": "module", "module_id": "zips", "title": "ЗиП", "text": "Сведения о ЗиП"},
         {"id": "map", "audience": "module", "module_id": "mingkh", "title": "Карта", "text": "Сведения о карте"},
         {"id": "admin", "audience": "admin", "title": "Администраторам", "text": "Настройка управления"},
     )},
)


@pytest.fixture
def isolated(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "DB_FILE", tmp_path / "release.sqlite")
    monkeypatch.setattr(db, "_SCHEMA_READY", False)
    # Most storage tests need a module-only release to exercise the invisible
    # release case; the real v3 catalog deliberately includes global news.
    monkeypatch.setattr(releases, "RELEASES", ({
        "id": RELEASE_ID, "version": "3", "title": "Обновление сервиса",
        "published_at": "2026-10-15T00:00:00+03:00",
        "entries": ({"id": "zip", "audience": "module", "module_id": "zips",
                     "title": "ЗиП", "text": "Проверка исходных данных"},),
    },))
    with db.get_db_connection() as conn:
        conn.executemany("INSERT INTO users(id,username,password_hash,role,modules) VALUES (?,?,?,?,?)", [
            (1, "zip_user", "unused", "Пользователь", '["zips"]'),
            (2, "map_user", "unused", "Администратор", '["mingkh"]'),
            (3, "manager", "unused", "Управление пользователями", '[]'),
            (4, "delegated", "unused", "Пользователь", '["zips"]'),
        ])
        conn.execute("UPDATE account_control SET manager_user_id=3,grants_migrated=1")
        conn.execute("INSERT INTO account_managers(user_id) VALUES(4)")
    return {str(i): {"id": i, "username": name, "modules": ["zips", "mingkh"], "is_active": True}
            for i, name in [(1, "zip_user"), (2, "map_user"), (3, "manager"), (4, "delegated")]}


@pytest.fixture
def app(isolated, monkeypatch):
    application = FastAPI()
    monkeypatch.setattr(routes, "get_user_from_token", lambda token: None)

    @application.middleware("http")
    async def identity(request, call_next):
        request.state.user = isolated.get(request.headers.get("X-Test-User"))
        return await call_next(request)

    application.include_router(routes.router)
    return application


def api(app, actor="1"):
    return TestClient(app, headers={"X-Test-User": actor} if actor else {})


def test_deployed_v32_catalog_has_only_current_news_and_filters_module_details(app, monkeypatch):
    monkeypatch.setattr(releases, "RELEASES", DEPLOYED_RELEASES)
    releases._validate_catalog()
    assert len(DEPLOYED_RELEASES) == 1
    with api(app) as client:
        response = client.get("/api/releases/pending")
    assert response.status_code == 200
    assert response.headers["cache-control"] == "private, no-store"
    payload = response.json()
    assert payload["pending_count"] == 1
    note, = payload["releases"]
    assert note["id"] == "2026-10-22-max-v32"
    assert note["version"] == "3.2" and note["read_at"] is None
    assert {entry["id"] for entry in note["entries"]} == {
        "max-guided-connection", "max-portable-certificate",
        "max-contacts-source-explanation", "max-mobile-connect-controls",
    }
    assert all(entry["module_id"] == "zips" for entry in note["entries"])
    for user in ("2", "3"):
        with api(app, user) as client:
            assert client.get("/api/releases/pending").json() == {"releases": [], "pending_count": 0}
            assert client.get("/api/releases/history").json()["releases"] == []
    with db.get_db_connection() as conn:
        conn.execute("UPDATE users SET modules='[]' WHERE id=1")
    with api(app) as client:
        assert client.get("/api/releases/pending").json()["releases"] == []


@pytest.mark.parametrize("path", ["/api/releases/pending", "/api/releases/history"])
def test_unauthenticated_cannot_read_release_notes(app, path):
    with api(app, None) as client:
        assert client.get(path).status_code == 401
        assert client.post(f"/api/releases/{RELEASE_ID}/read").status_code == 401


def test_current_db_grants_override_stale_token_and_role(app):
    with api(app, "1") as client:
        assert client.get("/api/releases/pending").json()["pending_count"] == 1
        with db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[]' WHERE id=1")
        assert client.get("/api/releases/pending").json()["releases"] == []
        assert client.get("/api/releases/history").json()["releases"] == []
        assert client.post(f"/api/releases/{RELEASE_ID}/read").status_code == 404
    with api(app, "2") as client:
        assert client.get("/api/releases/history").json()["releases"] == []


def test_acknowledgement_persists_across_clients_and_does_not_change_other_account(app):
    with api(app) as first, api(app) as second, api(app, "4") as another:
        assert first.get("/api/releases/pending").json()["pending_count"] == 1
        acknowledged = first.post(f"/api/releases/{RELEASE_ID}/read")
        assert acknowledged.status_code == 200
        assert second.get("/api/releases/pending").json() == {"releases": [], "pending_count": 0}
        assert another.get("/api/releases/pending").json()["pending_count"] == 1
        history = second.get("/api/releases/history").json()
        assert history["pending_count"] == 0 and len(history["releases"]) == 1
        assert history["releases"][0]["read_at"] == acknowledged.json()["read_at"]
        assert second.post(f"/api/releases/{RELEASE_ID}/read").json() == acknowledged.json()


def test_parallel_first_api_calls_and_repeated_acknowledgement_are_idempotent(app, isolated):
    with ThreadPoolExecutor(max_workers=8) as executor:
        outcomes = list(executor.map(lambda _: releases.acknowledge(isolated["1"], RELEASE_ID), range(16)))
    assert all(item == outcomes[0] for item in outcomes)
    with db.get_db_connection() as conn:
        assert conn.execute("SELECT count(*) FROM release_note_receipts").fetchone()[0] == 1
        assert conn.execute("SELECT count(*) FROM app_schema_migrations WHERE name=?", (releases.MIGRATION_ID,)).fetchone()[0] == 1


def test_request_cannot_choose_someone_else_or_rewrite_acknowledgement(app):
    with api(app) as client:
        response = client.post(f"/api/releases/{RELEASE_ID}/read", json={"user_id": 4})
        assert response.status_code == 422
        assert client.get("/api/releases/pending").json()["pending_count"] == 1
        assert client.post(f"/api/releases/{RELEASE_ID}/read", json={}).status_code == 200
    with db.get_db_connection() as conn:
        account, = conn.execute("SELECT account_key FROM release_note_receipts").fetchone()
    assert account == "local:1"


def test_unknown_and_inaccessible_release_ids_are_not_acknowledged(app):
    with api(app) as client:
        assert client.post("/api/releases/unknown/read").status_code == 404
    with api(app, "2") as client:
        assert client.post(f"/api/releases/{RELEASE_ID}/read").status_code == 404
    with db.get_db_connection() as conn:
        assert conn.execute("SELECT count(*) FROM release_note_receipts").fetchone()[0] == 0


@pytest.mark.parametrize("actor, expected", [("1", {"everyone", "zip"}), ("2", {"everyone", "map"}),
                                             ("3", {"everyone", "admin"}), ("4", {"everyone", "zip", "admin"})])
def test_mixed_release_is_filtered_per_entry_and_current_manager_rights(app, monkeypatch, actor, expected):
    monkeypatch.setattr(releases, "RELEASES", CATALOG)
    with api(app, actor) as client:
        payload = client.get("/api/releases/pending").json()
        assert {entry["id"] for entry in payload["releases"][0]["entries"]} == expected
        if actor == "4":
            with db.get_db_connection() as conn:
                conn.execute("DELETE FROM account_managers WHERE user_id=4")
            payload = client.get("/api/releases/history").json()
            assert {entry["id"] for entry in payload["releases"][0]["entries"]} == {"everyone", "zip"}


def test_disabled_or_deleted_local_account_is_rejected(app):
    with api(app) as client:
        with db.get_db_connection() as conn:
            conn.execute("UPDATE users SET is_active=0 WHERE id=1")
        assert client.get("/api/releases/pending").status_code == 401
        assert client.post(f"/api/releases/{RELEASE_ID}/read").status_code == 401
        with db.get_db_connection() as conn:
            conn.execute("DELETE FROM users WHERE id=1")
        assert client.get("/api/releases/history").status_code == 401


def test_keycloak_identity_is_namespaced_and_uses_existing_permission_model(app, isolated):
    isolated["external"] = {"kc_sub": "1", "roles": ["zips"], "role": "admin", "is_active": True}
    with api(app, "external") as external, api(app, "1") as local:
        assert external.get("/api/releases/pending").json()["pending_count"] == 1
        assert external.post(f"/api/releases/{RELEASE_ID}/read").status_code == 200
        assert local.get("/api/releases/pending").json()["pending_count"] == 1
    with db.get_db_connection() as conn:
        assert conn.execute("SELECT account_key FROM release_note_receipts").fetchone()[0] == "keycloak:1"


def test_reading_pending_does_not_mark_seen_or_touch_ordinary_notifications(app):
    with db.get_db_connection() as conn:
        before = [dict(row) for row in conn.execute("SELECT * FROM account_notifications")]
    with api(app) as client:
        assert client.get("/api/releases/pending").json()["pending_count"] == 1
        assert client.get("/api/releases/history").json()["pending_count"] == 1
        assert client.get("/api/releases/pending").json()["pending_count"] == 1
        client.post(f"/api/releases/{RELEASE_ID}/read")
    with db.get_db_connection() as conn:
        assert [dict(row) for row in conn.execute("SELECT * FROM account_notifications")] == before


def test_migration_reentry_preserves_accounts_receipts_and_existing_schema(app):
    with api(app) as client:
        first = client.post(f"/api/releases/{RELEASE_ID}/read").json()
        releases.initialize()
        releases.initialize()
        assert client.post(f"/api/releases/{RELEASE_ID}/read").json() == first
    with db.get_db_connection() as conn:
        assert conn.execute("SELECT count(*) FROM users").fetchone()[0] == 4
        assert conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='ux_survey_responses'").fetchone()


def test_failed_receipt_write_stays_unread_and_can_be_retried(app):
    releases.initialize()
    with db.get_db_connection() as conn:
        conn.execute("CREATE TRIGGER fail_receipt BEFORE INSERT ON release_note_receipts "
                     "BEGIN SELECT RAISE(ABORT, 'synthetic write failure'); END")
    with TestClient(app, headers={"X-Test-User": "1"}, raise_server_exceptions=False) as client:
        assert client.post(f"/api/releases/{RELEASE_ID}/read").status_code == 500
        assert client.get("/api/releases/pending").json()["pending_count"] == 1
        with db.get_db_connection() as conn:
            conn.execute("DROP TRIGGER fail_receipt")
        assert client.post(f"/api/releases/{RELEASE_ID}/read").status_code == 200
        assert client.get("/api/releases/pending").json()["pending_count"] == 0


def test_new_release_is_independent_from_previous_acknowledgement(app, monkeypatch):
    with api(app) as client:
        client.post(f"/api/releases/{RELEASE_ID}/read")
        newer = copy.deepcopy(releases.RELEASES[0])
        newer.update(id="2026-10-16-next", version="16.10.2026", published_at="2026-10-16T00:00:00+03:00")
        monkeypatch.setattr(releases, "RELEASES", (releases.RELEASES[0], newer))
        pending = client.get("/api/releases/pending").json()
        assert pending["pending_count"] == 1 and pending["releases"][0]["id"] == newer["id"]
        history = client.get("/api/releases/history").json()
        assert [item["id"] for item in history["releases"]] == [newer["id"], RELEASE_ID]


@pytest.mark.parametrize("bad_modules", ['null', '42', '{}', 'not-json', '[null, {}, "unknown"]'])
def test_malformed_local_grants_fail_closed(app, bad_modules):
    with db.get_db_connection() as conn:
        conn.execute("UPDATE users SET modules=? WHERE id=1", (bad_modules,))
    with api(app) as client:
        assert client.get("/api/releases/pending").json() == {"releases": [], "pending_count": 0}


def test_invisible_release_is_not_acknowledged_implicitly_before_later_grant(app):
    with api(app, "2") as client:
        assert client.get("/api/releases/pending").json()["releases"] == []
        with db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[\"mingkh\",\"zips\"]' WHERE id=2")
        assert client.get("/api/releases/pending").json()["pending_count"] == 1


def test_catalog_rejects_unknown_module_instead_of_broadening_audience(isolated, monkeypatch):
    broken = copy.deepcopy(CATALOG)
    broken[0]["entries"][1]["module_id"] = "zip_curator"
    monkeypatch.setattr(releases, "RELEASES", broken)
    with pytest.raises(ValueError, match="module"):
        releases.pending(isolated["1"])


def test_water_dashboard_details_require_current_water_grant(app, monkeypatch):
    catalog = copy.deepcopy(CATALOG)
    catalog[0]["entries"][2]["module_id"] = "water-dashboard"
    monkeypatch.setattr(releases, "RELEASES", catalog)
    with api(app) as client:
        assert "map" not in {entry["id"] for note in client.get("/api/releases/pending").json()["releases"] for entry in note["entries"]}
        with db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[\"water-dashboard\"]' WHERE id=1")
        notes = client.get("/api/releases/pending").json()["releases"]
        assert {entry["id"] for note in notes for entry in note["entries"]} == {"everyone", "map"}


def test_old_receipts_do_not_keep_previous_versions_in_history(app, monkeypatch):
    with api(app) as client:
        assert client.post(f"/api/releases/{RELEASE_ID}/read").status_code == 200
        monkeypatch.setattr(releases, "RELEASES", DEPLOYED_RELEASES)
        history = client.get("/api/releases/history").json()
        assert history["pending_count"] == 1
        assert [item["version"] for item in history["releases"]] == ["3.2"] * len(DEPLOYED_RELEASES)
        assert client.post(f"/api/releases/{RELEASE_ID}/read").status_code == 404


def test_collective_release_details_require_collective_grant(app, monkeypatch):
    catalog = copy.deepcopy(CATALOG)
    catalog[0]["entries"][2]["module_id"] = "collective"
    monkeypatch.setattr(releases, "RELEASES", catalog)
    with db.get_db_connection() as conn:
        conn.execute("UPDATE users SET modules=? WHERE id=1", ('["collective"]',))
    with api(app) as client:
        notes = client.get("/api/releases/pending").json()["releases"]
        details = [entry for note in notes for entry in note["entries"] if entry["module_id"]]
        assert [(entry["id"], entry["module_id"]) for entry in details] == [
            ("map", "collective")]
        with db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[]' WHERE id=1")
        notes = client.get("/api/releases/pending").json()["releases"]
        assert not any(entry["module_id"] for note in notes for entry in note["entries"])
