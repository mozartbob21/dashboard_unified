"""Permanent MAX links use temporary databases, never live credentials or routing."""
import hashlib
import sqlite3
import time
from concurrent.futures import ThreadPoolExecutor
from types import SimpleNamespace

import pytest

from services.max_bot.store import MIGRATION, PERMANENT_BINDING_EXPIRY, Store
from test.test_max_bot_store import CHAT_ID, MAX_ID, NOW, change, fails, linked, pending, store


@pytest.fixture
def legacy_settings(tmp_path):
    """A pre-upgrade store, before the new migration marker exists."""
    settings = SimpleNamespace(db_path=tmp_path / "legacy.sqlite", accounts_path=tmp_path / "accounts.sqlite",
                               link_ttl=300, audit_days=30)
    with sqlite3.connect(settings.accounts_path) as conn:
        conn.executescript("""
            CREATE TABLE users(id INTEGER PRIMARY KEY,username TEXT,role TEXT,modules TEXT,is_active INTEGER,password_hash TEXT);
            CREATE TABLE account_control(id INTEGER PRIMARY KEY, manager_user_id INTEGER);
            CREATE TABLE account_managers(user_id INTEGER PRIMARY KEY);
            INSERT INTO account_control VALUES(1,1);
            INSERT INTO users VALUES(1,'reader','Пользователь','["presentations"]',1,'synthetic-password');
        """)
    with sqlite3.connect(settings.db_path) as conn:
        conn.executescript(MIGRATION.read_text(encoding="utf-8"))
    return settings


def seed_legacy(settings, expires_at, created_at=None, *, user_id=1):
    created_at = time.time() - 60 if created_at is None else created_at
    row = (user_id, str(MAX_ID + user_id - 1), str(CHAT_ID + user_id - 1),
           hashlib.sha256(b"synthetic-password").hexdigest(), created_at, expires_at)
    with sqlite3.connect(settings.db_path) as conn:
        conn.execute("INSERT INTO max_bot_bindings VALUES(?,?,?,?,?,?)", row)
    return row


def binding_rows(path):
    with sqlite3.connect(path) as conn:
        return conn.execute("SELECT * FROM max_bot_bindings ORDER BY user_id").fetchall()


def test_confirmed_binding_survives_restart_and_years_without_activity(store):
    confirmed = linked(store)
    assert confirmed["binding"] == {"max_user_mask": "•••• 6789", "expires_at": None, "permanent": True}
    assert binding_rows(store.db_path)[0][-1] == PERMANENT_BINDING_EXPIRY
    restarted = Store(store.settings)
    future = NOW + 20 * 365 * 86400
    assert restarted.authorize(MAX_ID, CHAT_ID, now=future)["id"] == 1
    assert restarted.binding_status(1, now=future)["binding"] == confirmed["binding"]


def test_old_worker_cleanup_during_restart_does_not_delete_permanent_link(store):
    linked(store)
    with sqlite3.connect(store.db_path) as conn:
        # Match the exact legacy cleanup statement during mixed-version startup.
        conn.execute("DELETE FROM max_bot_bindings WHERE expires_at<=?", (NOW + 365 * 86400,))
    assert Store(store.settings).authorize(MAX_ID, CHAT_ID, now=NOW + 365 * 86400)["id"] == 1


def test_permanent_marker_is_not_a_far_future_expiration_deadline(store):
    linked(store)
    after_marker = PERMANENT_BINDING_EXPIRY + 100
    assert store.authorize(MAX_ID, CHAT_ID, now=after_marker)["id"] == 1
    assert store.binding_status(1, now=after_marker)["binding"]["permanent"] is True


def test_cleanup_expires_codes_and_pending_but_never_permanent_binding(store):
    linked(store)
    pending(store, max_id=MAX_ID + 1, chat_id=CHAT_ID + 1, now=NOW + 1)
    state = store.binding_status(1, now=NOW + 20 * 365 * 86400)
    assert state == {"active": True, "pending": [], "binding": {
        "max_user_mask": "•••• 6789", "expires_at": None, "permanent": True}}
    with sqlite3.connect(store.db_path) as conn:
        assert conn.execute("SELECT COUNT(*) FROM max_bot_codes").fetchone()[0] == 0
        assert conn.execute("SELECT COUNT(*) FROM max_bot_pending").fetchone()[0] == 0
        assert conn.execute("SELECT COUNT(*) FROM max_bot_bindings").fetchone()[0] == 1


def test_active_legacy_upgrade_changes_only_expiration_preserving_subscription_generation(legacy_settings):
    before = seed_legacy(legacy_settings, time.time() + 300)
    # The subscription captures the binding creation time, not its expiry.
    with sqlite3.connect(legacy_settings.db_path) as conn:
        conn.execute("CREATE TABLE synthetic_subscription(user_id INTEGER PRIMARY KEY,binding_created_at REAL)")
        conn.execute("INSERT INTO synthetic_subscription VALUES(1,?)", (before[-2],))
    migrated = Store(legacy_settings)
    assert binding_rows(legacy_settings.db_path) == [before[:-1] + (PERMANENT_BINDING_EXPIRY,)]
    assert migrated.authorize(MAX_ID, CHAT_ID, now=time.time() + 5 * 365 * 86400)["id"] == 1
    with sqlite3.connect(legacy_settings.db_path) as conn:
        assert conn.execute("SELECT binding_created_at FROM synthetic_subscription").fetchone()[0] == before[-2]
        assert conn.execute("SELECT COUNT(*) FROM max_bot_meta WHERE key='migration_20261023_max_persistent_bindings'").fetchone()[0] == 1


def test_expired_legacy_link_is_never_revived(legacy_settings):
    before = seed_legacy(legacy_settings, time.time() - 1)
    migrated = Store(legacy_settings)
    assert binding_rows(legacy_settings.db_path) == [before]
    fails("LINK_REQUIRED", migrated.authorize, MAX_ID, CHAT_ID)
    assert not migrated.binding_status(1)["active"]
    assert binding_rows(legacy_settings.db_path) == []
    assert not Store(legacy_settings).binding_status(1)["active"]


def test_expired_fractional_second_cannot_be_rounded_into_a_permanent_binding(legacy_settings, monkeypatch):
    frozen_now = NOW + 0.75
    before = seed_legacy(legacy_settings, NOW + 0.25)
    original_connect = sqlite3.connect

    def frozen_connect(*args, **kwargs):
        conn = original_connect(*args, **kwargs)
        # Freeze the SQLite clock used by the migration, including its old
        # whole-second representation, so this boundary test is deterministic.
        conn.create_function("julianday", 1, lambda _: 2440587.5 + frozen_now / 86400.0)
        conn.create_function("strftime", 2, lambda *_: str(NOW))
        return conn

    monkeypatch.setattr(sqlite3, "connect", frozen_connect)
    migrated = Store(legacy_settings)
    assert binding_rows(legacy_settings.db_path) == [before]
    fails("LINK_REQUIRED", migrated.authorize, MAX_ID, CHAT_ID, now=frozen_now)


@pytest.mark.parametrize("invalid_expiry", [0, -1, "corrupt", float("inf"), 253402300800, 1e99])
def test_malformed_legacy_expiration_does_not_become_permanent(legacy_settings, invalid_expiry):
    seed_legacy(legacy_settings, invalid_expiry)
    migrated = Store(legacy_settings)
    assert binding_rows(legacy_settings.db_path)[0][-1] == invalid_expiry
    fails("LINK_REQUIRED", migrated.authorize, MAX_ID, CHAT_ID)
    assert not migrated.binding_status(1)["active"]


def test_migration_is_atomic_and_idempotent_when_web_and_worker_start_together(legacy_settings):
    before = seed_legacy(legacy_settings, time.time() + 300)
    with ThreadPoolExecutor(max_workers=6) as pool:
        stores = list(pool.map(lambda _: Store(legacy_settings), range(6)))
    assert binding_rows(legacy_settings.db_path) == [before[:-1] + (PERMANENT_BINDING_EXPIRY,)]
    assert all(item.authorize(MAX_ID, CHAT_ID)["id"] == 1 for item in stores)
    # Once complete, a later restart must not upgrade newly inserted legacy rows.
    # This also prevents an accidental old-worker write from becoming permanent.
    with sqlite3.connect(legacy_settings.db_path) as conn:
        conn.execute("UPDATE max_bot_bindings SET expires_at=?", (time.time() + 500,))
    after = binding_rows(legacy_settings.db_path)
    Store(legacy_settings)
    assert binding_rows(legacy_settings.db_path) == after


def test_unlink_remains_effective_after_restart_and_migration(legacy_settings):
    seed_legacy(legacy_settings, time.time() + 300)
    migrated = Store(legacy_settings)
    migrated.unlink(1)
    restarted = Store(legacy_settings)
    fails("LINK_REQUIRED", restarted.authorize, MAX_ID, CHAT_ID)
    assert not restarted.binding_status(1)["active"]
    assert binding_rows(legacy_settings.db_path) == []


@pytest.mark.parametrize("statement,code", [
    ("UPDATE users SET is_active=0 WHERE id=1", "ACCESS_REVOKED"),
    ("UPDATE users SET modules='[]' WHERE id=1", "MODULE_FORBIDDEN"),
    ("DELETE FROM users WHERE id=1", "ACCESS_REVOKED"),
    ("UPDATE users SET password_hash='reset' WHERE id=1", "LINK_REQUIRED"),
])
def test_permanent_link_does_not_bypass_live_revocation(store, statement, code):
    linked(store)
    change(store, statement)
    restarted = Store(store.settings)
    fails(code, restarted.authorize, MAX_ID, CHAT_ID, now=NOW + 10 * 365 * 86400)
    restarted.unlink_max(MAX_ID, CHAT_ID)
    fails("LINK_REQUIRED", restarted.authorize, MAX_ID, CHAT_ID, now=NOW + 10 * 365 * 86400)


def test_legacy_expiry_inserted_after_migration_still_expires(store):
    linked(store)
    with sqlite3.connect(store.db_path) as conn:
        conn.execute("UPDATE max_bot_bindings SET expires_at=?", (NOW + 100,))
    assert store.authorize(MAX_ID, CHAT_ID, now=NOW + 99)["id"] == 1
    assert store.binding_status(1, now=NOW)["binding"] == {
        "max_user_mask": "•••• 6789", "expires_at": NOW + 100, "permanent": False}
    fails("LINK_REQUIRED", store.authorize, MAX_ID, CHAT_ID, now=NOW + 100)


def test_binding_expiry_configuration_cannot_reintroduce_monthly_disconnect(store):
    store.settings.binding_ttl = 300  # Legacy callers may still pass the old option.
    changed = Store(store.settings)
    linked(changed)
    assert changed.authorize(MAX_ID, CHAT_ID, now=NOW + 365 * 86400)["id"] == 1
