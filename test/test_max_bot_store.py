"""Linking tests use only temporary databases and synthetic routing identities."""
import hashlib
import sqlite3
from concurrent.futures import ThreadPoolExecutor
from types import SimpleNamespace

import pytest

from services.max_bot.store import ALPHABET, Store, StoreError

NOW = 1_800_000_000
MAX_ID = 123456789
CHAT_ID = 987654321


@pytest.fixture
def store(tmp_path):
    account_db = tmp_path / "accounts.sqlite"
    with sqlite3.connect(account_db) as conn:
        conn.execute("CREATE TABLE users(id INTEGER PRIMARY KEY,username TEXT,role TEXT,modules TEXT,is_active INTEGER,password_hash TEXT)")
        conn.executemany("INSERT INTO users VALUES (?,?,?,?,?,?)", [
            (1, "reader", "Пользователь", '["zips"]', 1, "synthetic-password-hash-1"),
            (2, "second", "Пользователь", '["zips"]', 1, "synthetic-password-hash-2"),
            (3, "admin", "Администратор", '[]', 1, "synthetic-password-hash-3"),
            (4, "blocked", "Пользователь", '["zips"]', 0, "synthetic-password-hash-4"),
        ])
    return Store(SimpleNamespace(accounts_path=account_db, db_path=tmp_path / "max_bot.sqlite",
                                 link_ttl=300, binding_ttl=2592000, audit_days=30))


def fails(code, function, *args, **kwargs):
    with pytest.raises(StoreError) as caught:
        function(*args, **kwargs)
    assert caught.value.code == code
    return caught.value


def pending(store, user_id=1, max_id=MAX_ID, chat_id=CHAT_ID, now=NOW):
    code = store.create_link(user_id, now=now)["code"]
    return store.redeem_link(code, max_id, chat_id, now=now)["request_id"]


def linked(store, user_id=1, max_id=MAX_ID, chat_id=CHAT_ID, now=NOW):
    request_id = pending(store, user_id, max_id, chat_id, now)
    return store.confirm_link(user_id, request_id, now=now)


def change(store, statement, parameters=()):
    with sqlite3.connect(store.accounts_path) as conn:
        conn.execute(statement, parameters)


def test_two_steps_required_and_code_not_stored_in_clear(store):
    original = store.accounts_path.read_bytes()
    result = store.create_link(1, now=NOW)
    code = result["code"]
    assert len(code) == 12 and all(c in ALPHABET for c in code)
    assert result["expires_at"] == NOW + 300
    assert code.encode() not in store.db_path.read_bytes()
    request = store.redeem_link(code, MAX_ID, CHAT_ID, now=NOW)
    fails("LINK_REQUIRED", store.authorize, MAX_ID, CHAT_ID, now=NOW)
    state = store.binding_status(1, now=NOW)
    assert state["active"] is False
    assert state["pending"][0]["request_id"] == request["request_id"]
    assert str(MAX_ID) not in repr(state) and str(CHAT_ID) not in repr(state)
    assert state["pending"][0]["max_user_mask"] == "•••• 6789"
    store.confirm_link(1, request["request_id"], now=NOW)
    user = store.authorize(MAX_ID, CHAT_ID, now=NOW)
    assert user["id"] == 1 and "password_hash" not in user
    assert store.accounts_path.read_bytes() == original


def test_code_human_format_normalizes_without_extending_validity(store):
    code = store.create_link(1, now=NOW)["code"]
    formatted = "-".join(code[i:i+4].lower() for i in range(0, 12, 4))
    assert store.redeem_link(formatted, MAX_ID, CHAT_ID, now=NOW + 299)["status"] == "pending"
    fails("LINK_INVALID", store.redeem_link, code, MAX_ID, CHAT_ID, now=NOW + 299)


@pytest.mark.parametrize("time_delta", [300, 301, 600])
def test_code_expires_at_five_minutes(store, time_delta):
    code = store.create_link(1, now=NOW)["code"]
    fails("LINK_INVALID", store.redeem_link, code, MAX_ID, CHAT_ID, now=NOW + time_delta)


def test_pending_requires_owner_and_expires(store):
    request_id = pending(store)
    fails("LINK_INVALID", store.confirm_link, 2, request_id, now=NOW)
    fails("LINK_INVALID", store.confirm_link, 1, request_id, now=NOW + 300)
    assert store.binding_status(1, now=NOW + 300)["pending"] == []


def test_new_code_invalidates_old_codes_and_pending_not_existing_binding(store):
    linked(store)
    old = store.create_link(1, now=NOW + 1)["code"]
    newer = store.create_link(1, now=NOW + 2)["code"]
    fails("LINK_INVALID", store.redeem_link, old, MAX_ID + 1, CHAT_ID + 1, now=NOW + 2)
    request_id = store.redeem_link(newer, MAX_ID + 1, CHAT_ID + 1, now=NOW + 2)["request_id"]
    store.create_link(1, now=NOW + 3)
    fails("LINK_INVALID", store.confirm_link, 1, request_id, now=NOW + 3)
    assert store.authorize(MAX_ID, CHAT_ID, now=NOW + 3)["id"] == 1


def test_confirm_replaces_only_own_binding_and_never_transfers_other_account(store):
    linked(store)
    foreign = pending(store, user_id=2)
    fails("LINK_CONFLICT", store.confirm_link, 2, foreign, now=NOW)
    assert store.authorize(MAX_ID, CHAT_ID, now=NOW)["id"] == 1
    request_id = pending(store, max_id=MAX_ID + 1, chat_id=CHAT_ID + 1)
    store.confirm_link(1, request_id, now=NOW)
    fails("LINK_REQUIRED", store.authorize, MAX_ID, CHAT_ID, now=NOW)
    assert store.authorize(MAX_ID + 1, CHAT_ID + 1, now=NOW)["id"] == 1


def test_one_time_code_and_confirmation_are_atomic(store):
    code = store.create_link(1, now=NOW)["code"]
    def redeem(n):
        try:
            return store.redeem_link(code, MAX_ID + n, CHAT_ID + n, now=NOW)["request_id"]
        except StoreError as exc:
            return exc.code
    with ThreadPoolExecutor(max_workers=4) as pool:
        results = list(pool.map(redeem, range(4)))
    assert results.count("LINK_INVALID") == 3
    request_id = next(value for value in results if value != "LINK_INVALID")
    def confirm(_):
        try:
            store.confirm_link(1, request_id, now=NOW)
            return "ok"
        except StoreError as exc:
            return exc.code
    with ThreadPoolExecutor(max_workers=2) as pool:
        results = list(pool.map(confirm, range(2)))
    assert sorted(results) == ["LINK_INVALID", "ok"]


@pytest.mark.parametrize("user_id,code", [(3, "MODULE_FORBIDDEN"), (4, "ACCESS_REVOKED"), (999, "ACCESS_REVOKED")])
def test_admin_without_grant_and_disabled_unknown_accounts_cannot_start_linking(store, user_id, code):
    fails(code, store.create_link, user_id, now=NOW)


@pytest.mark.parametrize("stage", ["before_redeem", "before_confirm", "after_confirm"])
def test_password_change_invalidates_code_pending_and_binding(store, stage):
    code = store.create_link(1, now=NOW)["code"]
    if stage == "before_redeem":
        change(store, "UPDATE users SET password_hash='changed' WHERE id=1")
        fails("LINK_INVALID", store.redeem_link, code, MAX_ID, CHAT_ID, now=NOW)
        return
    request_id = store.redeem_link(code, MAX_ID, CHAT_ID, now=NOW)["request_id"]
    if stage == "after_confirm":
        store.confirm_link(1, request_id, now=NOW)
    change(store, "UPDATE users SET password_hash='changed' WHERE id=1")
    if stage == "before_confirm":
        fails("LINK_INVALID", store.confirm_link, 1, request_id, now=NOW)
    else:
        fails("LINK_REQUIRED", store.authorize, MAX_ID, CHAT_ID, now=NOW)
    assert store.binding_status(1, now=NOW) == {"active": False, "pending": []}


@pytest.mark.parametrize("statement,code", [
    ("UPDATE users SET is_active=0 WHERE id=1", "ACCESS_REVOKED"),
    ("DELETE FROM users WHERE id=1", "ACCESS_REVOKED"),
    ("UPDATE users SET modules='[]' WHERE id=1", "MODULE_FORBIDDEN"),
    ("UPDATE users SET modules='invalid-json' WHERE id=1", "MODULE_FORBIDDEN"),
])
def test_live_revocation_every_request(store, statement, code):
    linked(store)
    change(store, statement)
    fails(code, store.authorize, MAX_ID, CHAT_ID, now=NOW)
    # Unlink can always be completed by the existing authenticated platform session.
    store.unlink(1)


def test_identity_and_chat_must_both_match_and_expiry_is_absolute(store):
    linked(store)
    fails("LINK_REQUIRED", store.authorize, MAX_ID + 1, CHAT_ID, now=NOW)
    fails("LINK_REQUIRED", store.authorize, MAX_ID, CHAT_ID + 1, now=NOW)
    assert store.authorize(MAX_ID, CHAT_ID, now=NOW + 2591999)["id"] == 1
    fails("LINK_REQUIRED", store.authorize, MAX_ID, CHAT_ID, now=NOW + 2592000)
    assert not store.binding_status(1, now=NOW + 2592000)["active"]


def test_unlink_clears_pending_codes_and_persisted_binding(store):
    linked(store)
    code = store.create_link(1, now=NOW)["code"]
    store.unlink(1)
    fails("LINK_REQUIRED", store.authorize, MAX_ID, CHAT_ID, now=NOW)
    fails("LINK_INVALID", store.redeem_link, code, MAX_ID, CHAT_ID, now=NOW)
    assert store.binding_status(1, now=NOW) == {"active": False, "pending": []}


@pytest.mark.parametrize("value", [True, 0, -1, "1' OR 1=1", "1e2", [], "9" * 30])
def test_invalid_account_ids_fail_before_sql(store, value):
    fails("INVALID_ID", store.create_link, value, now=NOW)


def test_missing_database_is_not_empty_inventory_or_new_account_database(store):
    store.accounts_path.unlink()
    fails("SYSTEM_UNAVAILABLE", store.create_link, 1, now=NOW)
    assert not store.accounts_path.exists()


def test_read_only_empty_user_store_never_seeds_accounts(store):
    change(store, "DELETE FROM users")
    before = store.accounts_path.read_bytes()
    fails("ACCESS_REVOKED", store.create_link, 1, now=NOW)
    assert store.accounts_path.read_bytes() == before


def test_rate_limit_is_atomic_persistent_and_keyed(store):
    def attempt(_):
        try:
            store.rate_limit("test", "routing-id", 3, 60, now=NOW)
            return True
        except StoreError as exc:
            assert exc.code == "RATE_LIMITED"
            return False
    with ThreadPoolExecutor(max_workers=8) as pool:
        assert sum(pool.map(attempt, range(12))) == 3
    restarted = Store(store.settings)
    fails("RATE_LIMITED", restarted.rate_limit, "test", "routing-id", 3, 60, now=NOW)
    restarted.rate_limit("test", "routing-id", 3, 60, now=NOW + 60)
    with sqlite3.connect(store.db_path) as conn:
        digest = conn.execute("SELECT key_hash FROM max_bot_rate_events LIMIT 1").fetchone()[0]
    assert digest != hashlib.sha256(b"routing-id").hexdigest()


def test_invalid_redeem_attempts_are_limited(store):
    for _ in range(5):
        fails("LINK_INVALID", store.redeem_link, "not-a-code", MAX_ID, CHAT_ID, now=NOW)
    fails("RATE_LIMITED", store.redeem_link, "not-a-code", MAX_ID, CHAT_ID, now=NOW)


def test_polling_lease_excludes_second_worker_and_stale_owner_cannot_advance_marker(store):
    assert store.lease("one", 60, now=NOW)
    assert not store.lease("two", 60, now=NOW + 1)
    assert store.lease("one", 60, now=NOW + 2)
    assert store.get_marker() is None
    store.set_marker(1234567890123456789, owner="one", now=NOW + 3)
    assert store.get_marker() == 1234567890123456789
    assert store.lease("two", 60, now=NOW + 62)
    fails("LEASE_LOST", store.set_marker, 2, owner="one", now=NOW + 62)
    store.release_lease("one")
    assert not store.lease("three", 60, now=NOW + 63)
    store.release_lease("two")
    assert store.lease("three", 60, now=NOW + 64)


def test_event_claim_is_durable_atomic_and_uncertain_sends_are_never_retried(store):
    with ThreadPoolExecutor(max_workers=8) as pool:
        assert sum(pool.map(lambda _: store.claim_event("same-message-id", now=NOW), range(8))) == 1
    store.finish_event("same-message-id", "outbound_uncertain", now=NOW)
    restarted = Store(store.settings)
    assert not restarted.claim_event("same-message-id", now=NOW + 100)
    assert "same-message-id".encode() not in store.db_path.read_bytes()
    assert restarted.claim_event("different-message-id", now=NOW)


def test_audit_only_fixed_codes_retention_and_no_secret_output(store, capsys):
    code = store.create_link(1, now=NOW)["code"]
    assert store.audit("link_create", "ok", now=NOW - 31 * 86400)
    assert store.audit("search", "ok", now=NOW)
    with sqlite3.connect(store.db_path) as conn:
        assert conn.execute("SELECT occurred_at,event,outcome FROM max_bot_audit").fetchall() == [(NOW, "search", "ok")]
    with pytest.raises(ValueError):
        store.audit("search", "private query content", now=NOW)
    output = capsys.readouterr()
    assert not output.out and not output.err
    assert code.encode() not in store.db_path.read_bytes()


def test_migration_idempotence_and_separate_account_database(store):
    linked(store)
    assert Store(store.settings).authorize(MAX_ID, CHAT_ID, now=NOW)["id"] == 1
    store.settings.db_path = store.accounts_path
    with pytest.raises(ValueError):
        Store(store.settings)


def test_bot_identity_is_immutable_and_preserves_marker_and_bindings(store):
    store.ensure_bot_id(111)
    store.set_marker(42)
    linked(store)
    restarted = Store(store.settings)
    restarted.ensure_bot_id("111")
    fails("BOT_ID_CHANGED", restarted.ensure_bot_id, 222)
    assert restarted.get_marker() == 42
    assert restarted.authorize(MAX_ID, CHAT_ID, now=NOW)["id"] == 1


def test_stop_unlinks_revoked_user_but_other_chat_cannot_revoke(store):
    linked(store)
    store.unlink_max(MAX_ID, CHAT_ID + 1)
    assert store.authorize(MAX_ID, CHAT_ID, now=NOW)["id"] == 1
    change(store, "UPDATE users SET is_active=0 WHERE id=1")
    store.unlink_max(MAX_ID, CHAT_ID)
    fails("LINK_REQUIRED", store.authorize, MAX_ID, CHAT_ID, now=NOW)
    # Repeated stop is harmless and does not require local credentials.
    store.unlink_max(MAX_ID, CHAT_ID)
