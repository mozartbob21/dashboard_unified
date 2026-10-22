"""Full local linking/search/contact path, with synthetic files and no MAX calls."""
from datetime import date, timedelta
import json
import sqlite3

import pytest

from services.max_bot.config import Settings
from services.max_bot.handler import BotHandler
from services.max_bot.stock import StockService
from services.max_bot.store import Store, StoreError
from services import report_municipalities


def event(value="муфта", uid=123456789, chat=321, kind="dialog", mid="test-mid"):
    return {"update_type": "message_created", "message": {
        "sender": {"user_id": uid, "is_bot": False}, "recipient": {"chat_id": chat, "chat_type": kind},
        "body": {"mid": mid, "text": value}}}


def button(reply, index=0):
    return reply["body"]["attachments"][0]["payload"]["buttons"][index][0]["payload"]


def callback(payload, uid=123456789, chat=321, kind="dialog"):
    return {"update_type": "message_callback", "callback": {"callback_id": "cb-test", "payload": payload,
            "user": {"user_id": uid}}, "message": {"recipient": {"chat_type": kind, "chat_id": chat}}}


@pytest.fixture
def handler(tmp_path, monkeypatch):
    monkeypatch.setattr(report_municipalities, "REGISTRY", tmp_path / "municipalities.json")
    accounts, published, contacts = (tmp_path / name for name in ("accounts.db", "published.json", "contacts.json"))
    with sqlite3.connect(accounts) as conn:
        conn.execute("CREATE TABLE users(id INTEGER PRIMARY KEY,username TEXT,role TEXT,modules TEXT,is_active INTEGER,password_hash TEXT)")
        conn.execute("CREATE TABLE account_control(id INTEGER PRIMARY KEY, manager_user_id INTEGER)")
        conn.execute("INSERT INTO account_control VALUES(1,1)")
        conn.execute("CREATE TABLE account_managers(user_id INTEGER PRIMARY KEY)")
        conn.execute("INSERT INTO account_managers VALUES(2)")
        conn.executemany("INSERT INTO users VALUES (?,?,?,?,?,?)", [(i, f"synthetic-{i}", "user", '["zips"]', 1, f"test-hash-{i}") for i in (1, 2)])
    published.write_text(json.dumps({"rows": [
        ["РСО", "Округ", "Наименование", "Категория", "Группа", "Количество", "Ед.", "Водоподготовка", "Дата"],
        ["РСО Тест", "Власиха", "Муфта D100", "Фитинги", "Муфта", 5, "шт", "", date.today().isoformat()],
        ["РСО Тест 2", "Дубна", "Насос D20", "Насосы", "Насос", 2, "шт", "", (date.today() - timedelta(days=31)).isoformat()]],
        "published_at": "2026-10-22T12:00:00"}), encoding="utf-8")
    contacts.write_text(json.dumps({"items": [{"id": "c1", "rso": "РСО Тест", "municipality": "Власиха",
        "person": "Тестовый инженер", "position": "Ответственный", "phone": "+7 000 111-22-33",
        "email": "not-exported@example.invalid", "note": "PRIVATE_NOTE"}]}), encoding="utf-8")
    settings = Settings(accounts_path=accounts, db_path=tmp_path / "max.db", published_path=published, contacts_path=contacts)
    return BotHandler(Store(settings), StockService(published, contacts))


def link(handler, user=1, uid=123456789, chat=321):
    code = handler.store.create_link(user)["code"]
    response = handler.handle(event("/link " + code, uid, chat))
    assert "До подтверждения данные недоступны" in response["body"]["text"]
    pending = handler.store.binding_status(user)["pending"][0]["request_id"]
    handler.store.confirm_link(user, pending)


def test_complete_path_no_data_before_confirmation_then_contact_on_click(handler):
    assert "max" in handler.handle(event())["body"]["text"].lower()
    code = handler.store.create_link(1)["code"]
    assert "6789" in handler.handle(event("/link " + code))["body"]["text"]
    assert "Муфта" not in handler.handle(event())["body"]["text"]
    handler.store.confirm_link(1, handler.store.binding_status(1)["pending"][0]["request_id"])
    result = handler.handle(event("Нужна муфта D100"))
    assert "Муфта D100" in result["body"]["text"]
    assert "Власиха" in result["body"]["text"]
    assert "+7" not in str(result)
    contact = handler.handle(callback(button(result)))
    assert "+7 000 111-22-33" in contact["body"]["text"]
    assert "PRIVATE_NOTE" not in str(contact) and "example.invalid" not in str(contact)
    assert contact["callback_id"] == "cb-test"


@pytest.mark.parametrize("kind", ["chat", "channel", "group", None])
def test_never_respond_to_business_queries_or_links_in_groups(handler, kind):
    assert handler.handle(event(kind=kind)) is None
    assert handler.handle(callback("any", kind=kind)) is None
    assert handler.event_key(event(kind=kind)) is None


def test_callback_not_transferable_to_other_user_or_chat(handler):
    link(handler)
    token = button(handler.handle(event()))
    link(handler, user=2, uid=777, chat=888)
    assert "устарела" in handler.handle(callback(token, uid=777, chat=888))["body"]["text"]
    assert "устарела" in handler.handle(callback("forged"))["body"]["text"]


@pytest.mark.parametrize("sql", ["UPDATE users SET is_active=0 WHERE id=1", "UPDATE users SET modules='[]' WHERE id=1", "UPDATE users SET password_hash='changed' WHERE id=1"])
def test_revoke_rights_checked_again_before_contact(handler, sql):
    link(handler)
    token = button(handler.handle(event()))
    with sqlite3.connect(handler.store.accounts_path) as conn:
        conn.execute(sql)
    response = handler.handle(callback(token))
    assert "+7" not in str(response) and "Тестовый инженер" not in str(response)
    assert "отозван" in response["body"]["text"].lower() or "доступ" in response["body"]["text"].lower() or "привяжите" in response["body"]["text"].lower()


def test_stop_revokes_even_after_local_account_block(handler):
    link(handler)
    with sqlite3.connect(handler.store.accounts_path) as conn:
        conn.execute("UPDATE users SET is_active=0")
    assert "отключён" in handler.handle(event("/stop"))["body"]["text"]
    with sqlite3.connect(handler.store.db_path) as conn:
        assert conn.execute("SELECT count(*) FROM max_bot_bindings").fetchone()[0] == 0


def test_source_unavailable_is_not_reported_as_zero_results(handler):
    link(handler)
    handler.stock.published_path.unlink()
    response = handler.handle(event())
    assert "недоступ" in response["body"]["text"].lower()
    assert "не найдено" not in response["body"]["text"]


def test_changed_publication_invalidates_old_contact_button(handler):
    link(handler)
    token = button(handler.handle(event()))
    handler.stock.published_path.write_text('{"rows":[]}', encoding="utf-8")
    assert "+7" not in str(handler.handle(callback(token)))


def test_multiple_positions_are_selected_separately_and_stale_not_confirmed(handler):
    link(handler)
    result = handler.handle(event("муфта;насос"))
    assert "всего комплекта" in result["body"]["text"]
    selected = handler.handle(callback(button(result, 1)))
    assert "Насос D20" in selected["body"]["text"]
    assert "уточн" in selected["body"]["text"].lower() or "устар" in selected["body"]["text"].lower()


def test_expired_button_and_restart_have_safe_fallback(handler):
    link(handler)
    token = button(handler.handle(event()))
    handler.clock = lambda: float("inf")
    assert "устарела" in handler.handle(callback(token))["body"]["text"]
    handler.actions.clear()
    assert "устарела" in handler.handle(callback(token))["body"]["text"]


def test_plain_text_no_markup_mode_or_private_data_in_payloads(handler):
    link(handler)
    result = handler.handle(event())
    assert "format" not in result["body"] and len(result["body"]["text"]) <= 4000
    token = button(result)
    assert "Власиха" not in token and "муфта" not in token and len(token) >= 24
    assert handler.event_key(event()) == handler.event_key(event())
    assert handler.event_key(event(mid="next")) != handler.event_key(event())


def test_env_token_is_server_only_no_default_secret_or_accidental_enable():
    secret = "synthetic-secret-token"
    configured = Settings.from_env({"MAX_ZIP_BOT_ENABLED": "true", "MAX_ZIP_BOT_MODE": "development-polling", "MAX_ZIP_BOT_TOKEN": secret})
    assert configured.bot_token == secret and secret not in repr(configured)
    assert not Settings.from_env({}).enabled
    with pytest.raises(ValueError):
        Settings.from_env({"MAX_ZIP_BOT_ENABLED": "true"})
