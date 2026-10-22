"""Synthetic contact-source recovery through the same linked MAX dialog."""
import json
import os
import sqlite3

import pytest

from test.test_max_bot_handler import handler, event, callback, link, button


def _atomic_contacts(handler, payload):
    """Match the canonical importer's replace-on-success publication pattern."""
    destination = handler.stock.contacts_path
    temporary = destination.with_name(destination.name + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    os.replace(temporary, destination)


def _source():
    return {"items": [{"rso": "РСО Тест", "municipality": "Власиха",
                       "person": "Новый тестовый инженер", "position": "Ответственный",
                       "phone": "+7 000 222-33-44", "email": "private@example.invalid",
                       "note": "PRIVATE_CONTACT_NOTE"}]}


def _contact_button(handler):
    link(handler)
    return button(handler.handle(event("муфта D100")))


def test_missing_contacts_explain_same_installation_upload_without_false_retry(handler):
    token = _contact_button(handler)
    handler.stock.contacts_path.unlink()
    reply = handler.handle(callback(token))
    message = reply["body"]["text"]
    assert "ещё не загружен" in message
    assert "том же компьютере или сервере" in message
    assert "Остатки ЗиП РСО" in message
    assert "Контакты ответственных от РСО" in message
    assert "Загрузить контакты" in message
    assert "позже" not in message.lower()
    assert str(handler.stock.contacts_path) not in message
    assert "contacts.json" not in message and "Телефон:" not in message
    assert not handler.stock.contacts_path.exists()  # A failed read never invents a source.
    assert reply["callback_id"] == "cb-test"


def test_existing_handler_and_button_recover_after_atomic_canonical_import(handler):
    token = _contact_button(handler)
    handler.stock.contacts_path.unlink()
    assert "ещё не загружен" in handler.handle(callback(token))["body"]["text"]

    _atomic_contacts(handler, _source())
    recovered = handler.handle(callback(token))
    assert "Новый тестовый инженер" in recovered["body"]["text"]
    assert "+7 000 222-33-44" in recovered["body"]["text"]
    assert "Тестовый инженер" not in recovered["body"]["text"]
    assert "private@example.invalid" not in str(recovered)
    assert "PRIVATE_CONTACT_NOTE" not in str(recovered)

    # The importer can also correct an existing contact without restarting polling.
    corrected = _source()
    corrected["items"][0]["phone"] = "+7 000 555-66-77"
    _atomic_contacts(handler, corrected)
    current = handler.handle(callback(token))["body"]["text"]
    assert "+7 000 555-66-77" in current and "+7 000 222-33-44" not in current


@pytest.mark.parametrize("raw", ["{broken", '{"items":[null]}', '{"items":{}}'])
def test_corrupt_existing_source_remains_unavailable_until_replaced(handler, raw):
    token = _contact_button(handler)
    handler.stock.contacts_path.write_text(raw, encoding="utf-8")
    message = handler.handle(callback(token))["body"]["text"]
    assert "временно недоступен" in message
    assert "ещё не загружен" not in message
    assert "не найден" not in message and "Телефон:" not in message
    assert handler.stock.contacts_path.read_text(encoding="utf-8") == raw

    _atomic_contacts(handler, _source())
    assert "+7 000 222-33-44" in handler.handle(callback(token))["body"]["text"]


def test_valid_empty_directory_is_explained_without_outage_or_false_contact(handler):
    token = _contact_button(handler)
    _atomic_contacts(handler, {"items": []})
    message = handler.handle(callback(token))["body"]["text"]
    assert "пуст" in message and "Загрузите контакты" in message
    assert "временно недоступен" not in message and "Телефон:" not in message


@pytest.mark.parametrize("sql", [
    "UPDATE users SET modules='[]' WHERE id=1",
    "UPDATE users SET is_active=0 WHERE id=1",
    "UPDATE users SET password_hash='changed' WHERE id=1",
])
def test_source_recovery_never_bypasses_fresh_account_authorization(handler, monkeypatch, sql):
    token = _contact_button(handler)
    handler.stock.contacts_path.unlink()
    assert "ещё не загружен" in handler.handle(callback(token))["body"]["text"]
    _atomic_contacts(handler, _source())
    with sqlite3.connect(handler.store.accounts_path) as connection:
        connection.execute(sql)

    def forbidden_read(*args, **kwargs):
        pytest.fail("Contact source must not be read after account access is revoked")

    monkeypatch.setattr(handler.stock, "contacts", forbidden_read)
    reply = handler.handle(callback(token))
    assert "Новый тестовый инженер" not in str(reply)
    assert "+7 000 222-33-44" not in str(reply)
    message = reply["body"]["text"].lower()
    assert "доступ" in message or "отозван" in message or "привяжите" in message
