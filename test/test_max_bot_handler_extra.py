"""Independent contract checks of rendered messages and callback pagination."""
import json
from datetime import datetime
from zoneinfo import ZoneInfo

import pytest

from test.test_max_bot_handler import handler, event, callback, link, button
from services.max_bot.client import _body
from services.max_bot.config import Settings, ROOT
from services.max_bot.handler import BotHandler


def payload_by_label(reply, label):
    rows = reply["body"].get("attachments", [])[0]["payload"]["buttons"]
    return next(btn["payload"] for row in rows for btn in row if btn["text"] == label)


def publish(handler, names):
    today = datetime.now(ZoneInfo("Europe/Moscow")).date().isoformat()
    headers = ["РСО", "Округ", "Наименование", "Категория", "Группа", "Количество", "Ед.", "Дата"]
    handler.stock.published_path.write_text(json.dumps({"rows": [headers, *[
        ["РСО Тест", "Власиха", name, "Фитинги", "Муфта", 5, "шт", today] for name in names
    ]]}), encoding="utf-8")


def test_search_pagination_forward_backward_is_bound_to_same_query_and_private_identity(handler):
    link(handler)
    publish(handler, [f"Муфта D{100 + i}" for i in range(12)])
    first = handler.handle(event("муфта"))
    assert "Страница 1" in first["body"]["text"] and "Муфта D104" in first["body"]["text"]
    assert "Муфта D105" not in first["body"]["text"]
    second = handler.handle(callback(payload_by_label(first, "Следующие →")))
    assert "Страница 2" in second["body"]["text"] and "Муфта D105" in second["body"]["text"]
    assert "Муфта D100" not in second["body"]["text"]
    third = handler.handle(callback(payload_by_label(second, "Следующие →")))
    assert "Страница 3" in third["body"]["text"] and "Муфта D111" in third["body"]["text"]
    assert "Следующие →" not in str(third)
    previous = handler.handle(callback(payload_by_label(third, "← Назад")))
    assert previous["body"]["text"] == second["body"]["text"]
    for reply in (first, second, third, previous):
        assert _body(reply["body"]) == reply["body"]
        assert reply["user_id"] == 123456789


def test_contacts_have_their_own_pagination_without_spilling_other_organizations(handler):
    link(handler)
    source = [{"rso": "РСО Тест", "municipality": "Власиха", "person": f"Ответственный {i:02d}",
               "position": "Инженер", "phone": f"000-{i:02d}"} for i in range(12)]
    source.append({"rso": "Другое РСО", "municipality": "Власиха", "person": "ЧУЖОЙ КОНТАКТ", "phone": "99999"})
    handler.stock.contacts_path.write_text(json.dumps({"items": source}), encoding="utf-8")
    first = handler.handle(callback(button(handler.handle(event("муфта")))))
    assert "Ответственный 00" in first["body"]["text"] and "Ответственный 04" in first["body"]["text"]
    assert "Ответственный 05" not in first["body"]["text"]
    second = handler.handle(callback(payload_by_label(first, "Ещё контакты →")))
    assert "Ответственный 05" in second["body"]["text"] and "Ответственный 09" in second["body"]["text"]
    third = handler.handle(callback(payload_by_label(second, "Ещё контакты →")))
    assert "Ответственный 11" in third["body"]["text"] and "Ещё контакты →" not in str(third)
    previous = handler.handle(callback(payload_by_label(third, "← Контакты")))
    assert previous["body"]["text"] == second["body"]["text"]
    for reply in (first, second, third, previous):
        assert "ЧУЖОЙ КОНТАКТ" not in str(reply) and "99999" not in str(reply)
        assert len(reply["body"]["text"]) <= 4000 and _body(reply["body"]) == reply["body"]


def test_contact_buttons_keep_the_specific_item_organization_and_municipality(handler):
    link(handler)
    today = datetime.now(ZoneInfo("Europe/Moscow")).date().isoformat()
    grid = json.loads(handler.stock.published_path.read_text())
    grid["rows"].append(["РСО Тест 2", "Дубна", "Муфта D100", "Фитинги", "Муфта", 6, "шт", "", today])
    handler.stock.published_path.write_text(json.dumps(grid), encoding="utf-8")
    contacts = json.loads(handler.stock.contacts_path.read_text())
    contacts["items"].append({"rso": "РСО Тест 2", "municipality": "Дубна", "person": "Другой инженер", "phone": "000-222"})
    handler.stock.contacts_path.write_text(json.dumps(contacts), encoding="utf-8")
    found = handler.handle(event("муфта D100"))
    rows = found["body"]["attachments"][0]["payload"]["buttons"]
    token = next(row[0]["payload"] for row in rows if "Дубна" in row[0]["text"])
    selected = handler.handle(callback(token))
    assert "Другой инженер" in selected["body"]["text"] and "000-222" in selected["body"]["text"]
    assert "Тестовый инженер" not in selected["body"]["text"] and "+7 000 111" not in selected["body"]["text"]


def test_large_source_fields_and_markup_remain_bounded_plain_text(handler):
    link(handler)
    publish(handler, [f'Муфта <b>ТЕСТ {i}</b> <img src=x onerror=alert(1)> ' + "Ж" * 800 for i in range(6)])
    handler.stock.contacts_path.write_text(json.dumps({"items": [{"rso": "РСО Тест", "municipality": "Власиха",
        "person": "<b>ИМЯ</b>" + "Я" * 250, "position": "<script>1</script>" + "П" * 240,
        "phone": "000-123"}]}), encoding="utf-8")
    found = handler.handle(event("муфта"))
    assert "<b>ТЕСТ" in found["body"]["text"]
    assert "format" not in found["body"] and len(found["body"]["text"]) <= 4000
    _body(found["body"])
    contacted = handler.handle(callback(button(found)))
    assert "<b>ИМЯ</b>" in contacted["body"]["text"] and "format" not in contacted["body"]
    _body(contacted["body"])


@pytest.mark.parametrize("update", [None, [], {"update_type": "message_created", "message": [1]},
    {"update_type": "message_created", "message": {"recipient": [1]}},
    {"update_type": "message_callback", "callback": [1], "message": {"recipient": {"chat_type": "dialog", "chat_id": 1}}},
    {"update_type": "bot_started", "user": [1], "chat_id": 1},
    {"update_type": "bot_started", "user": {"user_id": 123}, "chat_id": 1},
])
def test_malformed_envelopes_do_not_generate_callback_or_identity_keys(handler, update):
    assert handler.event_key(update) is None


def test_invalid_message_body_is_ignored_without_business_or_contact_access(handler):
    link(handler)
    update = event()
    update["message"]["body"] = ["broken"]
    assert handler.handle(update) is None


def test_environment_cannot_redirect_business_source_paths():
    settings = Settings.from_env({"MAX_ZIP_BOT_PUBLISHED_PATH": "/untrusted/other.json", "MAX_ZIP_BOT_ACCOUNTS_PATH": "/untrusted/users.db"})
    assert settings.published_path == ROOT / "data/zip_curator/published.json"
    assert settings.accounts_path == ROOT / "data/dashboard.db"
    assert not settings.enabled and settings.mode == "disabled"


def test_production_mode_not_accepted_and_token_redacted():
    with pytest.raises(ValueError):
        Settings.from_env({"MAX_ZIP_BOT_ENABLED": "true", "MAX_ZIP_BOT_MODE": "production", "MAX_ZIP_BOT_TOKEN": "TEST_TOKEN"})
    settings = Settings.from_env({"MAX_ZIP_BOT_ENABLED": "true", "MAX_ZIP_BOT_MODE": "development-polling", "MAX_ZIP_BOT_TOKEN": "TEST_TOKEN"})
    assert "TEST_TOKEN" not in repr(settings)


def test_missing_source_is_never_rendered_as_empty_inventory(handler):
    link(handler)
    handler.stock.published_path.write_text('{"rows":"corrupt"}')
    response = handler.handle(event("муфта"))
    assert "не готов" in response["body"]["text"] or "недоступ" in response["body"]["text"]
    assert "позиций не найдено" not in response["body"]["text"]
