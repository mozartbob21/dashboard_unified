"""Read-only MAX stock contract, using synthetic publications and contacts only."""
import json
from datetime import date, timedelta
from pathlib import Path

import pytest

from services.max_bot import stock
from services import report_municipalities


TODAY = date(2026, 10, 22)
HEADERS = ["РСО", "Округ", "Наименование", "Категория", "Группа", "Количество", "Ед.", "Водоподготовка", "Дата"]


def row(name="Муфта D100", municipality="Тестовый", organization="РСО Тест", quantity=4, unit="шт", day="22.10.2026"):
    return [organization, municipality, name, "Фитинги", "Муфта", quantity, unit, "", day]


def write(path, payload):
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


@pytest.fixture
def store(tmp_path, monkeypatch):
    monkeypatch.setattr(report_municipalities, "REGISTRY", tmp_path / "municipalities.json")
    published, contacts = tmp_path / "published.json", tmp_path / "contacts.json"
    write(published, {"rows": [HEADERS, row()], "published_at": "2026-10-22T12:00:00",
                      "rso_metadata": {"secret": {"source_path": "/private/source.xlsx", "source_identity": "not-public"}}})
    write(contacts, {"items": [{"id": "existing-id", "rso": "РСО Тест", "municipality": "Тестовый",
                                "person": "Тестовый инженер", "position": "Главный инженер", "phone": "+7 000 111-22-33",
                                "email": "not-exported@example.invalid", "note": "PRIVATE_NOTE"}],
                     "updated_at": "2026-10-21T11:00:00", "source_file": "private-directory.xlsx"})
    return stock.StockService(published, contacts)


def publish(service, rows, **extra):
    write(service.published_path, {"rows": [HEADERS, *rows], "published_at": "2026-10-22T12:00:00", **extra})


def search(service, query="муфта", **kwargs):
    return service.search(query, today=TODAY, **kwargs)


def items(service, query="муфта", **kwargs):
    return search(service, query, **kwargs)["queries"][0]["items"]


def details(service, item=None):
    item = item or items(service)[0]
    return service.contacts(item["id"], item["municipality_id"])


def test_projection_contains_only_public_fields_and_never_reads_contacts(store):
    store.contacts_path.unlink()
    result = search(store, "Нужна муфта D100")
    item = result["queries"][0]["items"][0]
    assert result["queries"][0]["status"] == "ok"
    assert item["match_type"] == "exact"
    assert item["availability"] == "available"
    assert item["stock_date"] == "2026-10-22"
    assert item["quantity"] == 4
    assert result["policy"] == {"stale_days": 30}
    assert set(item) == {"id", "name", "category", "group", "quantity", "unit", "municipality_id", "municipality",
                         "organization", "stock_date", "availability", "availability_label", "match_type"}
    serialized = json.dumps(result)
    assert "source_path" not in serialized and "source_identity" not in serialized
    assert "/private" not in serialized and "phone" not in serialized and "contacts" not in serialized


@pytest.mark.parametrize(("quantity", "day", "unit", "expected"), [
    (4, "22.10.2026", "шт", "available"), (4, "22.09.2026", "шт", "available"),
    (4, "21.09.2026", "шт", "stale"), (0, "22.10.2026", "шт", "zero"),
    (None, "22.10.2026", "шт", "unknown"), ("", "22.10.2026", "шт", "unknown"),
    (-1, "22.10.2026", "шт", "unknown"), (True, "22.10.2026", "шт", "unknown"),
    ("1e99", "22.10.2026", "шт", "unknown"), ("NaN", "22.10.2026", "шт", "unknown"),
    ("3,25", "2026-10-22", "пог. м", "available"), (4, "", "шт", "unknown"),
    (4, "23.10.2026", "шт", "unknown"), (4, "31.02.2026", "шт", "unknown"),
    (4, "22.10.2026", "", "unknown"), (4, "22.10.2026", "—", "unknown"),
    (4, "22.10.2026", "не указано", "unknown"),
])
def test_freshness_quantity_and_date_are_not_invented(store, quantity, day, unit, expected):
    publish(store, [row(quantity=quantity, day=day, unit=unit)])
    item = items(store)[0]
    assert item["availability"] == expected
    if expected != "available":
        assert not item["availability_label"].startswith("Есть")


def test_publication_or_source_date_does_not_refresh_an_old_or_unknown_report(store):
    publish(store, [row(day="01.01.2020"), row(name="Муфта D50", day="")],
            rso_metadata={"рсо тест": {"source_created_at": "2026-10-22T12:00:00", "loaded_at": "2026-10-22T12:00:00"}})
    assert {item["availability"] for item in items(store)} == {"stale", "unknown"}


@pytest.mark.parametrize("municipality", ["", "—", "не указан", "Московская область"])
def test_missing_municipality_cannot_confirm_stock_or_expose_contacts(store, municipality):
    publish(store, [row(municipality=municipality)])
    item = items(store)[0]
    assert item["availability"] == "unknown"
    with pytest.raises(stock.StockError) as error:
        details(store, item)
    assert error.value.status_code == 404


def test_preserves_numeric_boundaries_model_order_and_dimensions(store):
    publish(store, [row("Муфта D100"), row("Муфта D1000"), row("Муфта D108"),
                    row("Насос K50-32-125"), row("Насос K32-50-125"),
                    row("Переход 100×50"), row("Переход 50×100")])
    assert [r["name"] for r in items(store, "муфта 100")] == ["Муфта D100"]
    assert [r["name"] for r in items(store, "Нужен насос K50-32-125")] == ["Насос K50-32-125"]
    assert [r["name"] for r in items(store, "переход 100х50")] == ["Переход 100×50"]
    assert not items(store, "муфта D101")
    assert not items(store, "насос K50-32-126")


def test_normalizes_yo_whitespace_inflection_but_labels_partial_match(store):
    publish(store, [row("Муфта разъёмная D100")])
    found = items(store, "  Где   есть  муфту   разъемная 100  ")
    assert len(found) == 1 and found[0]["match_type"] == "terms"


def test_typo_suggestions_are_marked_and_numbers_are_not_fuzzy(store):
    result = search(store, "мфута D100")["queries"][0]
    assert result["status"] == "clarify"
    assert result["items"][0]["match_type"] == "similar"
    assert not items(store, "мфута D101")


def test_model_word_typo_not_treated_as_exact(store):
    publish(store, [row("Насос Grundfos 100")])
    item = items(store, "насос Grundfoz 100")[0]
    assert item["match_type"] == "similar"


def test_multiple_variants_require_explicit_selection(store):
    publish(store, [row("Муфта D100"), row("Муфта D108")])
    result = search(store)["queries"][0]
    assert result["status"] == "clarify" and len(result["items"]) == 2
    assert result["warnings"]


def test_same_position_in_multiple_municipalities_keeps_individual_values(store):
    publish(store, [row(municipality="Первый", quantity=4), row(municipality="Второй", quantity=9)])
    result = search(store)["queries"][0]
    assert result["status"] == "ok"
    assert {r["quantity"] for r in result["items"]} == {4, 9}
    assert len({r["municipality_id"] for r in result["items"]}) == 2


def test_multiple_positions_have_independent_paginated_results_not_a_bundle(store):
    publish(store, [row("Муфта D100", municipality="Первый"), row("Муфта D100", municipality="Второй"),
                    row("Насос K50", municipality="Третий")])
    result = search(store, "муфта; насос\nкран", page_size=1)
    assert len(result["queries"]) == 3
    assert result["queries"][0]["pagination"] == {"page": 1, "page_size": 1, "total": 2, "has_next": True}
    assert result["queries"][1]["pagination"]["total"] == 1
    assert result["queries"][2]["status"] == "empty"
    assert "комплекта" in result["warnings"][0]
    assert len(items(store, "муфта", page=2, page_size=1)) == 1
    assert items(store, "муфта", page=3, page_size=1) == []


@pytest.mark.parametrize("query", [None, {}, "", "где есть", "x", "а" * 241, "а" * 1201,
                                  "муфта\x00", ";".join(["муфта"] * 6), " ".join(["муфта"] * 25)])
def test_invalid_queries_are_bounded(store, query):
    with pytest.raises(stock.StockError) as error:
        search(store, query)
    assert error.value.status_code == 400


@pytest.mark.parametrize("kwargs", [{"page": 0}, {"page": -1}, {"page": True}, {"page": 1.5},
                                    {"page_size": 0}, {"page_size": 51}, {"page_size": "20"}])
def test_pagination_validation(store, kwargs):
    with pytest.raises(stock.StockError) as error:
        search(store, **kwargs)
    assert error.value.status_code == 400


def test_sql_and_html_are_plain_search_text(store):
    assert not items(store, "' OR 1=1 --")
    assert not items(store, "<script>alert(1)</script>")
    publish(store, [row("Муфта <img src=x onerror=alert(1)>")])
    item = items(store)[0]
    assert item["name"] == "Муфта <img src=x onerror=alert(1)>"
    assert set(item) >= {"name", "availability"}  # JSON text, frontend must use textContent


@pytest.mark.parametrize("raw", ["", "[]", "{broken", '{"rows":[],"rows":[]}',
                                '{"rows": NaN}', '{"rows": null}', '{"rows": [["wrong"]]}'])
def test_invalid_or_unpublished_snapshot_is_service_error_not_empty_result(store, raw):
    store.published_path.write_text(raw)
    with pytest.raises(stock.StockError) as error:
        search(store)
    assert error.value.code == "STOCK_UNAVAILABLE" and error.value.status_code == 503
    assert str(store.published_path) not in str(error.value)


def test_missing_snapshot_or_contact_is_service_error_without_path(store):
    item = items(store)[0]
    store.contacts_path.unlink()
    with pytest.raises(stock.StockError) as error:
        details(store, item)
    assert error.value.code == "CONTACTS_NOT_CONFIGURED" and error.value.status_code == 409
    assert "contacts.json" not in str(error.value)
    store.published_path.unlink()
    with pytest.raises(stock.StockError) as error:
        search(store)
    assert error.value.code == "STOCK_UNAVAILABLE"


def test_empty_valid_publication_is_legitimate_no_result(store):
    publish(store, [])
    assert search(store)["queries"][0]["status"] == "empty"


def test_oversized_snapshot_is_bounded(store, monkeypatch):
    monkeypatch.setattr(stock, "MAX_STOCK_BYTES", 10)
    with pytest.raises(stock.StockError):
        search(store)


def test_rejects_duplicate_or_missing_required_headers(store):
    write(store.published_path, {"rows": [[*HEADERS, "РСО"], [*row(), "Another"]]})
    with pytest.raises(stock.StockError):
        search(store)
    write(store.published_path, {"rows": [["Наименование", "Количество"], ["Муфта", 4]]})
    with pytest.raises(stock.StockError):
        search(store)


def test_omitted_invalid_rows_are_reported_but_never_zero(store):
    publish(store, [row(), [], {"not": "row"}, row(name="")])
    result = search(store)
    assert len(result["queries"][0]["items"]) == 1
    assert result["warnings"]
    publish(store, [[], row(name="")])
    with pytest.raises(stock.StockError):
        search(store)


def test_identical_copies_are_not_double_counted_and_different_units_not_summed(store):
    publish(store, [row(), row(), row(unit="пог. м")])
    found = items(store)
    assert len(found) == 2
    assert {r["unit"] for r in found} == {"шт", "пог. м"}
    assert {r["quantity"] for r in found} == {4}


def test_contacts_are_exactly_scoped_and_safe_projection(store):
    original = json.loads(store.contacts_path.read_text())
    original["items"].extend([
        {"rso": "Другое РСО", "municipality": "Тестовый", "phone": "000-00", "person": "WRONG_ORG"},
        {"rso": "РСО Тест", "municipality": "Другой", "phone": "111-11", "person": "WRONG_MUNI"},
        {"rso": "РСО Тест", "municipality": "Тестовый", "phone": "", "person": "Тестовый директор", "position": "Директор"},
    ])
    write(store.contacts_path, original)
    result = details(store)
    assert len(result["contacts"]) == 2
    assert any(c["phone"] is None for c in result["contacts"])
    assert set(result["contacts"][0]) == {"id", "person", "position", "organization", "phone"}
    serialized = json.dumps(result)
    for secret in ("WRONG_ORG", "WRONG_MUNI", "PRIVATE_NOTE", "not-exported", "private-directory", "existing-id"):
        assert secret not in serialized


def test_missing_contact_has_explicit_message_not_fake_number(store):
    write(store.contacts_path, {"items": []})
    result = details(store)
    assert result["contacts"] == [] and result["warnings"]


def test_contact_without_municipality_requires_unique_published_organization(store):
    write(store.contacts_path, {"items": [{"rso": "РСО Тест", "phone": "12345", "person": "Тест"}]})
    assert details(store)["contacts"][0]["phone"] == "12345"
    publish(store, [row(), row(municipality="Второй")])
    result = details(store)
    assert result["contacts"] == [] and any("принадлежность" in w for w in result["warnings"])
    publish(store, [row(), row(municipality="")])
    assert details(store, next(r for r in items(store) if r["municipality"] == "Тестовый"))["contacts"] == []


def test_shared_phone_preserves_distinct_responsibles(store):
    write(store.contacts_path, {"items": [{"rso": "РСО Тест", "municipality": "Тестовый", "phone": "12345", "person": person}
                                         for person in ("Тест Один", "Тест Два")]})
    assert len(details(store)["contacts"]) == 2


def test_org_matching_does_not_ignore_legal_form_numbers_or_use_fuzzy_names(store):
    write(store.contacts_path, {"items": [{"rso": "РСО Тест2", "municipality": "Тестовый", "phone": "12345", "person": "Не тот"},
                                         {"rso": "МУП РСО Тест", "municipality": "Тестовый", "phone": "56789", "person": "Не тот"}]})
    assert details(store)["contacts"] == []


def test_municipality_aliases_use_existing_registry_not_fuzzy_guess(store):
    write(report_municipalities.REGISTRY, {"municipalities": {"Тестовый": {"aliases": ["Тестовый округ"]}}})
    write(store.contacts_path, {"items": [{"rso": "РСО Тест", "municipality": "Тестовый округ", "phone": "12345"}]})
    assert details(store)["contacts"][0]["phone"] == "12345"


@pytest.mark.parametrize("phone", ["Иванов Иван", "<img onerror=alert(1)>", "name1@example.invalid", ""])
def test_legacy_name_or_markup_is_never_a_phone(store, phone):
    write(store.contacts_path, {"items": [{"rso": "РСО Тест", "municipality": "Тестовый", "phone": phone}]})
    assert details(store)["contacts"][0]["phone"] is None


@pytest.mark.parametrize("payload", [{"items": {}}, {"items": [None]}, {"items": "bad"}, {}])
def test_malformed_contact_data_is_service_error(store, payload):
    write(store.contacts_path, payload)
    with pytest.raises(stock.StockError) as error:
        details(store)
    assert error.value.code == "CONTACTS_UNAVAILABLE"


def test_contacts_limit_and_missing_phone_warning(store):
    write(store.contacts_path, {"items": [{"rso": "РСО Тест", "municipality": "Тестовый", "person": f"Тестовый {i}"} for i in range(55)]})
    result = details(store)
    assert len(result["contacts"]) == 50 and any("50" in warning for warning in result["warnings"])


@pytest.mark.parametrize(("item_id", "muni_id"), [("../state.json", "x"), (None, "x"), ("i_" + "a" * 32, "m_" + "a" * 32)])
def test_fake_or_path_ids_are_rejected(store, item_id, muni_id):
    with pytest.raises(stock.StockError) as error:
        store.contacts(item_id, muni_id)
    assert error.value.status_code == 404


def test_municipality_cannot_be_swapped_between_existing_items(store):
    publish(store, [row(municipality="Первый"), row(municipality="Второй")])
    first, second = items(store)
    with pytest.raises(stock.StockError) as error:
        store.contacts(first["id"], second["municipality_id"])
    assert error.value.status_code == 404


def test_removed_or_changed_result_id_is_revalidated_on_contact_lookup(store):
    item = items(store)[0]
    publish(store, [row(quantity=0)])
    with pytest.raises(stock.StockError) as error:
        details(store, item)
    assert error.value.status_code == 404
    current = items(store)[0]
    assert current["availability"] == "zero" and current["id"] != item["id"]
    publish(store, [])
    with pytest.raises(stock.StockError):
        details(store, current)


def test_cache_never_hides_source_corruption_or_deletion(store):
    assert items(store)
    store.published_path.write_text("broken")
    with pytest.raises(stock.StockError):
        search(store)
    store.published_path.unlink()
    with pytest.raises(stock.StockError):
        search(store)


def test_read_only_operations_do_not_mutate_existing_files(store):
    before = {p: (p.read_bytes(), p.stat().st_mtime_ns) for p in (store.published_path, store.contacts_path)}
    search(store)
    details(store)
    assert {p: (p.read_bytes(), p.stat().st_mtime_ns) for p in before} == before
    assert {p.name for p in store.published_path.parent.iterdir()} == {"published.json", "contacts.json"}


def test_ten_thousand_synthetic_rows_correct_pagination_and_freshness(store):
    publish(store, [row(f"Муфта D{i}", municipality=f"Округ {i % 50}", quantity=i, day="22.10.2026") for i in range(10_000)])
    result = search(store, "муфта", page=200, page_size=50)["queries"][0]
    assert result["pagination"] == {"page": 200, "page_size": 50, "total": 10_000, "has_next": False}
    assert len(result["items"]) == 50
    assert items(store, "Муфта D9999")[0]["quantity"] == 9999


def test_surrogate_in_source_is_sanitized_and_query_is_rejected(store):
    store.published_path.write_text(json.dumps({"rows": [HEADERS, row("Муфта D100\ud800")]}), encoding="utf-8")
    assert items(store)[0]["name"] == "Муфта D100"
    with pytest.raises(stock.StockError) as error:
        search(store, "Муфта\ud800")
    assert error.value.status_code == 400
