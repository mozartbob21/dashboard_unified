from services.zip_curator import contacts
import json
import tempfile
from pathlib import Path
from unittest.mock import patch


def test_wide_rso_table_creates_leader_and_chief_engineer_contacts():
    rows = [
        [
            "ОМСУ",
            "Наименование   РСО",
            "ФИО   руководителя",
            "Моб.   тел. руководителя",
            "ФИО   главного инженера",
            "Моб.   тел. главного инженера",
        ],
        [
            "Одинцовский",
            "АО Водоканал",
            "Иванов Иван Иванович",
            "+7 999 111-22-33",
            "Петров Пётр Петрович",
            "+7 999 444-55-66",
        ],
    ]

    parsed = contacts._rows_to_contacts(rows)

    assert parsed == [
        {
            "municipality": "Одинцовский",
            "rso": "АО Водоканал",
            "person": "Иванов Иван Иванович",
            "phone": "+7 999 111-22-33",
            "position": "Руководитель",
        },
        {
            "municipality": "Одинцовский",
            "rso": "АО Водоканал",
            "person": "Петров Пётр Петрович",
            "phone": "+7 999 444-55-66",
            "position": "Главный инженер",
        },
    ]


def test_header_matching_supports_future_wording_without_substring_collisions():
    rows = [
        [
            "Городской округ",
            "Организация",
            "Ф.И.О. директора",
            "Телефон директора",
            "Главный инженер — ФИО",
            "Тел. гл. инженера",
        ],
        ["Истра", "РСО 2", "Сидоров С.С.", "8 900 100-20-30", "Орлов О.О.", "8 900 400-50-60"],
    ]

    parsed = contacts._rows_to_contacts(rows)

    assert [(item["position"], item["person"], item["phone"]) for item in parsed] == [
        ("Руководитель", "Сидоров С.С.", "8 900 100-20-30"),
        ("Главный инженер", "Орлов О.О.", "8 900 400-50-60"),
    ]
    assert contacts._field_for_header("ФИО руководителя") == "person"
    assert contacts._field_for_header("Моб. тел. руководителя") == "phone"


def test_contacts_with_different_roles_do_not_merge_on_shared_phone():
    leader = {
        "rso": "РСО 3",
        "person": "Иванов И.И.",
        "phone": "+7 495 000-00-00",
        "position": "Руководитель",
    }
    engineer = {
        "rso": "РСО 3",
        "person": "Петров П.П.",
        "phone": "+7 495 000-00-00",
        "position": "Главный инженер",
    }

    assert contacts._same_contact(leader, engineer) is False


def test_reimport_same_file_recovers_both_legacy_names_and_phones_without_duplicates():
    headers = ["ОМСУ", "Наименование РСО", "ФИО руководителя", "Моб. тел. руководителя",
               "ФИО главного инженера", "Моб. тел. главного инженера"]
    content = (";".join(headers) + "\nИстра;РСО;Иванов Иван Иванович;+7 999 111-22-33;"
               "Петров Петр Петрович;+7 999 444-55-66\n").encode()
    original = {"items": [{"id": "old", "rso": "РСО", "municipality": "Истра",
                           "person": "Петров Петр Петрович", "phone": "Иванов Иван Иванович"}]}
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "contacts.json"
        path.write_text(json.dumps(original), encoding="utf-8")
        with patch.object(contacts, "CONTACTS_FILE", path):
            old_view = contacts.load_contacts()
            assert old_view["needs_reimport"]
            assert len(old_view["items"]) == 2
            assert all(not item["phone"] for item in old_view["items"])
            for _ in range(2):
                result = contacts.import_contacts("same.csv", content)
                assert result["total"] == 2
                assert {item["person"]: item["phone"] for item in result["payload"]["items"]} == {
                    "Иванов Иван Иванович": "+7 999 111-22-33",
                    "Петров Петр Петрович": "+7 999 444-55-66",
                }
            assert not contacts.load_contacts()["needs_reimport"]
            assert json.loads(path.with_name(path.name + ".before-import.bak").read_text()) == original


def test_repeated_director_label_moves_to_position_and_matches_reimport():
    original = {"rso": "РСО", "person": "Директор филиала Иванов Иван Иванович "
                "Директор филиала Иванов Иван Иванович", "phone": "", "position": "Ответственный"}
    clean = contacts._normalize_contact(original)
    assert clean["person"] == "Иванов Иван Иванович"
    assert clean["position"] == "Директор филиала"
    assert contacts._same_contact(clean, {"rso": "РСО", "person": clean["person"], "position": "Руководитель"})


def test_migration_preserves_distinct_people_and_short_numbers():
    people = "Директор филиала Иванов Иван Иванович Директор филиала Петров Петр Петрович"
    clean = contacts._normalize_contact({"person": people})
    assert "Иванов" in clean["person"] and "Петров" in clean["person"]
    assert not contacts._name_in_phone("1234")
    assert not contacts._name_in_phone("доб. 123")
    assert not contacts._same_contact(
        {"rso": "РСО", "person": "Иванов", "phone": "12345"},
        {"rso": "РСО", "person": "Петров", "phone": "12345"},
    )
