from services.zip_curator import contacts


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
