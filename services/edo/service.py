from __future__ import annotations

from pathlib import Path

from services.common import DATA_DIR, load_json, save_json, now_str, print_stage


EDO_DIR = DATA_DIR / "edo"
EDO_RESULT_FILE = EDO_DIR / "result.json"


def _build_demo_result_if_missing() -> dict:
    return {
        "created_at": now_str(),
        "summary_message": (
            "Служебная сводка:\n"
            "Проведён контроль заполненности данных.\n"
            "Выявлены записи, требующие актуализации."
        ),
        "public_chat_message": (
            "Коллеги, выполнена проверка заполненности данных. "
            "Просьба обратить внимание на проблемные позиции и актуализировать сведения."
        ),
        "extraction_note": "Источник доступен. Данные обработаны в локальном контуре.",
        "redmine_url": "",
        "screenshot_path": "",
        "screenshot_paths": [],
        "missing_data_issues": [
            {
                "municipality": "Балашиха",
                "organization": "МБУ Тест 1",
                "responsible_name": "Иванов И.И.",
                "responsible_phone": "+7 (999) 111-11-11",
                "message": "Не заполнены отдельные поля по входящим и исходящим документам."
            }
        ],
        "personal_messages": [
            {
                "municipality": "Балашиха",
                "organization": "МБУ Тест 1",
                "responsible_name": "Иванов И.И.",
                "responsible_phone": "+7 (999) 111-11-11",
                "status": "risk",
                "is_edited": False,
                "message": (
                    "Добрый день.\n"
                    "По результатам проверки выявлены незаполненные данные. "
                    "Просьба актуализировать сведения."
                )
            }
        ],
        "rows": [
            {
                "municipality": "Балашиха",
                "organization": "МБУ Тест 1",
                "total_docs_vn": 100,
                "electronic_docs_vn": 90,
                "total_docs_vh": 80,
                "electronic_docs_vh": 70,
                "total_docs_ish": 60,
                "electronic_docs_ish": 55,
                "responsible_name": "Иванов И.И.",
                "responsible_phone": "+7 (999) 111-11-11",
                "status": "risk",
                "reason": "Есть незаполненные показатели"
            },
            {
                "municipality": "Химки",
                "organization": "МБУ Тест 2",
                "total_docs_vn": 120,
                "electronic_docs_vn": 120,
                "total_docs_vh": 95,
                "electronic_docs_vh": 95,
                "total_docs_ish": 77,
                "electronic_docs_ish": 77,
                "responsible_name": "Петров П.П.",
                "responsible_phone": "+7 (999) 222-22-22",
                "status": "ok",
                "reason": "Замечаний нет"
            }
        ]
    }


def run_edo_check() -> dict:
    print_stage("Подготовка", "Инициализация сценария EDO...")
    print_stage("Загрузка данных", "Чтение текущего результата EDO...")

    result = load_json(EDO_RESULT_FILE)

    if not result:
        print_stage("Формирование демо-результата", "Файл результата не найден, создаётся демонстрационный результат.")
        result = _build_demo_result_if_missing()
    else:
        print_stage("Обновление результата", "Результат найден, обновляется метка времени.")
        result["created_at"] = now_str()

        for item in result.get("personal_messages", []) or []:
            if "is_edited" not in item:
                item["is_edited"] = False

    print_stage("Сохранение", "Сохранение результата EDO...")
    save_json(EDO_RESULT_FILE, result)

    print_stage("Завершение", "Проверка EDO завершена.")
    return result