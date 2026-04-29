import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError, sync_playwright


BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data" / "watercontrol"
SCREENSHOTS_DIR = DATA_DIR / "screenshots"
RESULT_FILE = DATA_DIR / "result.json"

RESPONSIBLES_CANDIDATES = [
    DATA_DIR / "responsible_by_municipalyty.json",
    DATA_DIR / "responsibles_by_municipality.json",
    BASE_DIR / "data" / "responsibles_by_municipality.json",
]

WATERCONTROL_URL = os.getenv(
    "WATERCONTROL_URL",
    os.getenv("WATERCONTROL_DASHBOARD_URL", "https://datalens.yandex/j9dqqujx03qa3"),
)

HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
DEBUG_ROWS = os.getenv("DEBUG_ROWS", "false").lower() == "true"

SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
RESULT_FILE.parent.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)

EMPTY_MARKERS = {"", "null", "none", "-", "—", "n/a", "nan", "не заполнено"}

REQUIRED_FIELDS_ALIASES = {
    "Дата плановой промывки": [
        "Дата плановой промывки",
        "Плановая промывка",
        "Дата промывки",
        "date_planned_flush",
        "planned_flush_date",
    ],
    "Ссылка на акт": [
        "Ссылка на акт",
        "Акт",
        "Ссылка",
        "URL акта",
        "act_link",
        "act_url",
    ],
}

COLUMN_ALIASES = {
    "municipality": [
        "Муниципалитет",
        "municipality",
        "Городской округ",
        "Округ",
    ],
    "organization": [
        "Организация",
        "Объект",
        "organization",
        "Наименование",
    ],
    "address": [
        "Адрес",
        "Адрес объекта",
        "address",
        "Местоположение",
    ],
    "task_id": [
        "ID задачи",
        "Номер задачи",
        "ID",
        "№",
        "task_id",
        "id",
    ],
    "responsible_name": [
        "Ответственный",
        "ФИО",
        "Исполнитель",
        "responsible_name",
    ],
    "responsible_phone": [
        "Телефон",
        "Контактный телефон",
        "responsible_phone",
    ],
}

COMMON_HEADER_HINTS = []
for aliases in COLUMN_ALIASES.values():
    COMMON_HEADER_HINTS.extend(aliases)
for aliases in REQUIRED_FIELDS_ALIASES.values():
    COMMON_HEADER_HINTS.extend(aliases)


def emit_stage(text: str):
    print(f"STAGE: {text}", flush=True)


def normalize_text(value):
    return (value or "").replace("\xa0", " ").strip()


def normalize_check(value):
    return normalize_text(value).lower()


def normalize_municipality_name(value):
    return normalize_text(value).upper()


def normalize_org_name(value):
    return normalize_text(value).lower()


def is_empty_value(value):
    return normalize_check(value) in EMPTY_MARKERS


def get_responsibles_file():
    for path in RESPONSIBLES_CANDIDATES:
        if path.exists():
            print(f"RESPONSIBLES FILE FOUND: {path}", flush=True)
            return path
    print("RESPONSIBLES FILE NOT FOUND", flush=True)
    return None


def load_responsibles_by_municipality():
    file_path = get_responsibles_file()
    if not file_path:
        return {}

    try:
        raw = json.loads(file_path.read_text(encoding="utf-8"))
        result = {}

        for municipality, item in raw.items():
            key = normalize_municipality_name(municipality)

            if isinstance(item, dict):
                result[key] = {
                    "name": normalize_text(item.get("name", "Ответственный не указан")) or "Ответственный не указан",
                    "phone": normalize_text(item.get("phone", "")),
                }
            elif isinstance(item, str):
                result[key] = {
                    "name": normalize_text(item) or "Ответственный не указан",
                    "phone": "",
                }

        print("RESPONSIBLES_MAP SIZE:", len(result), flush=True)
        return result

    except Exception as e:
        print("ERROR LOADING RESPONSIBLES:", str(e), flush=True)
        return {}


def get_responsible_info_by_municipality(municipality, responsibles_map):
    key = normalize_municipality_name(municipality)
    item = responsibles_map.get(key)

    if item:
        return {
            "name": item.get("name", "Ответственный не указан"),
            "phone": item.get("phone", ""),
        }

    return {
        "name": "Ответственный не указан",
        "phone": "",
    }


def save_result(payload):
    RESULT_FILE.parent.mkdir(parents=True, exist_ok=True)
    RESULT_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print("RESULT SAVED:", RESULT_FILE, flush=True)


def save_error_result(timestamp, screenshot_paths, message):
    first_screenshot = screenshot_paths[0] if screenshot_paths else ""

    result = {
        "created_at": timestamp,
        "summary_message": "Проверка не завершена.",
        "public_chat_message": "",
        "personal_messages": [],
        "rows": [],
        "screenshot_path": first_screenshot,
        "screenshot_paths": screenshot_paths,
        "source_url": WATERCONTROL_URL,
        "extraction_note": message,
    }
    save_result(result)


def is_captcha_page(page):
    try:
        content = page.content().lower()
    except Exception:
        return False

    markers = [
        "i'm not a robot",
        "smartcaptcha",
        "please confirm that you and not a robot",
        "yandex cloud",
        "press to continue",
        "captcha",
    ]
    return any(marker in content for marker in markers)


def wait_for_manual_captcha_pass(page, max_wait_seconds=180):
    print("Обнаружена капча. Пройдите её вручную в открытом браузере...", flush=True)
    start_time = time.time()

    while time.time() - start_time < max_wait_seconds:
        page.wait_for_timeout(2000)

        try:
            if not is_captcha_page(page):
                page.wait_for_load_state("domcontentloaded", timeout=30000)
                page.wait_for_timeout(5000)
                print("Капча пройдена, продолжаем обработку.", flush=True)
                return True
        except Exception:
            pass

    return False


def find_scrollable_container(page):
    selectors = [
        '[role="grid"]',
        'table',
        '.table',
        '.grid-container',
        '.scrollable',
        '[class*="scroll"]',
        '[class*="table"]',
    ]

    for selector in selectors:
        locator = page.locator(selector)

        try:
            count = locator.count()
        except Exception:
            count = 0

        for i in range(min(count, 15)):
            candidate = locator.nth(i)
            handle = candidate.element_handle()

            if handle is None:
                continue

            try:
                metrics = page.evaluate(
                    """
                    (el) => ({
                        scrollHeight: el.scrollHeight,
                        clientHeight: el.clientHeight,
                        scrollWidth: el.scrollWidth,
                        clientWidth: el.clientWidth,
                        overflowY: window.getComputedStyle(el).overflowY
                    })
                    """,
                    handle,
                )

                if metrics and metrics["scrollHeight"] > metrics["clientHeight"] + 80:
                    return candidate, handle
            except Exception:
                continue

    return None, None


def save_table_screenshots(page, base_name="table_part", max_parts=8):
    scrollable_locator, scrollable_handle = find_scrollable_container(page)

    if scrollable_locator is None or scrollable_handle is None:
        file_name = f"{base_name}_1.png"
        file_path = SCREENSHOTS_DIR / file_name
        page.screenshot(path=str(file_path), full_page=True)
        return [f"data/watercontrol/screenshots/{file_name}"]

    try:
        scrollable_locator.click(timeout=3000)
    except Exception:
        pass

    try:
        page.evaluate("(el) => { el.scrollTop = 0; }", scrollable_handle)
        page.wait_for_timeout(1500)
    except Exception:
        pass

    saved_files = []
    previous_scroll_top = -1

    for i in range(max_parts):
        file_name = f"{base_name}_{i + 1}.png"
        file_path = SCREENSHOTS_DIR / file_name

        try:
            scrollable_locator.screenshot(path=str(file_path))
        except Exception:
            page.screenshot(path=str(file_path), full_page=True)

        saved_files.append(f"data/watercontrol/screenshots/{file_name}")

        try:
            current_scroll_top = page.evaluate("(el) => el.scrollTop", scrollable_handle)
            client_height = page.evaluate("(el) => el.clientHeight", scrollable_handle)
            scroll_height = page.evaluate("(el) => el.scrollHeight", scrollable_handle)

            page.evaluate(
                """
                (el) => {
                    const step = Math.max(el.clientHeight - 120, 200);
                    el.scrollTop = el.scrollTop + step;
                }
                """,
                scrollable_handle,
            )

            page.wait_for_timeout(1800)

            new_scroll_top = page.evaluate("(el) => el.scrollTop", scrollable_handle)

            if new_scroll_top == current_scroll_top:
                try:
                    page.mouse.wheel(0, 900)
                    page.wait_for_timeout(1800)
                    new_scroll_top = page.evaluate("(el) => el.scrollTop", scrollable_handle)
                except Exception:
                    pass

            if new_scroll_top == current_scroll_top:
                try:
                    page.keyboard.press("PageDown")
                    page.wait_for_timeout(1800)
                    new_scroll_top = page.evaluate("(el) => el.scrollTop", scrollable_handle)
                except Exception:
                    pass

            if new_scroll_top == current_scroll_top or new_scroll_top == previous_scroll_top:
                break

            if new_scroll_top + client_height >= scroll_height:
                final_name = f"{base_name}_{i + 2}.png"
                final_path = SCREENSHOTS_DIR / final_name
                try:
                    scrollable_locator.screenshot(path=str(final_path))
                except Exception:
                    page.screenshot(path=str(final_path), full_page=True)
                saved_files.append(f"data/watercontrol/screenshots/{final_name}")
                break

            previous_scroll_top = current_scroll_top

        except Exception:
            break

    unique_files = []
    seen = set()
    for item in saved_files:
        if item not in seen:
            seen.add(item)
            unique_files.append(item)

    return unique_files


def extract_rows_from_html_table(page):
    table_data = []

    tables = page.locator("table")
    try:
        tables_count = tables.count()
    except Exception:
        tables_count = 0

    for t in range(min(tables_count, 10)):
        table = tables.nth(t)

        headers = []
        header_cells = table.locator("thead tr th")
        try:
            header_count = header_cells.count()
        except Exception:
            header_count = 0

        for i in range(header_count):
            try:
                headers.append(normalize_text(header_cells.nth(i).inner_text()))
            except Exception:
                headers.append("")

        rows = table.locator("tbody tr")
        try:
            rows_count = rows.count()
        except Exception:
            rows_count = 0

        table_rows = []
        for i in range(rows_count):
            row = rows.nth(i)
            cells = row.locator("td")

            values = []
            try:
                cells_count = cells.count()
            except Exception:
                cells_count = 0

            for j in range(cells_count):
                try:
                    values.append(normalize_text(cells.nth(j).inner_text()))
                except Exception:
                    values.append("")

            if values:
                table_rows.append(values)

        if headers and table_rows:
            table_data.append({"headers": headers, "rows": table_rows})

    return table_data


def extract_rows_from_generic_grid(page):
    rows_data = []

    selectors = [
        '[role="row"]',
        '.table__row',
        '.grid-row',
        '.datalens-table__row',
        '[role="grid"] [role="row"]',
        'tr',
    ]

    for selector in selectors:
        rows = page.locator(selector)

        try:
            rows_count = rows.count()
        except Exception:
            rows_count = 0

        current = []
        for i in range(min(rows_count, 500)):
            row = rows.nth(i)
            try:
                text = normalize_text(row.inner_text())
            except Exception:
                continue

            if not text:
                continue

            parts = [p.strip() for p in text.split("\n") if normalize_text(p)]

            if len(parts) >= 4:
                current.append(parts)

        if current:
            rows_data.extend(current)
            return rows_data

    return rows_data


def looks_like_target_headers(headers):
    normalized_headers = [normalize_check(h) for h in headers if normalize_text(h)]
    if not normalized_headers:
        return False

    hints = [normalize_check(x) for x in COMMON_HEADER_HINTS]
    matched = sum(1 for h in normalized_headers if h in hints)

    return matched >= 3


def cleanup_table_rows(headers, rows):
    cleaned = []
    seen = set()

    normalized_headers = [normalize_text(h) for h in headers]

    for values in rows:
        normalized_values = [normalize_text(v) for v in values]

        if len(normalized_values) < 2:
            continue

        while len(normalized_values) < len(normalized_headers):
            normalized_values.append("")

        normalized_values = normalized_values[:len(normalized_headers)]

        key = tuple(normalized_values)
        if key in seen:
            continue

        seen.add(key)
        cleaned.append(dict(zip(normalized_headers, normalized_values)))

    return cleaned


def map_row_by_aliases(row_dict, aliases_map):
    result = {}

    lowered = {normalize_check(k): v for k, v in row_dict.items()}

    for target_field, aliases in aliases_map.items():
        value = ""

        for alias in aliases:
            alias_key = normalize_check(alias)

            if alias in row_dict:
                value = normalize_text(row_dict.get(alias))
                break

            if alias_key in lowered:
                value = normalize_text(lowered[alias_key])
                break

        result[target_field] = value

    return result


def find_missing_fields(row_dict):
    missing_fields = []

    lowered = {normalize_check(k): v for k, v in row_dict.items()}

    for pretty_name, aliases in REQUIRED_FIELDS_ALIASES.items():
        found_value = ""

        for alias in aliases:
            alias_key = normalize_check(alias)
            if alias in row_dict:
                found_value = normalize_text(row_dict.get(alias))
                break
            if alias_key in lowered:
                found_value = normalize_text(lowered[alias_key])
                break

        if is_empty_value(found_value):
            missing_fields.append(pretty_name)

    return missing_fields


def build_public_chat_message(issues):
    if not issues:
        return "По результатам автоматической проверки WaterControl незаполненных обязательных полей не выявлено."

    lines = [
        "В ходе автоматической проверки WaterControl выявлены объекты, по которым требуется актуализация данных:",
        "",
    ]

    for index, item in enumerate(issues, start=1):
        municipality = item.get("municipality", "Не указан")
        organization = item.get("organization", "Не указана")
        missing_fields = ", ".join(item.get("missing_fields", []))

        lines.append(
            f"{index}. {municipality} / {organization} — не заполнены поля: {missing_fields}."
        )

    lines.append("")
    lines.append("Просьба актуализировать данные по указанным объектам.")

    return "\n".join(lines)


def build_default_personal_message(item):
    municipality = item.get("municipality", "не указан")
    organization = item.get("organization", "не указана")
    task_id = item.get("task_id", "")
    missing_fields = ", ".join(item.get("missing_fields", []))

    task_part = f" по задаче №{task_id}" if task_id else ""

    return (
        "Здравствуйте!\n\n"
        f"В системе WaterControl по объекту {organization} ({municipality}){task_part} "
        "обнаружены незаполненные обязательные поля:\n"
        f"{missing_fields}.\n\n"
        "Просим актуализировать сведения в возможно короткий срок.\n\n"
        "Буду признателен за оперативное реагирование."
    )


def build_personal_messages(issues):
    messages = []

    for item in issues:
        if not item.get("missing_fields"):
            continue

        text = build_default_personal_message(item)

        messages.append({
            "responsible_name": item.get("responsible_name", "Ответственный не указан"),
            "responsible_phone": item.get("responsible_phone", ""),
            "municipality": item.get("municipality", ""),
            "organization": item.get("organization", ""),
            "status": item.get("status", "risk"),
            "is_edited": False,
            "message": text,
        })

    return messages


def build_summary(issues, total_rows):
    if not issues:
        return (
            f"Проверка завершена. Проанализировано строк: {total_rows}. "
            "Незаполненных обязательных полей не обнаружено."
        )

    return (
        f"Проверка завершена. Проанализировано строк: {total_rows}. "
        f"Обнаружено проблемных строк с незаполненными обязательными полями: {len(issues)}."
    )


def merge_with_saved_personal_messages(new_personal_messages, issues):
    candidates = [
        RESULT_FILE,
        DATA_DIR / "latest_result.json",
        DATA_DIR / "last_result.json",
    ]

    old_result = None
    for path in candidates:
        if path.exists():
            try:
                old_result = json.loads(path.read_text(encoding="utf-8"))
                break
            except Exception:
                continue

    if not old_result:
        return new_personal_messages

    old_messages = old_result.get("personal_messages", [])
    old_rows = old_result.get("rows", [])

    old_map = {}
    for item in old_messages:
        key = (
            normalize_municipality_name(item.get("municipality", "")),
            normalize_org_name(item.get("organization", "")),
        )
        old_map[key] = item

    old_rows_map = {}
    for item in old_rows:
        key = (
            normalize_municipality_name(item.get("municipality", "")),
            normalize_org_name(item.get("organization", "")),
        )
        old_rows_map[key] = item

    new_rows_map = {}
    for item in issues:
        key = (
            normalize_municipality_name(item.get("municipality", "")),
            normalize_org_name(item.get("organization", "")),
        )
        new_rows_map[key] = item

    merged = []
    for item in new_personal_messages:
        key = (
            normalize_municipality_name(item.get("municipality", "")),
            normalize_org_name(item.get("organization", "")),
        )

        if key in old_map:
            old_item = old_map[key]
            old_issue = old_rows_map.get(key, {})
            new_issue = new_rows_map.get(key, {})

            old_missing_fields = old_issue.get("missing_fields", [])
            new_missing_fields = new_issue.get("missing_fields", [])

            if old_missing_fields == new_missing_fields:
                saved_message = normalize_text(old_item.get("message", ""))
                if saved_message:
                    item["message"] = saved_message

                item["is_edited"] = old_item.get("is_edited", False)
            else:
                item["is_edited"] = False

        merged.append(item)

    return merged


def detect_best_table(page):
    tables = extract_rows_from_html_table(page)

    for table in tables:
        headers = table.get("headers", [])
        rows = table.get("rows", [])
        if looks_like_target_headers(headers) and rows:
            return headers, rows, "Данные извлечены из HTML-таблицы."

    return [], [], ""


def build_row_result(row_dict, responsibles_map):
    mapped = map_row_by_aliases(row_dict, COLUMN_ALIASES)

    municipality = mapped.get("municipality", "")
    organization = mapped.get("organization", "")
    address = mapped.get("address", "")
    task_id = mapped.get("task_id", "")
    responsible_name = mapped.get("responsible_name", "")
    responsible_phone = mapped.get("responsible_phone", "")

    if not responsible_name and municipality:
        responsible_info = get_responsible_info_by_municipality(municipality, responsibles_map)
        responsible_name = responsible_info["name"]
        responsible_phone = responsible_phone or responsible_info["phone"]

    missing_fields = find_missing_fields(row_dict)

    status = "ok"
    reason = "Все обязательные поля заполнены"

    if missing_fields:
        status = "risk"
        reason = "Не заполнены обязательные поля: " + ", ".join(missing_fields)

    return {
        "municipality": municipality or "Не указан",
        "organization": organization or address or "Не указана",
        "address": address or organization or "",
        "task_id": task_id,
        "responsible_name": responsible_name or "Ответственный не указан",
        "responsible_phone": responsible_phone,
        "status": status,
        "reason": reason,
        "missing_fields": missing_fields,
    }


def try_extract_by_html_table(page, responsibles_map):
    headers, rows, extraction_note = detect_best_table(page)
    if not headers or not rows:
        return [], ""

    cleaned_rows = cleanup_table_rows(headers, rows)
    result_rows = [build_row_result(row, responsibles_map) for row in cleaned_rows]
    return result_rows, extraction_note


def try_extract_by_generic_grid(page, responsibles_map):
    raw_rows = extract_rows_from_generic_grid(page)
    result_rows = []

    for parts in raw_rows:
        if len(parts) < 4:
            continue

        row = {
            "Муниципалитет": parts[0] if len(parts) > 0 else "",
            "Организация": parts[1] if len(parts) > 1 else "",
            "ID задачи": parts[2] if len(parts) > 2 else "",
            "Дата плановой промывки": parts[3] if len(parts) > 3 else "",
            "Ссылка на акт": parts[4] if len(parts) > 4 else "",
            "Адрес": parts[1] if len(parts) > 1 else "",
        }
        result_rows.append(build_row_result(row, responsibles_map))

    return result_rows


def run():
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    screenshot_base = f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    issues = []
    rows_result = []
    extraction_note = ""
    total_rows = 0
    screenshot_paths = []

    responsibles_map = load_responsibles_by_municipality()

    try:
        with sync_playwright() as p:
            emit_stage("Открытие панели")

            browser = p.chromium.launch(
                headless=HEADLESS,
                slow_mo=300,
                args=[
                    "--start-maximized",
                    "--disable-blink-features=AutomationControlled",
                ],
            )

            context = browser.new_context(
                viewport={"width": 1600, "height": 1200},
                user_agent=(
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
                locale="ru-RU",
            )

            page = context.new_page()

            page.goto(WATERCONTROL_URL, wait_until="domcontentloaded", timeout=120000)
            page.wait_for_timeout(5000)

            if is_captcha_page(page):
                passed = wait_for_manual_captcha_pass(page, max_wait_seconds=180)
                if not passed:
                    screenshot_paths = save_table_screenshots(page, base_name=screenshot_base, max_parts=1)
                    browser.close()
                    save_error_result(
                        timestamp,
                        screenshot_paths,
                        "Доступ остановлен капчей. Капча не была пройдена вручную за отведённое время.",
                    )
                    return

            page.wait_for_timeout(4000)

            emit_stage("Снятие скриншотов")
            screenshot_paths = save_table_screenshots(page, base_name=screenshot_base, max_parts=8)

            if is_captcha_page(page):
                browser.close()
                save_error_result(
                    timestamp,
                    screenshot_paths,
                    "После ожидания страница всё ещё находится на капче.",
                )
                return

            emit_stage("Анализ данных")
            rows_result, extraction_note = try_extract_by_html_table(page, responsibles_map)

            if not rows_result:
                rows_result = try_extract_by_generic_grid(page, responsibles_map)
                if rows_result:
                    extraction_note = "Данные извлечены из grid-структуры."
                else:
                    extraction_note = (
                        "Строки таблицы не найдены. Возможно, структура WaterControl отличается "
                        "от ожидаемой или данные не успели прогрузиться."
                    )

            total_rows = len(rows_result)

            if DEBUG_ROWS:
                print("Извлечённые строки:", flush=True)
                for row in rows_result[:30]:
                    print(row, flush=True)

            for row in rows_result:
                if row.get("status") != "ok":
                    issues.append(row)

            browser.close()

    except TimeoutError:
        save_error_result(
            timestamp,
            screenshot_paths,
            "Timeout загрузки страницы WaterControl",
        )
        sys.exit(1)

    except PlaywrightError as e:
        save_error_result(
            timestamp,
            screenshot_paths,
            f"Ошибка открытия или обработки страницы: {str(e)}",
        )
        sys.exit(1)

    except Exception as e:
        save_error_result(
            timestamp,
            screenshot_paths,
            f"Непредвиденная ошибка: {str(e)}",
        )
        sys.exit(1)

    emit_stage("Формирование сводки")
    public_chat_message = build_public_chat_message(issues)
    personal_messages = build_personal_messages(issues)
    personal_messages = merge_with_saved_personal_messages(personal_messages, issues)
    summary_message = build_summary(issues, total_rows)

    result = {
        "created_at": timestamp,
        "summary_message": summary_message,
        "public_chat_message": public_chat_message,
        "personal_messages": personal_messages,
        "rows": rows_result,
        "screenshot_path": screenshot_paths[0] if screenshot_paths else "",
        "screenshot_paths": screenshot_paths,
        "source_url": WATERCONTROL_URL,
        "extraction_note": extraction_note,
    }

    emit_stage("Сохранение результата")
    save_result(result)
    print("Готово:", RESULT_FILE, flush=True)


if __name__ == "__main__":
    run()