import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError, Error as PlaywrightError


BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data" / "edo"
SCREENSHOTS_DIR = DATA_DIR / "screenshots"
RESULT_FILE = DATA_DIR / "result.json"

RESPONSIBLES_CANDIDATES = [
    DATA_DIR / "responsible_by_municipalyty.json",
    DATA_DIR / "responsibles_by_municipality.json",
    BASE_DIR / "data" / "responsibles_by_municipality.json",
]

DATALENS_URL = os.getenv(
    "EDO_DASHBOARD_URL",
    os.getenv("DASHBOARD_URL", "https://datalens.yandex/f5wqqij889haz?tab=EL")
)

REDMINE_URL = os.getenv(
    "EDO_REDMINE_URL",
    "https://docs.google.com/spreadsheets/d/1tZeDsYYNo0iD0cCwcAAQX6eS50jgknfn9R6Zx079PWY/edit?gid=1926946402#gid=1926946402"
)

HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
DEBUG_ROWS = os.getenv("DEBUG_ROWS", "false").lower() == "true"

SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)
RESULT_FILE.parent.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)

EMPTY_MARKERS = {"", "null", "none", "-", "—", "n/a", "nan"}
WEEK_RANGE_RE = re.compile(r"\d{4}\.\d{2}\.\d{2}\s*-\s*\d{4}\.\d{2}\.\d{2}")


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


def format_empty_fields(empty_fields):
    pretty_names = {
        "Общее кол-во документов в обороте (вн)": "Общее кол-во документов в обороте (вн)",
        "Кол-во документов в эл виде (вн)": "Кол-во документов в эл. виде (вн)",
        "Общее кол-во документов в обороте (вх)": "Общее кол-во документов в обороте (вх)",
        "Кол-во документов в эл виде (вх)": "Кол-во документов в эл. виде (вх)",
        "Общее кол-во документов в обороте (исх)": "Общее кол-во документов в обороте (исх)",
        "Кол-во документов в эл виде (исх)": "Кол-во документов в эл. виде (исх)",
    }
    return ", ".join(pretty_names.get(field, field) for field in empty_fields)


def get_responsibles_file():
    for path in RESPONSIBLES_CANDIDATES:
        if path.exists():
            print(f"RESPONSIBLES FILE FOUND: {path}", flush=True)
            return path
    print("RESPONSIBLES FILE NOT FOUND", flush=True)
    return None


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
    ]
    return any(marker in content for marker in markers)


def wait_for_manual_captcha_pass(page, max_wait_seconds=180):
    print("Обнаружена SmartCaptcha. Пройдите её вручную в открытом браузере...", flush=True)
    start_time = time.time()

    while time.time() - start_time < max_wait_seconds:
        page.wait_for_timeout(2000)

        try:
            if not is_captcha_page(page):
                page.wait_for_load_state("domcontentloaded", timeout=30000)
                page.wait_for_timeout(5000)
                print("SmartCaptcha пройдена, продолжаем обработку.", flush=True)
                return True
        except Exception:
            pass

    return False


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
        "missing_data_issues": [],
        "rows": [],
        "screenshot_path": first_screenshot,
        "screenshot_paths": screenshot_paths,
        "source_url": DATALENS_URL,
        "redmine_url": REDMINE_URL,
        "extraction_note": message,
    }
    save_result(result)


def build_public_chat_message(issues):
    if not issues:
        return "По результатам автоматической проверки незаполненных данных не выявлено."

    lines = [
        "В ходе автоматической проверки выявлены объекты, по которым требуется актуализация данных:",
        "",
    ]

    for index, item in enumerate(issues, start=1):
        municipality = item.get("municipality", "Не указан")
        organization = item.get("organization", "Не указана")
        empty_fields = format_empty_fields(item.get("empty_fields", []))

        lines.append(
            f"{index}. {municipality} / {organization} — не заполнены поля: {empty_fields}."
        )

    lines.append("")
    lines.append("Просьба актуализировать данные по указанным объектам.")

    return "\n".join(lines)


def build_default_personal_message(item):
    municipality = item.get("municipality", "не указан")
    organization = item.get("organization", "не указана")
    empty_fields = format_empty_fields(item.get("empty_fields", []))

    return (
        "Здравствуйте!\n\n"
        'Для расчёта показателя по трансформационному проекту "Переход РСО на ЭДО" '
        f"по организации {organization} ({municipality}) отсутствуют данные:\n"
        f"{empty_fields}.\n\n"
        "Из-за отсутствия расчетных значений Ваш муниципалитет находится в красной зоне рейтинга МинЖКХ.\n"
        f"Ссылка на таблицу: {REDMINE_URL}\n\n"
        "Сведения нужно внести до конца следующего рабочего дня.\n\n"
        "Буду признателен за оперативное реагирование."
    )


def build_personal_messages(issues):
    messages = []

    for item in issues:
        if not item.get("empty_fields"):
            continue

        text = build_default_personal_message(item)

        messages.append({
            "responsible_name": item.get("responsible_name", "Ответственный не указан"),
            "responsible_phone": item.get("responsible_phone", ""),
            "municipality": item.get("municipality", ""),
            "organization": item.get("organization", ""),
            "metric_name": "Незаполненные данные",
            "status": "risk",
            "is_edited": False,
            "message": text,
        })

    return messages


def build_summary(issues, total_rows):
    if not issues:
        return (
            f"Проверка завершена. Проанализировано строк: {total_rows}. "
            "Незаполненных данных не обнаружено."
        )

    return (
        f"Проверка завершена. Проанализировано строк: {total_rows}. "
        f"Обнаружено проблемных строк с незаполненными полями: {len(issues)}."
    )


def extract_rows_from_html_table(page):
    rows_data = []

    rows = page.locator("table tbody tr")
    count = rows.count()

    for i in range(count):
        row = rows.nth(i)
        cells = row.locator("td")

        values = []
        for j in range(cells.count()):
            try:
                values.append(normalize_text(cells.nth(j).inner_text()))
            except Exception:
                values.append("")

        if values:
            rows_data.append(values)

    return rows_data


def extract_rows_from_generic_grid(page):
    rows_data = []

    selectors = [
        '[role="row"]',
        '.table__row',
        '.grid-row',
        '.datalens-table__row',
        '[role="grid"] [role="row"]',
    ]

    for selector in selectors:
        rows = page.locator(selector)

        try:
            rows_count = rows.count()
        except Exception:
            rows_count = 0

        for i in range(rows_count):
            row = rows.nth(i)
            try:
                text = normalize_text(row.inner_text())
            except Exception:
                continue

            if not text:
                continue

            parts = [p.strip() for p in text.split("\n")]

            if len(parts) >= 8:
                rows_data.append(parts[:8])

        if rows_data:
            return rows_data

    return rows_data


def cleanup_rows(rows_data):
    cleaned = []
    seen = set()

    for values in rows_data:
        normalized = [normalize_text(v) for v in values]

        if len(normalized) < 8:
            continue

        normalized = normalized[:8]

        first_cell = normalized[0].lower()
        second_cell = normalized[1].lower()

        if first_cell == "организация" or second_cell == "муниципалитет":
            continue

        if not normalized[0] or not normalized[1]:
            continue

        key = tuple(normalized)
        if key in seen:
            continue

        seen.add(key)
        cleaned.append(normalized)

    return cleaned


def map_row_to_issue(values, row_index, responsibles_map):
    if len(values) < 8:
        return None

    organization = normalize_text(values[0])
    municipality = normalize_text(values[1])
    total_docs_vn = normalize_text(values[2])
    electronic_docs_vn = normalize_text(values[3])
    total_docs_vh = normalize_text(values[4])
    electronic_docs_vh = normalize_text(values[5])
    total_docs_ish = normalize_text(values[6])
    electronic_docs_ish = normalize_text(values[7])

    responsible_info = get_responsible_info_by_municipality(municipality, responsibles_map)
    responsible_name = responsible_info["name"]
    responsible_phone = responsible_info["phone"]

    field_checks = [
        ("Общее кол-во документов в обороте (вн)", total_docs_vn),
        ("Кол-во документов в эл виде (вн)", electronic_docs_vn),
        ("Общее кол-во документов в обороте (вх)", total_docs_vh),
        ("Кол-во документов в эл виде (вх)", electronic_docs_vh),
        ("Общее кол-во документов в обороте (исх)", total_docs_ish),
        ("Кол-во документов в эл виде (исх)", electronic_docs_ish),
    ]

    empty_fields = [
        field_name for field_name, field_value in field_checks
        if is_empty_value(field_value)
    ]

    if not empty_fields:
        return None

    return {
        "row_index": row_index + 1,
        "organization": organization or "Не указана",
        "municipality": municipality or "Не указан",
        "responsible_name": responsible_name,
        "responsible_phone": responsible_phone,
        "empty_fields": empty_fields,
        "message": (
            "Обнаружены незаполненные данные. "
            f"Муниципалитет: {municipality or 'не указан'}. "
            f"Организация: {organization or 'не указана'}. "
            f"Не заполнены поля: {', '.join(empty_fields)}."
        ),
    }


def find_scrollable_container(page):
    selectors = [
        '[role="grid"]',
        '.datalens-table',
        '.widget-table',
        '.table',
        '.grid-container',
        '.ql-table-container',
        '.njs-scrollable',
        '.scrollable',
        '[class*="scroll"]',
    ]

    for selector in selectors:
        locator = page.locator(selector)

        try:
            count = locator.count()
        except Exception:
            count = 0

        for i in range(min(count, 10)):
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
        return [f"data/edo/screenshots/{file_name}"]

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

        scrollable_locator.screenshot(path=str(file_path))
        saved_files.append(f"data/edo/screenshots/{file_name}")

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
                scrollable_locator.screenshot(path=str(final_path))
                saved_files.append(f"data/edo/screenshots/{final_name}")
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
    old_issues = old_result.get("missing_data_issues", [])

    old_map = {}
    for item in old_messages:
        key = (
            normalize_municipality_name(item.get("municipality", "")),
            normalize_org_name(item.get("organization", "")),
        )
        old_map[key] = item

    old_issues_map = {}
    for item in old_issues:
        key = (
            normalize_municipality_name(item.get("municipality", "")),
            normalize_org_name(item.get("organization", "")),
        )
        old_issues_map[key] = item

    new_issues_map = {}
    for item in issues:
        key = (
            normalize_municipality_name(item.get("municipality", "")),
            normalize_org_name(item.get("organization", "")),
        )
        new_issues_map[key] = item

    merged = []
    for item in new_personal_messages:
        key = (
            normalize_municipality_name(item.get("municipality", "")),
            normalize_org_name(item.get("organization", "")),
        )

        if key in old_map:
            old_item = old_map[key]
            old_issue = old_issues_map.get(key, {})
            new_issue = new_issues_map.get(key, {})

            old_empty_fields = old_issue.get("empty_fields", [])
            new_empty_fields = new_issue.get("empty_fields", [])

            if old_empty_fields == new_empty_fields:
                saved_message = normalize_text(old_item.get("message", ""))
                if saved_message:
                    item["message"] = saved_message

                item["is_edited"] = old_item.get("is_edited", False)
            else:
                item["is_edited"] = False

        merged.append(item)

    return merged


def deduplicate_week_candidates(candidates):
    result = []
    seen = set()

    for text, item in candidates:
        cleaned = normalize_text(text)
        if cleaned not in seen:
            seen.add(cleaned)
            result.append((cleaned, item))

    return result


def get_current_week_value(page):
    locator = page.locator(".yc-select-control__tokens-text")

    try:
        count = locator.count()
    except Exception:
        count = 0

    for i in range(min(count, 20)):
        item = locator.nth(i)
        try:
            text = normalize_text(item.inner_text())
        except Exception:
            text = ""

        match = WEEK_RANGE_RE.search(text)
        if match:
            return match.group(0)

    return ""


def find_week_control(page):
    tokens = page.locator(".yc-select-control__tokens-text")

    try:
        count = tokens.count()
    except Exception:
        count = 0

    for i in range(min(count, 20)):
        item = tokens.nth(i)

        try:
            text = normalize_text(item.inner_text())
        except Exception:
            text = ""

        if not WEEK_RANGE_RE.search(text):
            continue

        try:
            control = item.locator("xpath=ancestor::div[contains(@class, 'yc-select-control')]").first
            return control
        except Exception:
            continue

    return None


def open_week_dropdown(page):
    print("Пытаюсь открыть dropdown недели...", flush=True)

    control = find_week_control(page)
    if control is None:
        print("Не найден контрол недели.", flush=True)
        return False

    try:
        control.click(timeout=3000)
        page.wait_for_timeout(1200)
        print("Dropdown недели открыт.", flush=True)
        return True
    except Exception as e:
        print("Не удалось кликнуть по контролу недели:", str(e), flush=True)
        return False


def collect_week_candidates_from_portal(page):
    candidates = []

    portal_roots = page.locator("[data-floating-ui-portal]")

    try:
        portal_count = portal_roots.count()
    except Exception:
        portal_count = 0

    for p_idx in range(min(portal_count, 10)):
        portal = portal_roots.nth(p_idx)

        inner = portal.locator("*")
        try:
            count = inner.count()
        except Exception:
            count = 0

        for i in range(min(count, 500)):
            item = inner.nth(i)

            try:
                text = normalize_text(item.inner_text())
            except Exception:
                text = ""

            match = WEEK_RANGE_RE.search(text)
            if match:
                candidates.append((match.group(0), item))

    return deduplicate_week_candidates(candidates)


def collect_week_candidates_fallback(page):
    candidates = []

    selectors = [
        '[role="option"]',
        '[role="listbox"] *',
        '.popup *',
        '.menu *',
        '.popover *',
        '.rc-virtual-list *',
        'li',
        'div',
        'span',
    ]

    for selector in selectors:
        locator = page.locator(selector)

        try:
            count = locator.count()
        except Exception:
            count = 0

        for i in range(min(count, 800)):
            item = locator.nth(i)

            try:
                text = normalize_text(item.inner_text())
            except Exception:
                text = ""

            match = WEEK_RANGE_RE.search(text)
            if match:
                candidates.append((match.group(0), item))

        if candidates:
            break

    return deduplicate_week_candidates(candidates)


def select_latest_week(page):
    emit_stage("Выбор последней недели")

    try:
        page.wait_for_timeout(4000)

        current_week = get_current_week_value(page)
        print("Текущая неделя до выбора:", current_week or "НЕ НАЙДЕНА", flush=True)

        opened = open_week_dropdown(page)
        if not opened:
            return False

        page.wait_for_timeout(1200)

        candidates = collect_week_candidates_from_portal(page)

        if not candidates:
            candidates = collect_week_candidates_fallback(page)

        print("Найдено кандидатов недель:", len(candidates), flush=True)

        if not candidates:
            return False

        filtered = candidates
        if current_week:
            without_current = [item for item in candidates if item[0] != current_week]
            if without_current:
                filtered = without_current

        latest_text, latest_item = filtered[-1]
        print("Выбираю последнюю неделю:", latest_text, flush=True)

        try:
            latest_item.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            pass

        page.wait_for_timeout(300)

        try:
            latest_item.click(timeout=3000)
        except Exception:
            try:
                box = latest_item.bounding_box()
                if box:
                    page.mouse.click(
                        box["x"] + box["width"] / 2,
                        box["y"] + box["height"] / 2,
                    )
                else:
                    raise Exception("У элемента нет bounding_box")
            except Exception as e:
                print("Не удалось выбрать последнюю неделю:", str(e), flush=True)
                return False

        page.wait_for_timeout(5000)

        new_week = get_current_week_value(page)
        print("Неделя после выбора:", new_week or "НЕ НАЙДЕНА", flush=True)

        if new_week == latest_text:
            return True

        return False

    except Exception as e:
        print("Ошибка при выборе последней недели:", str(e), flush=True)
        return False


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

            page.goto(DATALENS_URL, wait_until="domcontentloaded", timeout=120000)
            page.wait_for_timeout(5000)

            if is_captcha_page(page):
                passed = wait_for_manual_captcha_pass(page, max_wait_seconds=180)
                if not passed:
                    screenshot_paths = save_table_screenshots(page, base_name=screenshot_base, max_parts=1)
                    browser.close()
                    save_error_result(
                        timestamp,
                        screenshot_paths,
                        "Доступ остановлен SmartCaptcha. Капча не была пройдена вручную за отведённое время.",
                    )
                    return

            page.wait_for_timeout(3000)
            select_latest_week(page)
            page.wait_for_timeout(5000)

            emit_stage("Снятие скриншотов")
            screenshot_paths = save_table_screenshots(page, base_name=screenshot_base, max_parts=8)

            if is_captcha_page(page):
                browser.close()
                save_error_result(
                    timestamp,
                    screenshot_paths,
                    "После ожидания страница всё ещё находится на SmartCaptcha.",
                )
                return

            emit_stage("Анализ данных")
            rows_data = extract_rows_from_html_table(page)

            if not rows_data:
                rows_data = extract_rows_from_generic_grid(page)
                if rows_data:
                    extraction_note = "Данные извлечены из grid-структуры."
                else:
                    extraction_note = (
                        "Строки таблицы не найдены. Возможно, структура DataLens отличается "
                        "от ожидаемой или данные не успели прогрузиться."
                    )
            else:
                extraction_note = "Данные извлечены из HTML-таблицы."

            rows_data = cleanup_rows(rows_data)
            total_rows = len(rows_data)

            if DEBUG_ROWS:
                print("Извлечённые строки:", flush=True)
                for row in rows_data[:30]:
                    print(row, flush=True)

            for idx, values in enumerate(rows_data):
                if len(values) < 8:
                    continue

                organization = values[0] if len(values) > 0 else ""
                municipality = values[1] if len(values) > 1 else ""
                total_docs_vn = values[2] if len(values) > 2 else ""
                electronic_docs_vn = values[3] if len(values) > 3 else ""
                total_docs_vh = values[4] if len(values) > 4 else ""
                electronic_docs_vh = values[5] if len(values) > 5 else ""
                total_docs_ish = values[6] if len(values) > 6 else ""
                electronic_docs_ish = values[7] if len(values) > 7 else ""

                responsible_info = get_responsible_info_by_municipality(municipality, responsibles_map)
                responsible_name = responsible_info["name"]
                responsible_phone = responsible_info["phone"]

                print(
                    "ROW DEBUG:",
                    normalize_municipality_name(municipality),
                    "->",
                    responsible_name,
                    responsible_phone,
                    flush=True,
                )

                issue = map_row_to_issue(values, idx, responsibles_map)

                row_status = "ok"
                reason = "Заполнено"
                empty_fields = []

                if issue:
                    issues.append(issue)
                    row_status = "risk"
                    reason = "Есть незаполненные поля"
                    empty_fields = issue.get("empty_fields", [])

                rows_result.append({
                    "municipality": municipality,
                    "organization": organization,
                    "metric_name": "Контроль заполненности",
                    "current_value": total_docs_vn,
                    "target_value": electronic_docs_vn,
                    "delay_days": total_docs_vh,
                    "responsible_name": responsible_name,
                    "responsible_phone": responsible_phone,
                    "status": row_status,
                    "reason": reason,
                    "empty_fields": empty_fields,
                    "total_docs_vn": total_docs_vn,
                    "electronic_docs_vn": electronic_docs_vn,
                    "total_docs_vh": total_docs_vh,
                    "electronic_docs_vh": electronic_docs_vh,
                    "total_docs_ish": total_docs_ish,
                    "electronic_docs_ish": electronic_docs_ish,
                })

            browser.close()

    except TimeoutError:
        save_error_result(
            timestamp,
            screenshot_paths,
            "Timeout загрузки страницы DataLens",
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
        "missing_data_issues": issues,
        "rows": rows_result,
        "screenshot_path": screenshot_paths[0] if screenshot_paths else "",
        "screenshot_paths": screenshot_paths,
        "source_url": DATALENS_URL,
        "redmine_url": REDMINE_URL,
        "extraction_note": extraction_note,
    }

    emit_stage("Сохранение результата")
    save_result(result)
    print("Готово:", RESULT_FILE, flush=True)


if __name__ == "__main__":
    run()