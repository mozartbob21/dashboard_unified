import json
import shutil
import time
from datetime import datetime

from playwright.sync_api import sync_playwright

from services.overdue.config import (
    DASHBOARD_URL,
    DATA_FILE,
    DEBUG_DIR,
    DEBUG_ENV_DIR,
    HEADLESS,
    IDLE_AFTER_DATA_SECONDS,
    MAX_WAIT_SECONDS,
    PLAYWRIGHT_PROFILE_DIR,
    RESPONSES_DIR,
)
from services.overdue.utils import load_json, save_json, safe_int


BLOCKED_HOST_PARTS = [
    "mc.yandex.ru",
    "metrika",
    "smartcaptcha",
    "showcaptcha",
    "captcha",
]


def prepare_dirs():
    PLAYWRIGHT_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    DEBUG_ENV_DIR.mkdir(parents=True, exist_ok=True)

    if RESPONSES_DIR.exists():
        shutil.rmtree(RESPONSES_DIR)

    RESPONSES_DIR.mkdir(parents=True, exist_ok=True)


def is_blocked_url(url: str) -> bool:
    lower_url = url.lower()
    return any(part in lower_url for part in BLOCKED_HOST_PARTS)


def looks_like_useful_json(data) -> bool:
    if isinstance(data, list):
        return len(data) > 0
    if isinstance(data, dict):
        return len(data) > 0
    return False


def try_launch_context(playwright):
    errors = []

    launch_variants = [
        {
            "name": "channel=chrome",
            "kwargs": {
                "user_data_dir": str(PLAYWRIGHT_PROFILE_DIR),
                "channel": "chrome",
                "headless": HEADLESS,
                "viewport": {"width": 1440, "height": 1100},
                "args": [
                    "--disable-blink-features=AutomationControlled",
                    "--start-maximized",
                ],
            },
        },
        {
            "name": "playwright chromium fallback",
            "kwargs": {
                "user_data_dir": str(PLAYWRIGHT_PROFILE_DIR),
                "headless": HEADLESS,
                "viewport": {"width": 1440, "height": 1100},
                "args": [
                    "--disable-blink-features=AutomationControlled",
                    "--start-maximized",
                ],
            },
        },
    ]

    for variant in launch_variants:
        try:
            print(f"[browser] trying: {variant['name']}")
            context = playwright.chromium.launch_persistent_context(**variant["kwargs"])
            print(f"[browser] success: {variant['name']}")
            return context, variant["name"]
        except Exception as e:
            errors.append(f"{variant['name']}: {e}")

    raise RuntimeError(
        "Не удалось запустить браузер.\n\n"
        + "\n\n".join(errors)
        + "\n\nУстановите браузер командой:\npython -m playwright install\n"
    )


def fetch_dashboard_data():
    prepare_dirs()

    screenshot_file = DEBUG_ENV_DIR / "dashboard_page.png"
    html_file = DEBUG_ENV_DIR / "dashboard_page.html"

    screenshot_web_path = "data/overdue/debug/dashboard_page.png"
    html_web_path = "data/overdue/debug/dashboard_page.html"

    saved_files = []
    saved_count = 0

    state = {
        "public_entry_seen": False,
        "dash_state_seen": False,
        "chart_run_count": 0,
        "last_useful_response_ts": None,
    }

    with sync_playwright() as p:
        context, browser_name = try_launch_context(p)
        page = context.pages[0] if context.pages else context.new_page()

        def handle_response(response):
            nonlocal saved_count

            try:
                url = response.url
                lower_url = url.lower()

                if is_blocked_url(lower_url):
                    return

                content_type = response.headers.get("content-type", "").lower()
                if "json" not in content_type:
                    return

                if response.status >= 400:
                    return

                data = response.json()
                if not looks_like_useful_json(data):
                    return

                saved_count += 1
                file_path = RESPONSES_DIR / f"{saved_count:03d}.json"

                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(
                        {
                            "url": url,
                            "status": response.status,
                            "content_type": content_type,
                            "data": data,
                        },
                        f,
                        ensure_ascii=False,
                        indent=2,
                    )

                saved_files.append(str(file_path.resolve()))
                state["last_useful_response_ts"] = time.time()

                if "getpublicentry" in lower_url:
                    state["public_entry_seen"] = True

                if "getpublicdashstate" in lower_url:
                    state["dash_state_seen"] = True

                if "/charts/api/run" in lower_url:
                    state["chart_run_count"] += 1

                print(f"[saved] {saved_count:03d} {url}")

            except Exception:
                pass

        page.on("response", handle_response)

        print(f"[open] {DASHBOARD_URL}")
        print(f"[browser] using: {browser_name}")
        page.goto(DASHBOARD_URL, wait_until="domcontentloaded", timeout=120000)

        print("")
        print("Если нужно — пройдите капчу / авторизацию вручную в окне браузера.")
        print("Ожидание завершится автоматически, когда дашборд догрузится.")
        print(f"Максимальный таймаут: {MAX_WAIT_SECONDS} сек.")
        print("")

        start_ts = time.time()

        while True:
            now = time.time()
            elapsed = now - start_ts

            enough_data_loaded = (
                state["public_entry_seen"]
                and state["dash_state_seen"]
                and state["chart_run_count"] >= 1
            )

            idle_enough = (
                state["last_useful_response_ts"] is not None
                and (now - state["last_useful_response_ts"]) >= IDLE_AFTER_DATA_SECONDS
            )

            if enough_data_loaded and idle_enough:
                print("[wait] Данные загружены, завершаем ожидание.")
                break

            if elapsed >= MAX_WAIT_SECONDS:
                print("[wait] Достигнут максимальный таймаут.")
                break

            page.wait_for_timeout(1000)

        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass

        final_url = page.url
        title = page.title()

        try:
            html_file.write_text(page.content(), encoding="utf-8")
        except Exception:
            pass

        try:
            page.screenshot(path=str(screenshot_file), full_page=True)
            print(f"[saved] screenshot: {screenshot_file}")
        except Exception:
            pass

        context.close()

    print(f"[done] saved JSON responses: {saved_count}")

    return {
        "screenshot_path": screenshot_web_path,
        "screenshot_paths": [screenshot_web_path],
        "html_path": html_web_path,
        "responses_dir": str(RESPONSES_DIR.resolve()),
        "response_files": saved_files,
        "saved_json_count": saved_count,
        "final_url": final_url,
        "title": title,
        "public_entry_seen": state["public_entry_seen"],
        "dash_state_seen": state["dash_state_seen"],
        "chart_run_count": state["chart_run_count"],
    }


def load_wrappers():
    wrappers = []
    if not RESPONSES_DIR.exists():
        return wrappers

    for file_path in sorted(RESPONSES_DIR.glob("*.json")):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                wrapper = json.load(f)

            wrappers.append(
                {
                    "file": str(file_path),
                    "url": wrapper.get("url", ""),
                    "status": wrapper.get("status", 0),
                    "data": wrapper.get("data"),
                }
            )
        except Exception:
            continue

    return wrappers


def find_omsu_chart_id(wrappers):
    for wrapper in wrappers:
        url = (wrapper.get("url") or "").lower()
        if "getpublicentry" not in url:
            continue

        payload = wrapper.get("data") or {}
        dash_data = payload.get("data") or {}
        tabs = dash_data.get("tabs") or []

        for tab in tabs:
            for item in tab.get("items", []):
                if item.get("type") != "widget":
                    continue

                item_data = item.get("data") or {}
                widget_tabs = item_data.get("tabs") or []

                for widget_tab in widget_tabs:
                    title = (widget_tab.get("title") or "").strip().lower()
                    chart_id = (widget_tab.get("chartId") or "").strip()

                    if title == "омсу" and chart_id:
                        return chart_id

    return None


def extract_from_chart_run_response(wrapper, omsu_chart_id):
    url = wrapper.get("url") or ""
    payload = wrapper.get("data") or {}

    if "/charts/api/run" not in url:
        return []

    wrapper_chart_id = payload.get("id") or payload.get("_confStorageConfig", {}).get("entryId")
    if omsu_chart_id and wrapper_chart_id != omsu_chart_id:
        return []

    chart_data = payload.get("data") or {}
    categories = chart_data.get("categories") or []
    graphs = chart_data.get("graphs") or []

    if not categories or not graphs:
        return []

    first_graph = graphs[0] or {}
    points = first_graph.get("data") or []

    items = []
    max_len = min(len(categories), len(points))

    for i in range(max_len):
        municipality = str(categories[i]).strip()
        point = points[i] or {}
        overdue_count = safe_int(point.get("y", 0), 0)

        if not municipality:
            continue

        items.append(
            {
                "municipality": municipality,
                "organization": municipality,
                "overdue_count": overdue_count,
                "responsible_name": "",
                "responsible_phone": "",
            }
        )

    return items


def extract_dashboard_data():
    wrappers = load_wrappers()
    omsu_chart_id = find_omsu_chart_id(wrappers)

    best_items = []
    matched_sources = []

    for wrapper in wrappers:
        items = extract_from_chart_run_response(wrapper, omsu_chart_id)

        if items:
            matched_sources.append(
                {
                    "file": wrapper.get("file", ""),
                    "url": wrapper.get("url", ""),
                    "items_found": len(items),
                    "chart_id": wrapper.get("data", {}).get("id")
                    or wrapper.get("data", {}).get("_confStorageConfig", {}).get("entryId", ""),
                }
            )

            if len(items) > len(best_items):
                best_items = items

    best_items.sort(key=lambda x: (-safe_int(x["overdue_count"], 0), x["municipality"]))

    return {
        "source": "saved_network" if best_items else "empty",
        "items": best_items,
        "matched_sources": matched_sources,
        "responses_scanned": len(wrappers),
        "items_count": len(best_items),
        "summary": {
            "total_records": len(best_items),
            "critical": sum(1 for x in best_items if safe_int(x["overdue_count"], 0) >= 20),
            "risk": sum(1 for x in best_items if 0 < safe_int(x["overdue_count"], 0) < 20),
            "ok": sum(1 for x in best_items if safe_int(x["overdue_count"], 0) == 0),
        },
        "debug": {
            "omsu_chart_id": omsu_chart_id,
        },
    }


def normalize_items(raw_items):
    normalized = []

    for item in raw_items or []:
        municipality = (item.get("municipality") or item.get("name") or "Не указано").strip()
        organization = (item.get("organization") or municipality).strip()
        overdue_count = safe_int(item.get("overdue_count", item.get("count", 0)), 0)

        normalized.append(
            {
                "municipality": municipality,
                "organization": organization,
                "overdue_count": overdue_count,
                "category": "Просроченные задачи",
                "responsible_name": item.get("responsible_name", ""),
                "responsible_phone": item.get("responsible_phone", ""),
            }
        )

    normalized.sort(key=lambda x: (-x["overdue_count"], x["municipality"]))
    return normalized


def build_summary(items):
    total_overdue = sum(safe_int(item.get("overdue_count", 0), 0) for item in items)

    critical_count = 0
    risk_count = 0
    ok_count = 0
    by_municipality = []

    for item in items:
        municipality = item.get("municipality", "Не указано")
        overdue_count = safe_int(item.get("overdue_count", 0), 0)

        if overdue_count >= 20:
            critical_count += 1
        elif overdue_count > 0:
            risk_count += 1
        else:
            ok_count += 1

        by_municipality.append(
            {
                "municipality": municipality,
                "count": overdue_count,
            }
        )

    by_municipality.sort(key=lambda x: (-x["count"], x["municipality"]))

    return {
        "total_overdue": total_overdue,
        "by_status": [
            {"status": "Критично", "count": critical_count},
            {"status": "Риск", "count": risk_count},
            {"status": "Норма", "count": ok_count},
        ],
        "by_municipality": by_municipality,
        "by_category": [],
    }


def build_public_message(summary, items):
    bad_items = [x for x in items if safe_int(x.get("overdue_count", 0), 0) > 0]
    bad_items.sort(key=lambda x: (-safe_int(x.get("overdue_count", 0), 0), x.get("municipality", "")))

    lines = [
        "Добрый день.",
        "По итогам проверки информационной панели выполнения поручений, зафиксированных в системе управления МинистерстаЖКХ МО, выявлены невыполненные задачи.",
        f"Общее количество просроченных задач: {summary.get('total_overdue', 0)}.",
    ]

    if bad_items:
        lines.append("")
        lines.append("ОМСУ с наибольшим количеством просроченных задач:")
        for item in bad_items[:15]:
            lines.append(f"- {item.get('municipality', 'Не указано')}: {item.get('overdue_count', 0)}")

        lines.append("")
        lines.append("Просьба оперативно отработать просроченные позиции и актуализировать сведения.")
    else:
        lines.append("")
        lines.append("Просроченные задачи не выявлены. Спасибо за своевременное обновление данных.")

    return "\n".join(lines)


def build_missing_data_issues(items):
    issues = []

    for item in items:
        municipality = item.get("municipality", "Не указано")
        organization = item.get("organization", municipality)
        responsible_name = item.get("responsible_name", "")
        responsible_phone = item.get("responsible_phone", "")
        overdue_count = safe_int(item.get("overdue_count", 0), 0)

        missing_fields = []
        if not responsible_name:
            missing_fields.append("не указан ответственный")
        if not responsible_phone:
            missing_fields.append("не указан телефон")

        if missing_fields:
            issues.append(
                {
                    "municipality": municipality,
                    "organization": organization,
                    "responsible_name": responsible_name or "Не указан",
                    "responsible_phone": responsible_phone or "",
                    "message": (
                        f"По записи '{municipality} / {organization}' обнаружены незаполненные данные: "
                        f"{', '.join(missing_fields)}."
                        f"{' Дополнительно зафиксировано просроченных задач: ' + str(overdue_count) + '.' if overdue_count > 0 else ''}"
                    ),
                }
            )

    return issues


def build_personal_messages(items):
    messages = []

    for item in items:
        overdue_count = safe_int(item.get("overdue_count", 0), 0)
        if overdue_count <= 0:
            continue

        status = "critical" if overdue_count >= 20 else "risk"
        municipality = item.get("municipality", "Не указано")
        organization = item.get("organization", municipality)
        responsible_name = item.get("responsible_name", "Коллега")
        responsible_phone = item.get("responsible_phone", "")

        message = (
            f"Добрый день!.\n\n"
            f"По итогам проверки информационной панели выполнения поручений, зафиксированных в системе управления МинистерстаЖКХ МО по ОМСУ '{municipality}' "
            f"выявлены невыполненные задачи.: {overdue_count}.\n"
            f"Просьба проверить блок '{organization}', отработать просроченные позиции и актуализировать отчёт.\n\n"
            f"Необходимо до конца следующего рабочего дня внести комментарии о текущем статусе исполнения поручения и перевести задачу на контролёра."
        )

        messages.append(
            {
                "municipality": municipality,
                "organization": organization,
                "responsible_name": responsible_name,
                "responsible_phone": responsible_phone,
                "status": status,
                "message": message,
                "is_edited": False,
            }
        )

    messages.sort(key=lambda x: (0 if x["status"] == "critical" else 1, -len(x["message"])))
    return messages


def build_report_text(summary, items):
    lines = [
        "Текстовый отчёт",
        "",
        f"Дата формирования: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Всего просроченных задач: {summary.get('total_overdue', 0)}",
        "",
        "Распределение по статусам:",
    ]

    for item in summary.get("by_status", []):
        lines.append(f"- {item['status']}: {item['count']}")

    lines.append("")
    lines.append("Топ ОМСУ:")

    for item in summary.get("by_municipality", [])[:15]:
        lines.append(f"- {item['municipality']}: {item['count']}")

    if items:
        lines.append("")
        lines.append("Детализация:")
        for item in items[:30]:
            lines.append(f"- {item.get('municipality', 'Не указано')}: {item.get('overdue_count', 0)}")

    return "\n".join(lines)


def run_overdue_pipeline():
    screenshot_paths = []
    extraction_note = ""

    try:
        fetch_result = fetch_dashboard_data()
        if isinstance(fetch_result, dict):
            screenshot_paths = fetch_result.get("screenshot_paths", []) or []
            if fetch_result.get("screenshot_path"):
                screenshot_paths.append(fetch_result["screenshot_path"])
    except Exception as e:
        extraction_note = f"Не удалось полностью выполнить Playwright-сценарий: {e}"

    try:
        extracted = extract_dashboard_data()
    except Exception as e:
        extracted = {"items": [], "source": "fallback"}
        if extraction_note:
            extraction_note += f"\nТакже не удалось извлечь данные: {e}"
        else:
            extraction_note = f"Не удалось извлечь данные: {e}"

    items = normalize_items(extracted.get("items", []))
    summary = build_summary(items)
    public_message = build_public_message(summary, items)
    personal_messages = build_personal_messages(items)
    missing_data_issues = build_missing_data_issues(items)
    report_text = build_report_text(summary, items)

    screenshot_paths = list(dict.fromkeys([x for x in screenshot_paths if x]))
    screenshot_path = screenshot_paths[0] if screenshot_paths else ""

    result = {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "public_message": public_message,
        "report_text": report_text,
        "summary": summary,
        "items": items,
        "screenshot_path": screenshot_path,
        "screenshot_paths": screenshot_paths,
        "missing_data_issues": missing_data_issues,
        "personal_messages": personal_messages,
        "extraction_note": extraction_note,
        "redmine_url": "",
    }

    save_json(DATA_FILE, result)
    return result


def load_overdue_result():
    return load_json(DATA_FILE, default=None)