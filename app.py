from pathlib import Path
import json
import os
import subprocess
import sys
import threading
import re
import asyncio
import time
from typing import Any
from datetime import datetime
from dotenv import load_dotenv
load_dotenv()


import bcrypt
from services.auth import registration
from services.auth import mailer
from services.notifications import store as notif_store


# ── Keycloak OIDC (активируется через AUTH_PROVIDER=keycloak в .env) ──
import os as _kc_os  # noqa: E402
KC_SERVER_URL = _kc_os.getenv("KC_SERVER_URL", "http://localhost:8080")
KC_REALM = _kc_os.getenv("KC_REALM", "neurona")
KC_CLIENT_ID = _kc_os.getenv("KC_CLIENT_ID", "neurona-web")
KC_CLIENT_SECRET = _kc_os.getenv("KC_CLIENT_SECRET", "")
AUTH_PROVIDER = _kc_os.getenv("AUTH_PROVIDER", "local")

from pydantic import BaseModel  # <-- Добавлен отсутствовавший импорт
from fastapi import FastAPI, Request, Header, HTTPException, Form
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, FileResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from jinja2 import ChainableUndefined


from services import run_history
from services.auth.security import (
    authenticate_user,
    create_access_token,
    get_user_from_token,
    require_admin_user,
    ADMIN_ROLES,
    has_module_access,
)

from services.appeals.storage import (
    create_appeal,
    list_appeals,
    get_appeal,
    update_appeal,
    append_appeal_history,
    calculate_stats as calculate_appeals_stats,
)
from services.appeals.reply_builder import DEFAULT_REPLY_TEMPLATE, generate_reply_from_template



BASE_DIR = Path(__file__).resolve().parent

GIT_UPDATE_TOKEN = os.getenv("GIT_UPDATE_TOKEN", "12345")
GIT_REMOTE_NAME = os.getenv("GIT_REMOTE_NAME", "origin")

DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"
GENERATED_DIR = BASE_DIR / "generated"

EDO_DATA_DIR = DATA_DIR / "edo"
OVERDUE_DATA_DIR = DATA_DIR / "overdue"
WATERCONTROL_DATA_DIR = DATA_DIR / "watercontrol"
UTNKR_DATA_DIR = DATA_DIR / "utnkr"
CAMERAS_DATA_DIR = DATA_DIR / "cameras"

EDO_RESULT_FILE = EDO_DATA_DIR / "result.json"
OVERDUE_RESULT_FILE = OVERDUE_DATA_DIR / "final_result.json"

MGKH_RM_DATA_DIR = DATA_DIR / "mgkh_rm"
MGKH_RM_RESULT_FILE = MGKH_RM_DATA_DIR / "result.json"
WATERCONTROL_RESULT_FILE = WATERCONTROL_DATA_DIR / "result.json"
UTNKR_RESULT_FILE = UTNKR_DATA_DIR / "violators.json"

CAMERAS_ADDRESSES_FILE = CAMERAS_DATA_DIR / "addresses.tsv"
CAMERAS_STATE_FILE = CAMERAS_DATA_DIR / "state" / "dashboard_state.json"
RUN_TIMES_FILE = DATA_DIR / "run_times.json"

GENERATED_PRESCRIPTIONS_DIR = GENERATED_DIR / "prescriptions"

UTNKR_UI_CONFIG = {
    "overdue_column_name": "Просрочка, дней",
    "object_column_name": "Объект",
    "municipality_column_name": "Муниципалитет",
    "organization_column_name": "Организация",
    "responsible_column_name": "Ответственный",
    "status_column_name": "Статус",
    "source_name": "Технадзор УТНКР",
    "source_url": "https://tehnadzor.utnkr.ru/",
}

CAMERAS_UI_CONFIG = {
    "source_name": "Проверка камер",
    "source_url": "https://fkr.eiasmo.ru",
    "addresses_file_name": "addresses.tsv",
    "prescriptions_dir": "generated/prescriptions",
}

# =========================
# РОЛИ С ПОЛНЫМ ДОСТУПОМ
# =========================

# Роли, которые видят ВСЕ модули и админ-интерфейс.
# Хочешь ещё одну — просто допиши название в нижнем регистре.
FULL_ACCESS_ROLES = {"администратор", "руководитель", "пользователь"}

# Полный список модулей платформы
ALL_MODULE_IDS = [
    "edo", "overdue", "watercontrol", "utnkr", "cameras", "appeals",
    "cds", "mgkh_rm", "ecur", "municipality-report", "water-dashboard",
    "water_rm", "tools",
]



# ── Маппинг ролей Keycloak → модули платформы ──
# Роль admin даёт полный доступ. Остальные роли = имя модуля.
KC_ROLE_TO_MODULE = {
    "admin":             "__full_access__",
    "администратор":     "__full_access__",
    "руководитель":      "__full_access__",
    "edo":               "edo",
    "overdue":           "overdue",
    "watercontrol":      "watercontrol",
    "utnkr":             "utnkr",
    "cameras":           "cameras",
    "appeals":           "appeals",
    "cds":               "cds",
    "mgkh_rm":           "mgkh_rm",
    "ecur":              "ecur",
    "municipality-report": "municipality-report",
    "water-dashboard":   "water-dashboard",
    "water_rm":          "water_rm",
    "tools":             "tools",
}

def _kc_user_modules(user: dict) -> list:
    """Собирает список модулей по ролям Keycloak."""
    roles = user.get("roles", []) or []
    role = (user.get("role") or "").strip().lower()
    if role in {"admin", "администратор", "руководитель"}:
        return list(ALL_MODULE_IDS)
    modules = []
    for r in roles:
        mapped = KC_ROLE_TO_MODULE.get(r.lower())
        if mapped and mapped != "__full_access__" and mapped not in modules:
            modules.append(mapped)
    return modules

def is_full_access(user) -> bool:
    role = (user.get("role") or "").strip().lower()
    if role in FULL_ACCESS_ROLES:
        return True
    if user.get("kc_sub") and role in {"admin", "администратор", "руководитель"}:
        return True
    return False


def check_module_access(user, module_id) -> bool:
    """Полные роли проходят на любой модуль, остальные — стандартная проверка."""
    if is_full_access(user):
        return True
    return has_module_access(user, module_id)


def effective_modules(user) -> list:
    """Модули для главной. Поддерживает локальных и Keycloak-пользователей."""
    if is_full_access(user):
        return list(ALL_MODULE_IDS)
    if user.get("kc_sub"):
        # Keycloak-пользователь — собираем модули из его ролей
        return _kc_user_modules(user) or ALL_MODULE_IDS  # fallback: показать всё
    return user.get("modules", [])


async def require_admin_or_full(request: Request):
    """Dependency для роутеров, доступных админу и полным ролям."""
    user = getattr(request.state, "user", None)
    if not user:
        user = get_user_from_token(request.cookies.get("access_token"))
    if not user:
        raise HTTPException(status_code=401, detail="Требуется авторизация")
    role = (user.get("role") or "").strip().lower()
    if role in ADMIN_ROLES or role in FULL_ACCESS_ROLES:
        return user
    raise HTTPException(status_code=403, detail="Недостаточно прав")

app = FastAPI(title="Unified Dashboard")

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
app.mount("/data", StaticFiles(directory=str(DATA_DIR)), name="data")

GENERATED_DIR.mkdir(parents=True, exist_ok=True)

app.mount(
    "/generated",
    StaticFiles(directory=str(GENERATED_DIR)),
    name="generated",
)
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))
templates.env.undefined = ChainableUndefined


run_status = {
    "edo": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выполнению проверки EDO.",
        "last_error": "",
    },
    "overdue": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выполнению проверки просроченных задач.",
        "last_error": "",
    },
    "watercontrol": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выполнению проверки WaterControl.",
        "last_error": "",
    },
    "utnkr": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выполнению проверки УТНКР.",
        "last_error": "",
    },
    "cameras": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выполнению проверки камер.",
        "last_error": "",
    },
    "camera_prescriptions": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к формированию предписаний по камерам.",
        "last_error": "",
    },
    "mgkh_rm": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к проверке дат в задачах Redmine.",
        "last_error": "",
    },
    "water_dashboard": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к обновлению сводного дашборда.",
        "last_error": "",
    },
    "cds": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выгрузке обращений ЦДС.",
        "last_error": "",
    },
}


def load_json_file(path: Path, default=None):
    if not path.exists():
        return default

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"Ошибка чтения JSON {path}: {e}")
        return default


def save_json_file(path: Path, data: Any):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )


def normalize_text(value):
    return (value or "").strip()


def normalize_key(municipality, organization):
    return (
        normalize_text(municipality).upper(),
        normalize_text(organization).lower(),
    )


def to_int(value, default=0):
    try:
        if value is None:
            return default

        if isinstance(value, str):
            value = (
                value.strip()
                .replace(" ", "")
                .replace("\u00A0", "")
                .replace(",", ".")
            )
            if not value:
                return default

        return int(float(value))
    except Exception:
        return default


def as_list_from_result(result):
    if not result:
        return []

    if isinstance(result, list):
        return result

    if isinstance(result, dict):
        for key in (
            "rows",
            "items",
            "violators",
            "cameras",
            "results",
            "data",
            "records",
        ):
            value = result.get(key)
            if isinstance(value, list):
                return value

    return []


def ensure_personal_message_flags(result, result_file: Path):
    if not result or not isinstance(result, dict):
        return result

    personal_messages = result.get("personal_messages", []) or []
    changed = False

    for item in personal_messages:
        if isinstance(item, dict) and "is_edited" not in item:
            item["is_edited"] = False
            changed = True

    if changed:
        try:
            save_json_file(result_file, result)
        except Exception as e:
            print("Ошибка при обновлении is_edited:", e)

    return result


def calculate_standard_status_metrics(result):
    rows = as_list_from_result(result)

    total = len(rows)
    critical = 0
    risk = 0
    ok = 0

    for row in rows:
        if not isinstance(row, dict):
            continue

        status = normalize_text(row.get("status")).lower()

        if status in ("critical", "red", "критично", "красный"):
            critical += 1
        elif status in ("risk", "warning", "yellow", "риск", "желтый", "жёлтый"):
            risk += 1
        elif status in ("ok", "green", "success", "норма", "зеленый", "зелёный"):
            ok += 1

    return {
        "total": total,
        "critical": critical,
        "risk": risk,
        "ok": ok,
    }


def calculate_edo_metrics(result):
    return calculate_standard_status_metrics(result)


def calculate_watercontrol_metrics(result):
    return calculate_standard_status_metrics(result)


def calculate_mgkh_rm_metrics(result):
    if not result or not isinstance(result, dict):
        return {"total": 0, "close": 0, "extend": 0, "rework": 0}

    b = result.get("buckets")
    if isinstance(b, dict):
        close = len(b.get("close", []) or [])
        extend = len(b.get("extend", []) or [])
        rework = len(b.get("rework", []) or [])
        return {
            "total": close + extend + rework,
            "close": close,
            "extend": extend,
            "rework": rework,
        }

    m = result.get("metrics")
    if isinstance(m, dict):
        return {
            "total": to_int(m.get("total")),
            "close": to_int(m.get("close")),
            "extend": to_int(m.get("extend")),
            "rework": to_int(m.get("rework")),
        }

    return {"total": 0, "close": 0, "extend": 0, "rework": 0}


def calculate_overdue_metrics(raw_result):
    if not raw_result or not isinstance(raw_result, dict):
        return {"total": 0, "critical": 0, "risk": 0, "ok": 0}

    items = raw_result.get("items", []) or []

    total = len(items)
    critical = len([x for x in items if to_int(x.get("overdue_count", 0)) >= 20])
    risk = len([x for x in items if 1 <= to_int(x.get("overdue_count", 0)) < 20])
    ok = len([x for x in items if to_int(x.get("overdue_count", 0)) <= 0])

    return {
        "total": total,
        "critical": critical,
        "risk": risk,
        "ok": ok,
    }


def calculate_utnkr_metrics(result):
    rows = as_list_from_result(result)

    total = len(rows)
    critical = 0
    risk = 0
    ok = 0

    for row in rows:
        if not isinstance(row, dict):
            continue

        status = normalize_text(
            row.get("status")
            or row.get("traffic_light")
            or row.get("color")
            or row.get("level")
        ).lower()

        overdue_days = to_int(
            row.get("overdue_days")
            or row.get("days_overdue")
            or row.get("days")
            or row.get("delay_days"),
            0,
        )

        if status in ("critical", "red", "критично", "красный"):
            critical += 1
        elif status in ("risk", "warning", "yellow", "риск", "желтый", "жёлтый"):
            risk += 1
        elif status in ("ok", "green", "success", "норма", "зеленый", "зелёный"):
            ok += 1
        elif overdue_days >= 20:
            critical += 1
        elif overdue_days > 0:
            risk += 1
        else:
            ok += 1

    return {
        "total": total,
        "critical": critical,
        "risk": risk,
        "ok": ok,
    }


def normalize_camera_status_value(row):
    if not isinstance(row, dict):
        return "unknown"

    raw_status = normalize_text(
        row.get("camera_status")
        or row.get("status")
        or row.get("stream_status")
        or row.get("check_status")
        or row.get("state")
    ).lower()

    online = row.get("online")
    available = row.get("available")
    has_stream = row.get("has_stream")
    stream_url = normalize_text(row.get("stream_url") or row.get("link_url"))

    if raw_status in (
        "working", "ok", "online", "success", "active", "green",
        "норма", "работает", "активна", "активный",
    ):
        return "working"

    if raw_status in (
        "not_working", "critical", "offline", "error", "failed", "fail",
        "broken", "red", "критично", "не работает", "неработает",
        "ошибка", "отключена", "офлайн",
    ):
        return "not_working"

    if raw_status in (
        "not_connected", "no_stream", "missing_stream", "no_url", "empty_url",
        "не подключена", "не подключен", "нет ссылки", "нет потока", "без ссылки",
    ):
        return "not_connected"

    if online is False or available is False:
        return "not_working"

    if has_stream is False:
        return "not_connected"

    if online is True or available is True:
        return "working"

    if not stream_url:
        return "not_connected"

    return "unknown"


def normalize_camera_row(row):
    if not isinstance(row, dict):
        row = {}

    item = dict(row)
    item["camera_status"] = normalize_camera_status_value(item)

    if not item.get("city"):
        item["city"] = item.get("municipality", "")

    if not item.get("owner"):
        item["owner"] = item.get("responsible", "") or item.get("organization", "")

    if not item.get("link_url") and item.get("stream_url"):
        item["link_url"] = item.get("stream_url")

    return item


def calculate_cameras_metrics(result):
    rows = [normalize_camera_row(row) for row in as_list_from_result(result)]

    total = len(rows)
    working = 0
    not_working = 0
    not_connected = 0
    unknown = 0

    for row in rows:
        status = row.get("camera_status")

        if status == "working":
            working += 1
        elif status == "not_working":
            not_working += 1
        elif status == "not_connected":
            not_connected += 1
        else:
            unknown += 1

    return {
        "total": total,
        "working": working,
        "not_working": not_working,
        "not_connected": not_connected,
        "unknown": unknown,
        "ok": working,
        "critical": not_working,
        "risk": not_connected + unknown,
    }


def filter_camera_rows(rows, status_filter="all"):
    status_filter = status_filter or "all"

    if status_filter == "all":
        return rows

    if status_filter == "problem":
        return [row for row in rows if row.get("camera_status") != "working"]

    return [row for row in rows if row.get("camera_status") == status_filter]


def group_camera_rows(rows, group_by="none"):
    group_by = group_by or "none"

    if group_by == "none":
        return []

    groups = {}

    for row in rows:
        if group_by == "company":
            key = row.get("owner") or "Не указано"
        elif group_by == "city":
            key = row.get("city") or row.get("municipality") or "Не указано"
        else:
            key = "Без группировки"

        groups.setdefault(key, []).append(row)

    return [
        {
            "name": name,
            "items": items,
        }
        for name, items in sorted(groups.items(), key=lambda item: item[0])
    ]


def transform_overdue_result_for_ui(raw):
    if not raw or not isinstance(raw, dict):
        return None

    summary = raw.get("summary", {}) or {}
    items = raw.get("items", []) or []

    rows = []

    for item in items:
        overdue_count = to_int(item.get("overdue_count", 0), 0)

        if overdue_count >= 20:
            status = "critical"
            reason = "Высокое количество просроченных задач"
        elif overdue_count > 0:
            status = "risk"
            reason = "Есть просроченные задачи"
        else:
            status = "ok"
            reason = "Просроченные задачи отсутствуют"

        rows.append(
            {
                "municipality": item.get("municipality", ""),
                "organization": item.get("organization", item.get("municipality", "")),
                "responsible_name": item.get("responsible_name", ""),
                "responsible_phone": item.get("responsible_phone", ""),
                "status": status,
                "reason": reason,
                "overdue_count": overdue_count,
            }
        )

    rows.sort(
        key=lambda x: (
            -to_int(x.get("overdue_count", 0)),
            x.get("municipality", ""),
        )
    )

    return {
        "created_at": raw.get("created_at", ""),
        "summary_message": raw.get("public_message", "Сводка пока не сформирована."),
        "public_chat_message": raw.get("public_message", ""),
        "rows": rows,
        "screenshot_path": raw.get("screenshot_path", ""),
        "screenshot_paths": raw.get("screenshot_paths", []),
        "missing_data_issues": raw.get("missing_data_issues", []),
        "personal_messages": raw.get("personal_messages", []),
        "extraction_note": raw.get("extraction_note", ""),
        "redmine_url": raw.get("redmine_url", ""),
        "report_text": raw.get("report_text", ""),
        "summary": summary,
        "by_status": summary.get("by_status", []),
        "by_municipality": summary.get("by_municipality", []),
        "by_category": summary.get("by_category", []),
    }


def set_status(service_name: str, **kwargs):
    if service_name not in run_status:
        return

    for key, value in kwargs.items():
        run_status[service_name][key] = value


def save_run_time(service_name: str):
    times = load_json_file(RUN_TIMES_FILE, default={}) or {}
    st = run_status.get(service_name, {})
    times[service_name] = {
        "started_at": st.get("started_at", ""),
        "finished_at": st.get("finished_at", ""),
    }
    save_json_file(RUN_TIMES_FILE, times)


def load_run_time(service_name: str) -> dict:
    times = load_json_file(RUN_TIMES_FILE, default={}) or {}
    return times.get(service_name, {})


def make_check_state(service_name: str):
    status = run_status.get(service_name, {})
    persisted = load_run_time(service_name)

    started_at = status.get("started_at") or persisted.get("started_at", "")
    finished_at = status.get("finished_at") or persisted.get("finished_at", "")

    return {
        "is_running": bool(status.get("running", False)),
        "running": bool(status.get("running", False)),
        "stage": status.get("stage", "Ожидание запуска"),
        "message": status.get("message", ""),
        "last_error": status.get("last_error", ""),
        "started_at": started_at,
        "finished_at": finished_at,
        "updated_at": finished_at or started_at or "",
    }


def build_table_info(file_path: Path):
    exists = file_path.exists()
    rel_path = str(file_path.relative_to(BASE_DIR)) if exists else ""

    return {
        "exists": exists,
        "path": rel_path,
        "name": file_path.name,
        "size": file_path.stat().st_size if exists else 0,
    }


def run_subprocess_worker(service_name: str, command: list[str], cwd: Path):
    status = run_status[service_name]
    run_id = None

    try:
        run_id = run_history.record_start(service_name, user="—")
        status["running"] = True
        status["last_error"] = ""
        status["started_at"] = datetime.now().isoformat(timespec="seconds")
        status["finished_at"] = ""
        save_run_time(service_name)
        status["stage"] = "Запуск"
        status["message"] = f"Запущена проверка: {service_name}."

        process = subprocess.Popen(
            command,
            cwd=str(cwd),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        if process.stdout:
            for line in process.stdout:
                text = line.strip()

                if not text:
                    continue

                print(f"[{service_name.upper()}]", text)

                if text.startswith("STAGE:"):
                    stage_name = text.replace("STAGE:", "", 1).strip()
                    status["stage"] = stage_name
                    status["message"] = stage_name
                    continue

                lowered = text.lower()

                if "captcha" in lowered or "капча" in lowered:
                    status["stage"] = "Ожидание подтверждения"
                    status["message"] = text
                elif "login" in lowered or "авторизац" in lowered or "вход" in lowered:
                    status["stage"] = "Авторизация"
                    status["message"] = text
                elif "screenshot" in lowered or "скриншот" in lowered:
                    status["stage"] = "Снятие скриншотов"
                    status["message"] = text
                elif "анализ" in lowered or "извлеч" in lowered or "parse" in lowered:
                    status["stage"] = "Анализ данных"
                    status["message"] = text
                elif "camera" in lowered or "кам" in lowered or "stream" in lowered:
                    status["stage"] = "Проверка камер"
                    status["message"] = text
                elif "ffmpeg" in lowered or "ffprobe" in lowered:
                    status["stage"] = "Проверка видеопотоков"
                    status["message"] = text
                elif "docx" in lowered or "предпис" in lowered:
                    status["stage"] = "Формирование документов"
                    status["message"] = text
                elif (
                    "сохранение" in lowered
                    or "result saved" in lowered
                    or "saved" in lowered
                    or "готово:" in lowered
                ):
                    status["stage"] = "Сохранение результата"
                    status["message"] = text
                else:
                    status["message"] = text

        return_code = process.wait()

        if return_code != 0:
            status["running"] = False
            status["stage"] = "Ошибка"
            status["message"] = f"Процесс {service_name} завершился с ошибкой."
            status["last_error"] = f"Процесс завершился с кодом {return_code}"
            if run_id:
                run_history.record_finish(run_id, "error", f"код {return_code}")
            return

        status["running"] = False
        status["stage"] = "Готово"
        status["message"] = f"Процесс {service_name} завершён успешно."
        if run_id:
            run_history.record_finish(run_id, "success")

    except Exception as e:
        status["running"] = False
        status["stage"] = "Ошибка"
        status["message"] = "Во время выполнения произошла ошибка."
        status["last_error"] = str(e)
        if run_id:
            run_history.record_finish(run_id, "error", str(e))

    finally:
        status["running"] = False
        status["finished_at"] = datetime.now().isoformat(timespec="seconds")
        save_run_time(service_name)


def start_background_service(service_name: str, command: list[str]):
    if run_status[service_name]["running"]:
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "message": f"Процесс {service_name} уже выполняется.",
            },
        )

    thread = threading.Thread(
        target=run_subprocess_worker,
        args=(service_name, command, BASE_DIR),
        daemon=True,
    )
    thread.start()

    return {"ok": True}


def save_personal_message_to_file(
    result_file: Path,
    payload: dict,
    not_found_message: str,
    allow_create: bool = False,
):
    municipality = normalize_text(payload.get("municipality"))
    organization = normalize_text(payload.get("organization"))
    message = normalize_text(payload.get("message"))

    if not municipality or not organization or not message:
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "message": "Не хватает данных для сохранения",
            },
        )

    data = load_json_file(result_file)

    if not data:
        return JSONResponse(
            status_code=404,
            content={
                "ok": False,
                "message": not_found_message,
            },
        )

    if not isinstance(data, dict):
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "message": "Некорректный формат файла результата",
            },
        )

    personal_messages = data.get("personal_messages", []) or []
    target_key = normalize_key(municipality, organization)
    updated = False

    for item in personal_messages:
        item_key = normalize_key(
            item.get("municipality", ""),
            item.get("organization", ""),
        )

        if item_key == target_key:
            item["message"] = message
            item["is_edited"] = True
            updated = True
            break

    if not updated:
        if not allow_create:
            return JSONResponse(
                status_code=404,
                content={
                    "ok": False,
                    "message": "Уведомление не найдено",
                },
            )

        personal_messages.append(
            {
                "municipality": municipality,
                "organization": organization,
                "message": message,
                "is_edited": True,
                "status": "risk",
                "responsible_name": "",
                "responsible_phone": "",
            }
        )

    data["personal_messages"] = personal_messages
    save_json_file(result_file, data)

    return {
        "ok": True,
        "message": "Сообщение сохранено.",
    }


def require_git_update_token(x_git_update_token: str | None):
    if not x_git_update_token or x_git_update_token != GIT_UPDATE_TOKEN:
        raise HTTPException(
            status_code=403,
            detail="Недостаточно прав для проверки или загрузки обновлений.",
        )


def run_git_command(command: list[str], timeout: int = 90):
    try:
        result = subprocess.run(
            command,
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            shell=False,
            timeout=timeout,
        )

        return {
            "ok": result.returncode == 0,
            "returncode": result.returncode,
            "command": " ".join(command),
            "stdout": result.stdout,
            "stderr": result.stderr,
        }

    except subprocess.TimeoutExpired:
        return {
            "ok": False,
            "returncode": -1,
            "command": " ".join(command),
            "stdout": "",
            "stderr": "Команда выполнялась слишком долго и была остановлена.",
        }

    except FileNotFoundError:
        return {
            "ok": False,
            "returncode": -1,
            "command": " ".join(command),
            "stdout": "",
            "stderr": "Git не найден. Проверьте, что Git установлен и доступен из терминала.",
        }

    except Exception as e:
        return {
            "ok": False,
            "returncode": -1,
            "command": " ".join(command),
            "stdout": "",
            "stderr": str(e),
        }


def get_current_git_branch():
    result = run_git_command(["git", "branch", "--show-current"], timeout=30)

    if result.get("ok") and result.get("stdout", "").strip():
        return result["stdout"].strip()

    return "main"


def is_git_repository():
    result = run_git_command(["git", "rev-parse", "--is-inside-work-tree"], timeout=30)

    return result.get("ok") and result.get("stdout", "").strip() == "true"


def has_local_git_changes():
    result = run_git_command(["git", "status", "--porcelain"], timeout=30)

    if not result.get("ok"):
        return {
            "ok": False,
            "has_changes": True,
            "status": result,
        }

    return {
        "ok": True,
        "has_changes": bool(result.get("stdout", "").strip()),
        "status": result,
    }


# =========================
# AUTH: КОНФИГ И MIDDLEWARE
# =========================

PATH_MODULE_MAP = {
    "/edo": "edo",
    "/overdue": "overdue",
    "/mgkh-rm": "mgkh_rm",
    "/watercontrol": "watercontrol",
    "/utnkr": "utnkr",
    "/cameras": "cameras",
    "/camera-prescriptions": "cameras",
    "/prescriptions": "cameras",
    "/appeals": "appeals",
    "/cds": "cds",
    "/municipality-report": "municipality-report",
    "/ecur": "ecur",
    "/water-dashboard": "water-dashboard",
}

PUBLIC_PATH_PREFIXES = (
    "/login",
    "/logout",
    "/register",        # страница регистрации
    "/api/register",    # покрывает /verify и /resend
    "/static",
    "/data",
    "/generated",
    "/favicon.ico",
    "/health",
    "/api/system/health",
)




# ── Keycloak: роуты логина (OAuth2 Authorization Code Flow) ──
@app.get("/login/keycloak")
async def login_keycloak_redirect(request: Request):
    redirect_uri = str(request.url_for("login_keycloak_callback"))
    auth_url = (
        f"{KC_SERVER_URL}/realms/{KC_REALM}/protocol/openid-connect/auth"
        f"?client_id={KC_CLIENT_ID}&response_type=code"
        f"&redirect_uri={redirect_uri}&scope=openid+profile+email"
    )
    return RedirectResponse(auth_url, status_code=302)


@app.get("/login/keycloak/callback", name="login_keycloak_callback")
async def login_keycloak_callback(request: Request, code: str = ""):
    if not code:
        return RedirectResponse("/login?error=no_code", status_code=302)
    try:
        import httpx
        redirect_uri = str(request.url_for("login_keycloak_callback"))
        token_url = f"{KC_SERVER_URL}/realms/{KC_REALM}/protocol/openid-connect/token"
        async with httpx.AsyncClient(timeout=15.0) as client:
            resp = await client.post(token_url, data={
                "grant_type": "authorization_code",
                "code": code,
                "redirect_uri": redirect_uri,
                "client_id": KC_CLIENT_ID,
                "client_secret": KC_CLIENT_SECRET,
            })
        if resp.status_code != 200:
            print(f"[keycloak] token exchange failed: {resp.status_code} {resp.text[:200]}")
            return RedirectResponse("/login?error=token_exchange", status_code=302)
        tokens = resp.json()
        access_token = tokens.get("access_token", "")
        refresh_token = tokens.get("refresh_token", "")
        expires_in = tokens.get("expires_in", 3600)
        response = RedirectResponse("/", status_code=302)
        response.set_cookie("access_token", access_token, httponly=True, max_age=expires_in, samesite="lax")
        if refresh_token:
            response.set_cookie("refresh_token", refresh_token, httponly=True, max_age=7 * 86400, samesite="lax")
        response.set_cookie("auth_provider", "keycloak", max_age=expires_in, samesite="lax")
        return response
    except Exception as e:
        print(f"[keycloak] login error: {e}")
        return RedirectResponse("/login?error=exception", status_code=302)


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    path = request.url.path

    if any(path.startswith(p) for p in PUBLIC_PATH_PREFIXES):
        return await call_next(request)

    token = request.cookies.get("access_token")
    user = get_user_from_token(token)

    if not user:
        if path == "/" or request.method == "GET":
            return RedirectResponse(url="/login", status_code=302)
        return JSONResponse(
            status_code=401,
            content={"ok": False, "message": "Требуется авторизация"},
        )

    for path_prefix, module_id in PATH_MODULE_MAP.items():
        if path.startswith(path_prefix):
            # ← ЗДЕСЬ единственная замена: has_module_access → check_module_access
            if not check_module_access(user, module_id):
                return RedirectResponse(url="/?error=no_access", status_code=302)
            break

    request.state.user = user
    return await call_next(request)

# =========================
# SECURITY: RATE-LIMIT + ЗАГОЛОВКИ
# =========================

LOGIN_ATTEMPTS: dict = {}


def rate_limit_ok(key: str, limit: int = 5, window: int = 300) -> bool:
    """Не более limit попыток за window секунд для одного key."""
    now = time.time()
    rec = LOGIN_ATTEMPTS.setdefault(key, [])
    rec[:] = [t for t in rec if now - t < window]
    if len(rec) >= limit:
        return False
    rec.append(now)
    return True


@app.middleware("http")
async def security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Permissions-Policy"] = "geolocation=(), camera=(), microphone=()"
    return response

# =========================
# AUTH: РОУТЫ ВХОДА И ВЫХОДА
# =========================

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, error: str = "", message: str = ""):
    token = request.cookies.get("access_token")
    if get_user_from_token(token):
        return RedirectResponse(url="/", status_code=302)

    return templates.TemplateResponse(
        request,
        "login.html",
        {
            "request": request,
            "error": error,
            "message": message,
        },
    )

@app.post("/login")
async def login_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
):
    client_ip = request.client.host if request.client else "unknown"
    rl_key = f"{client_ip}:{username.lower()}"

    if not rate_limit_ok(rl_key):
        return RedirectResponse(
            url="/login?error=Слишком много попыток входа. Подождите 5 минут.",
            status_code=303,
        )

    user = authenticate_user(username, password)
    if not user:
        return RedirectResponse(
            url="/login?error=Неверный логин или пароль",
            status_code=303,
        )

    access_token = create_access_token({
        "sub": user["username"],
        "role": user.get("role", ""),
    })

    response = RedirectResponse(url="/", status_code=303)
    response.set_cookie(
        key="access_token",
        value=access_token,
        httponly=True,          # кука недоступна из JavaScript (защита от XSS-кражи)
        max_age=60 * 60 * 8,
        samesite="strict",      # кука не отправляется с чужих сайтов (защита от CSRF)
        secure=False,           # остаёмся на HTTP
    )
    return response

@app.get("/logout")
async def logout():
    response = RedirectResponse(url="/login?message=Вы вышли из системы", status_code=303)
    response.delete_cookie("access_token")
    response.delete_cookie("auth_provider")
    return response


@app.get("/", response_class=HTMLResponse)
async def home(request: Request, error: str = ""):
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}

    return templates.TemplateResponse(
        request,
        "home.html",
        {
            "request": request,
            "user": user,
            "user_modules": effective_modules(user),
            "is_admin_ui": user.get("username") == "admin" or is_full_access(user),
            "user_role": user.get("role", ""),
            "user_username": user.get("username", ""),
            "access_error": error,
        },
    )


@app.get("/register", response_class=HTMLResponse)
async def register_page(request: Request):
    return templates.TemplateResponse(request, "register.html", {"request": request})


@app.post("/api/register")
async def api_register_start(payload: dict):
    try:
        username = normalize_text(payload.get("username"))
        email = normalize_text(payload.get("email"))
        password = str(payload.get("password") or "")

        if len(password) < 6:
            return JSONResponse(status_code=400,
                                content={"ok": False, "message": "Пароль должен быть не короче 6 символов"})

        password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        print(f"[register] Старт регистрации: username={username}, email={email}")
        
        ok, result = registration.start_registration(username, email, password_hash)
        if not ok:
            print(f"[register] Ошибка валидации: {result}")
            return JSONResponse(status_code=400, content={"ok": False, "message": result})

        print(f"[register] Заявка создана, отправляю код на {result['email']}")
        sent, err = mailer.send_verification_code(result["email"], result["code"])
        if not sent:
            print(f"[register] Ошибка отправки письма: {err}")
            if os.getenv("REGISTRATION_DEBUG_CODE") == "1":
                return {"ok": True, "debug_code": result["code"],
                        "message": "SMTP не настроен (dev-режим). Код: " + result["code"]}
            return JSONResponse(status_code=500, content={
                "ok": False,
                "message": f"Не удалось отправить письмо. Проверьте YANDEX_SMTP_USER / YANDEX_SMTP_PASSWORD. ({err})",
            })

        print(f"[register] Письмо отправлено успешно")
        return {"ok": True, "message": f"Код отправлен на {result['email']}"}
    except Exception as e:
        import traceback
        print("[register] НЕПРЕДВИДЕННАЯ ОШИБКА:")
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"ok": False, "message": f"Ошибка сервера: {str(e)}"})

    password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    ok, result = registration.start_registration(username, email, password_hash)
    if not ok:
        return JSONResponse(status_code=400, content={"ok": False, "message": result})

    sent, err = mailer.send_verification_code(result["email"], result["code"])
    if not sent:
        if os.getenv("REGISTRATION_DEBUG_CODE") == "1":
            return {"ok": True, "debug_code": result["code"],
                    "message": "SMTP не настроен (dev-режим). Код: " + result["code"]}
        return JSONResponse(status_code=500, content={
            "ok": False,
            "message": f"Не удалось отправить письмо. Проверьте YANDEX_SMTP_USER / YANDEX_SMTP_PASSWORD. ({err})",
        })

    return {"ok": True, "message": f"Код отправлен на {result['email']}"}


@app.post("/api/register/resend")
async def api_register_resend(payload: dict):
    ident = normalize_text(payload.get("username") or payload.get("email"))
    ok, result = registration.resend_code(ident)
    if not ok:
        return JSONResponse(status_code=400, content={"ok": False, "message": result})

    sent, err = mailer.send_verification_code(result["email"], result["code"])
    if not sent:
        if os.getenv("REGISTRATION_DEBUG_CODE") == "1":
            return {"ok": True, "debug_code": result["code"], "message": "Dev-режим. Код: " + result["code"]}
        return JSONResponse(status_code=500, content={"ok": False, "message": f"Ошибка отправки: {err}"})

    return {"ok": True, "message": f"Письмо отправлено на {result['email']}"}


@app.post("/api/register/verify")
async def api_register_verify(payload: dict):
    ident = normalize_text(payload.get("username") or payload.get("email"))
    code = normalize_text(payload.get("code"))

    ok, result = registration.verify_registration(ident, code)
    if not ok:
        return JSONResponse(status_code=400, content={"ok": False, "message": result})

    # Автовход: выдаём cookie, как при логине
    access_token = create_access_token({"sub": result["username"], "role": result["role"]})
    response = JSONResponse(content={"ok": True, "message": "Аккаунт создан! Добро пожаловать.", "redirect": "/"})
    response.set_cookie(key="access_token", value=access_token, httponly=True,
                        max_age=60 * 60 * 8, samesite="strict", secure=False)
    return response


@app.get("/api/me/settings")
async def api_me_settings(request: Request):
    return {"ok": True, "settings": registration.get_settings(request.state.user["id"])}


@app.post("/api/me/settings")
async def api_me_save_settings(request: Request, payload: dict):
    settings = registration.save_settings(request.state.user["id"], payload.get("settings") or {})
    return {"ok": True, "settings": settings}

@app.get("/edo", response_class=HTMLResponse)
async def edo_page(request: Request):
    result = load_json_file(EDO_RESULT_FILE)
    result = ensure_personal_message_flags(result, EDO_RESULT_FILE)
    metrics = calculate_edo_metrics(result)

    return templates.TemplateResponse(
        request,
        "edo.html",
        {
            "request": request,
            "result": result,
            "metrics": metrics,
            "status": run_status["edo"],
            "check_state": make_check_state("edo"),
            "run_status": run_status["edo"],
        },
    )


@app.get("/mgkh-rm", response_class=HTMLResponse)
async def mgkh_rm_page(request: Request):
    result = load_json_file(MGKH_RM_RESULT_FILE)
    metrics = calculate_mgkh_rm_metrics(result)

    return templates.TemplateResponse(
        request,
        "mgkh_rm.html",
        {
            "request": request,
            "result": result,
            "metrics": metrics,
            "status": run_status["mgkh_rm"],
            "check_state": make_check_state("mgkh_rm"),
            "run_status": run_status["mgkh_rm"],
        },
    )


@app.get("/mgkh-rm/run-status")
async def mgkh_rm_run_status():
    return run_status["mgkh_rm"]


@app.post("/mgkh-rm/run-check")
async def mgkh_rm_run_check():
    command = [sys.executable, "-m", "services.mgkh_php.runner"]
    return start_background_service("mgkh_rm", command)


@app.get("/scheduler", response_class=HTMLResponse)
async def scheduler_page(request: Request):
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}

    role = (user.get("role") or "").strip().lower()
    if role not in ADMIN_ROLES and role not in FULL_ACCESS_ROLES:
        return RedirectResponse(url="/", status_code=303)

    return templates.TemplateResponse(
        request,
        "scheduler.html",
        {
            "request": request,
            "user": user,
            "user_role": user.get("role", ""),
            "user_username": user.get("username", ""),
        },
    )
@app.get("/overdue", response_class=HTMLResponse)
async def overdue_page(request: Request):
    raw_result = load_json_file(OVERDUE_RESULT_FILE)
    raw_result = ensure_personal_message_flags(raw_result, OVERDUE_RESULT_FILE)

    result = transform_overdue_result_for_ui(raw_result)
    metrics = calculate_overdue_metrics(raw_result)

    return templates.TemplateResponse(
        request,
        "overdue.html",
        {
            "request": request,
            "result": result,
            "raw_result": raw_result,
            "metrics": metrics,
            "status": run_status["overdue"],
            "check_state": make_check_state("overdue"),
            "run_status": run_status["overdue"],
        },
    )


@app.get("/watercontrol", response_class=HTMLResponse)
async def watercontrol_page(request: Request):
    result = load_json_file(WATERCONTROL_RESULT_FILE)
    result = ensure_personal_message_flags(result, WATERCONTROL_RESULT_FILE)
    metrics = calculate_watercontrol_metrics(result)

    return templates.TemplateResponse(
        request,
        "watercontrol.html",
        {
            "request": request,
            "result": result,
            "metrics": metrics,
            "status": run_status["watercontrol"],
            "check_state": make_check_state("watercontrol"),
            "run_status": run_status["watercontrol"],
        },
    )

@app.get("/water-dashboard", response_class=HTMLResponse)
async def water_dashboard_page(request: Request):
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}

    snap = load_json_file(BASE_DIR / "data" / "water_dashboard" / "snapshot.json", default={}) or {}

    return templates.TemplateResponse(
        request,
        "water_dashboard.html",
        {
            "request": request,
            "user": user,
            "user_role": user.get("role", ""),
            "user_username": user.get("username", ""),
            "snapshot_date": snap.get("snapshot_date", "—"),
            "table_json": json.dumps(snap.get("table", []), ensure_ascii=False),
            "meta_json": json.dumps({
                "updated_at": snap.get("updated_at", ""),
                "refresh": snap.get("sources_refresh", {}),
                "kpi_live": snap.get("kpi_live", {}),
            }, ensure_ascii=False),
        },
    )
@app.get("/utnkr", response_class=HTMLResponse)
async def utnkr_page(request: Request):
    result = load_json_file(UTNKR_RESULT_FILE, default={})
    rows = as_list_from_result(result)
    metrics = calculate_utnkr_metrics(result)

    return templates.TemplateResponse(
        request,
        "utnkr.html",
        {
            "request": request,
            "result": result,
            "data": result,
            "rows": rows,
            "items": rows,
            "violators": rows,
            "metrics": metrics,
            "config": UTNKR_UI_CONFIG,
            "check_state": make_check_state("utnkr"),
            "status": run_status["utnkr"],
            "run_status": run_status["utnkr"],
        },
    )


@app.get("/cameras", response_class=HTMLResponse)
async def cameras_page(
    request: Request,
    status: str = "all",
    group_by: str = "none",
):
    state = load_json_file(CAMERAS_STATE_FILE, default={}) or {}
    cameras = [normalize_camera_row(row) for row in as_list_from_result(state)]

    filtered_results = filter_camera_rows(cameras, status)
    grouped_results = group_camera_rows(filtered_results, group_by)
    metrics = calculate_cameras_metrics({"rows": cameras})
    table_info = build_table_info(CAMERAS_ADDRESSES_FILE)

    return templates.TemplateResponse(
        request,
        "cameras.html",
        {
            "request": request,
            "result": state,
            "state": state,
            "data": state,
            "rows": cameras,
            "items": cameras,
            "cameras": cameras,
            "results": cameras,
            "filtered_results": filtered_results,
            "grouped_results": grouped_results,
            "metrics": metrics,
            "config": CAMERAS_UI_CONFIG,
            "check_state": make_check_state("cameras"),
            "prescription_state": make_check_state("camera_prescriptions"),
            "status": run_status["cameras"],
            "run_status": run_status["cameras"],
            "status_filter": status,
            "group_by": group_by,
            "addresses_exists": table_info["exists"],
            "addresses_file": table_info["path"],
            "table_info": table_info,
        },
    )


@app.get("/prescriptions", response_class=HTMLResponse)
async def prescriptions_alias(request: Request):
    return RedirectResponse(url="/camera-prescriptions", status_code=302)


@app.get("/camera-prescriptions", response_class=HTMLResponse)
async def camera_prescriptions_page(
    request: Request,
    mode: str = "individual",
):
    allowed_modes = {"individual", "combined", "zip", "all"}

    if mode not in allowed_modes:
        mode = "individual"

    state = load_json_file(CAMERAS_STATE_FILE, default={}) or {}
    camera_rows = [normalize_camera_row(row) for row in as_list_from_result(state)]

    problem_rows = [
        row for row in camera_rows
        if row.get("camera_status") in ("not_working", "not_connected")
    ]

    prescriptions = []
    generated_items = []

    combined_filename = None
    combined_url = None

    zip_filename = None
    zip_url = None

    generation_result_file = GENERATED_PRESCRIPTIONS_DIR / "generation_result.json"
    generation_result = load_json_file(generation_result_file, default={}) or {}

    if isinstance(generation_result, dict):
        result_items = generation_result.get("items", []) or []

        if isinstance(result_items, list):
            generated_items = result_items

        combined_filename = generation_result.get("combined_filename")
        combined_url = generation_result.get("combined_url")

        zip_filename = generation_result.get("zip_filename")
        zip_url = generation_result.get("zip_url")

    if GENERATED_PRESCRIPTIONS_DIR.exists():
        all_files = sorted(
            GENERATED_PRESCRIPTIONS_DIR.glob("*"),
            key=lambda item: item.stat().st_mtime if item.exists() else 0,
            reverse=True,
        )

        for path in all_files:
            if not path.is_file():
                continue

            if path.name == "generation_result.json":
                continue

            file_info = {
                "name": path.name,
                "filename": path.name,
                "path": f"/generated/prescriptions/{path.name}",
                "url": f"/generated/prescriptions/{path.name}",
                "size": path.stat().st_size,
            }

            prescriptions.append(file_info)

            lower_name = path.name.lower()

            if not zip_filename and lower_name.endswith(".zip"):
                zip_filename = path.name
                zip_url = f"/generated/prescriptions/{path.name}"

            if not combined_filename and lower_name.endswith(".docx") and (
                "combined" in lower_name
                or "общ" in lower_name
                or "all" in lower_name
            ):
                combined_filename = path.name
                combined_url = f"/generated/prescriptions/{path.name}"

    def parse_generated_prescription_filename(filename):
        stem = Path(filename).stem
        stem_without_number = re.sub(r"^\d+[_-]+", "", stem)

        reason_map = [
            {
                "prefixes": ("no_video_stream", "no_stream", "missing_stream"),
                "label": "Нет видеопотока",
                "status": "not_connected",
            },
            {
                "prefixes": ("not_connected", "camera_not_connected", "no_camera"),
                "label": "Камера отсутствует / не подключена",
                "status": "not_connected",
            },
            {
                "prefixes": ("not_working", "camera_not_working", "offline", "camera_offline"),
                "label": "Камера не работает",
                "status": "not_working",
            },
            {
                "prefixes": ("broken", "error", "failed", "fail"),
                "label": "Ошибка работы камеры",
                "status": "not_working",
            },
        ]

        reason_label = "Статус не определён"
        camera_status = ""
        rest = stem_without_number

        for item in reason_map:
            for prefix in item["prefixes"]:
                if rest.lower().startswith(prefix.lower() + "_"):
                    reason_label = item["label"]
                    camera_status = item["status"]
                    rest = rest[len(prefix) + 1:]
                    break

                if rest.lower() == prefix.lower():
                    reason_label = item["label"]
                    camera_status = item["status"]
                    rest = ""

            if camera_status:
                break

        def cleanup(value):
            value = str(value or "")
            value = value.replace("__", " ").replace("_", " ")
            value = re.sub(r"\s+", " ", value)
            value = re.sub(r"\s+,", ",", value)
            value = re.sub(r",\s*", ", ", value)
            value = value.strip(" -_.,;")

            replacements = [
                (r"\bг\s+", "г. "),
                (r"\bул\s+", "ул. "),
                (r"\bд\s+", "д. "),
                (r"\bмкр\s+", "мкр. "),
                (r"\bпгт\s+", "пгт. "),
                (r"\bрп\s+", "рп. "),
            ]

            for pattern_item, replacement in replacements:
                value = re.sub(pattern_item, replacement, value, flags=re.IGNORECASE)

            value = re.sub(r"\s+", " ", value)
            value = re.sub(r"\s+,", ",", value)

            return value.strip(" -_.,;")

        normalized = cleanup(rest)
        lowered = normalized.lower()

        owner = ""
        address = ""

        unknown_org_variants = [
            "не указана организация",
            "не указано организация",
            "организация не указана",
            "без организации",
        ]

        for variant in unknown_org_variants:
            if lowered.startswith(variant):
                owner = "Не указана организация"
                address = cleanup(normalized[len(variant):])
                break

        if not owner:
            address_pattern = re.compile(
                r"(?<![А-Яа-яA-Za-z0-9])"
                r"("
                r"г\.?|город|ул\.?|улица|д\.?|дом|"
                r"пр-кт|проспект|пр\.?|проезд|пер\.?|переулок|"
                r"ш\.?|шоссе|п\.?|пос\.?|поселок|посёлок|"
                r"с\.?|село|дер\.?|деревня|рп\.?|пгт\.?|мкр\.?"
                r")"
                r"\s+",
                re.IGNORECASE,
            )

            match = address_pattern.search(normalized)

            if match:
                if match.start() == 0:
                    owner = "Не указана организация"
                    address = normalized
                else:
                    owner = cleanup(normalized[:match.start()])
                    address = cleanup(normalized[match.start():])
            else:
                owner = cleanup(normalized)
                address = ""

        if owner.lower().startswith("не указана организация") and owner.lower() != "не указана организация":
            tail = owner[len("Не указана организация"):]
            owner = "Не указана организация"

            if not address:
                address = cleanup(tail)

        if address.lower().startswith("не указана организация"):
            address = cleanup(address[len("Не указана организация"):])

        owner = cleanup(owner) or "Не указана организация"
        address = cleanup(address) or normalized or filename

        return {
            "owner": owner,
            "address": address,
            "camera_status": camera_status,
            "camera_status_label": reason_label,
        }

    if not generated_items:
        generated_items = []

        for file_info in prescriptions:
            filename = file_info.get("filename", "")
            lower_name = filename.lower()

            if not lower_name.endswith(".docx"):
                continue

            if "combined" in lower_name or "общ" in lower_name or "all" in lower_name:
                continue

            parsed = parse_generated_prescription_filename(filename)

            generated_items.append(
                {
                    "owner": parsed.get("owner") or "Не указано",
                    "address": parsed.get("address") or filename,
                    "checked_at": "",
                    "camera_status": parsed.get("camera_status") or "",
                    "camera_status_label": parsed.get("camera_status_label") or "Статус не определён",
                    "filename": filename,
                    "url": file_info.get("url"),
                }
            )

    generated_at = (
        generation_result.get("generated_at")
        or generation_result.get("created_at")
        or state.get("generated_at")
        or state.get("updated_at")
        or state.get("checked_at")
        or state.get("created_at")
        or state.get("last_check")
        or ""
    )

    return templates.TemplateResponse(
        request,
        "camera_prescriptions.html",
        {
            "request": request,
            "mode": mode,
            "result": state,
            "state": state,
            "data": state,
            "rows": camera_rows,
            "cameras": camera_rows,
            "results": camera_rows,
            "items": problem_rows,
            "problem_items": problem_rows,
            "problem_rows": problem_rows,
            "problem_count": len(problem_rows),
            "generated_items": generated_items,
            "prescriptions": prescriptions,
            "generated_count": len(prescriptions),
            "generated_at": generated_at,
            "combined_filename": combined_filename,
            "combined_url": combined_url,
            "zip_filename": zip_filename,
            "zip_url": zip_url,
            "metrics": calculate_cameras_metrics({"rows": camera_rows}),
            "config": CAMERAS_UI_CONFIG,
            "check_state": make_check_state("camera_prescriptions"),
            "status": run_status["camera_prescriptions"],
            "run_status": run_status["camera_prescriptions"],
        },
    )


@app.get("/edo/run-status")
async def edo_run_status():
    return run_status["edo"]


@app.get("/overdue/run-status")
async def overdue_run_status():
    return run_status["overdue"]


@app.get("/watercontrol/run-status")
async def watercontrol_run_status():
    return run_status["watercontrol"]


@app.get("/utnkr/run-status")
async def utnkr_run_status():
    check_state = make_check_state("utnkr")
    current = run_status["utnkr"]
    return {
        **current,
        "started_at": check_state.get("started_at"),
        "finished_at": check_state.get("finished_at"),
        "check_state": check_state,
    }


@app.get("/cameras/run-status")
async def cameras_run_status():
    return run_status["cameras"]


@app.get("/cameras/address-table-info")
async def cameras_address_table_info():
    table_info = build_table_info(CAMERAS_ADDRESSES_FILE)

    return {
        "ok": True,
        "exists": table_info["exists"],
        "path": table_info["path"],
        "name": table_info["name"],
        "size": table_info["size"],
        "file": table_info,
    }


@app.get("/cameras/status")
async def cameras_status(
    status: str = "all",
    group_by: str = "none",
):
    state = load_json_file(CAMERAS_STATE_FILE, default={}) or {}
    rows = [normalize_camera_row(row) for row in as_list_from_result(state)]
    filtered_results = filter_camera_rows(rows, status)
    grouped_results = group_camera_rows(filtered_results, group_by)

    current = run_status["cameras"]
    check_state = make_check_state("cameras")

    if current.get("running"):
        check_state["status"] = "running"
    elif current.get("last_error"):
        check_state["status"] = "error"
    elif rows:
        check_state["status"] = "done"
    else:
        check_state["status"] = "idle"

    return {
        "ok": True,
        "status": check_state["status"],
        "check_state": check_state,
        "running": check_state["running"],
        "is_running": check_state["is_running"],
        "stage": check_state["stage"],
        "message": check_state["message"],
        "last_error": check_state["last_error"],
        "results": rows,
        "rows": rows,
        "items": rows,
        "filtered_results": filtered_results,
        "grouped_results": grouped_results,
        "status_filter": status,
        "group_by": group_by,
        "metrics": calculate_cameras_metrics({"rows": rows}),
        "table_info": build_table_info(CAMERAS_ADDRESSES_FILE),
    }


@app.post("/edo/run-check")
async def edo_run_check():
    command = [sys.executable, "-m", "services.edo.runner"]
    return start_background_service("edo", command)


@app.post("/overdue/run-check")
async def overdue_run_check():
    command = [sys.executable, "-m", "services.overdue.runner"]
    return start_background_service("overdue", command)


@app.post("/watercontrol/run-check")
async def watercontrol_run_check():
    command = [sys.executable, "-m", "services.watercontrol.runner"]
    return start_background_service("watercontrol", command)


@app.post("/utnkr/run-check")
async def utnkr_run_check():
    command = [sys.executable, "-m", "services.utnkr.scanner"]
    return start_background_service("utnkr", command)


@app.post("/cameras/run-check")
async def cameras_run_check():
    command = [sys.executable, "-m", "services.cameras.camera_checker"]
    return start_background_service("cameras", command)


@app.post("/camera-prescriptions/run-check")
async def camera_prescriptions_run_check():
    command = [sys.executable, "-m", "services.cameras.prescription_generator"]
    return start_background_service("camera_prescriptions", command)


@app.post("/cameras/generate-prescriptions")
async def cameras_generate_prescriptions():
    command = [sys.executable, "-m", "services.cameras.prescription_generator"]
    return start_background_service("camera_prescriptions", command)


@app.post("/edo/save-personal-message")
async def save_edo_personal_message(payload: dict):
    return save_personal_message_to_file(
        result_file=EDO_RESULT_FILE,
        payload=payload,
        not_found_message="Файл результата EDO не найден",
        allow_create=False,
    )


@app.post("/overdue/save-personal-message")
async def save_overdue_personal_message(payload: dict):
    return save_personal_message_to_file(
        result_file=OVERDUE_RESULT_FILE,
        payload=payload,
        not_found_message="Результаты проверки overdue ещё не сформированы",
        allow_create=True,
    )


@app.post("/watercontrol/save-personal-message")
async def save_watercontrol_personal_message(payload: dict):
    return save_personal_message_to_file(
        result_file=WATERCONTROL_RESULT_FILE,
        payload=payload,
        not_found_message="Файл результата WaterControl не найден",
        allow_create=True,
    )


@app.get("/api/summary")
async def api_summary():
    edo_result = load_json_file(EDO_RESULT_FILE)
    overdue_result = load_json_file(OVERDUE_RESULT_FILE)
    watercontrol_result = load_json_file(WATERCONTROL_RESULT_FILE)
    utnkr_result = load_json_file(UTNKR_RESULT_FILE)
    cameras_state = load_json_file(CAMERAS_STATE_FILE)

    return {
        "ok": True,
        "modules": {
            "edo": {
                "status": run_status["edo"],
                "metrics": calculate_edo_metrics(edo_result),
            },
            "overdue": {
                "status": run_status["overdue"],
                "metrics": calculate_overdue_metrics(overdue_result),
            },
            "watercontrol": {
                "status": run_status["watercontrol"],
                "metrics": calculate_watercontrol_metrics(watercontrol_result),
            },
            "utnkr": {
                "status": run_status["utnkr"],
                "metrics": calculate_utnkr_metrics(utnkr_result),
            },
            "cameras": {
                "status": run_status["cameras"],
                "metrics": calculate_cameras_metrics(cameras_state),
            },
            "camera_prescriptions": {
                "status": run_status["camera_prescriptions"],
            },
        },
    }


@app.get("/api/history/recent")
async def api_history_recent(request: Request, limit: int = 10):
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}
    role = (user.get("role") or "").strip().lower()
    if role not in {"admin", "администратор"} and role not in FULL_ACCESS_ROLES:
        raise HTTPException(status_code=403, detail="Доступ только для администратора")
    return run_history.get_recent(limit)

@app.get("/api/history/all")
async def api_history_all(request: Request, limit: int = 200):
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}
    role = (user.get("role") or "").strip().lower()
    if role not in {"admin", "администратор"} and role not in FULL_ACCESS_ROLES:
        raise HTTPException(status_code=403, detail="Доступ только для администратора")
    return run_history.get_all(limit)


@app.get("/system/git/check")
async def system_git_check(x_git_update_token: str | None = Header(default=None)):
    require_git_update_token(x_git_update_token)

    if not is_git_repository():
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "message": "Текущая папка проекта не является Git-репозиторием.",
                "project_root": str(BASE_DIR),
            },
        )

    branch = get_current_git_branch()
    fetch_result = run_git_command(["git", "fetch", GIT_REMOTE_NAME], timeout=120)
    status_result = run_git_command(["git", "status", "-sb"], timeout=30)
    incoming_result = run_git_command(
        [
            "git", "log", "--oneline", "--decorate", "--max-count=10",
            f"HEAD..{GIT_REMOTE_NAME}/{branch}",
        ],
        timeout=30,
    )

    local_changes = has_local_git_changes()
    has_updates = bool(incoming_result.get("stdout", "").strip())

    return {
        "ok": True,
        "message": "Проверка обновлений выполнена.",
        "project_root": str(BASE_DIR),
        "branch": branch,
        "remote": GIT_REMOTE_NAME,
        "has_updates": has_updates,
        "has_local_changes": local_changes.get("has_changes", True),
        "fetch": fetch_result,
        "status": status_result,
        "incoming_commits": incoming_result,
        "local_changes": local_changes,
    }


@app.post("/system/git/pull")
async def system_git_pull(x_git_update_token: str | None = Header(default=None)):
    require_git_update_token(x_git_update_token)

    if not is_git_repository():
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "message": "Текущая папка проекта не является Git-репозиторием.",
                "project_root": str(BASE_DIR),
            },
        )

    local_changes = has_local_git_changes()

    if not local_changes.get("ok"):
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "message": "Не удалось проверить локальные изменения.",
                "local_changes": local_changes,
            },
        )

    if local_changes.get("has_changes"):
        return JSONResponse(
            status_code=409,
            content={
                "ok": False,
                "message": "Есть локальные изменения. Автоматическое обновление остановлено, чтобы не потерять правки.",
                "hint": "Сначала выполните commit/stash или уберите локальные изменения.",
                "local_changes": local_changes,
            },
        )

    branch = get_current_git_branch()
    fetch_result = run_git_command(["git", "fetch", GIT_REMOTE_NAME], timeout=120)

    if not fetch_result.get("ok"):
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "message": "Не удалось выполнить git fetch.",
                "fetch": fetch_result,
            },
        )

    pull_result = run_git_command(["git", "pull", GIT_REMOTE_NAME, branch], timeout=180)
    response_status = 200 if pull_result.get("ok") else 400

    return JSONResponse(
        status_code=response_status,
        content={
            "ok": pull_result.get("ok"),
            "message": (
                "Обновления загружены. Если обновлялся Python-код, перезапустите приложение."
                if pull_result.get("ok")
                else "Не удалось загрузить обновления."
            ),
            "project_root": str(BASE_DIR),
            "branch": branch,
            "remote": GIT_REMOTE_NAME,
            "fetch": fetch_result,
            "pull": pull_result,
            "restart_required": True,
        },
    )


# =========================
# CDS / ДИСПЕТЧЕРСКАЯ ЖКХ
# =========================

CDS_DATA_DIR = DATA_DIR / "cds"
CDS_RESULT_FILE = CDS_DATA_DIR / "result.json"

run_status["cds"] = {
    "running": False,
    "stage": "Ожидание запуска",
    "message": "Система готова к выгрузке из CDS.",
    "last_error": "",
}


@app.get("/cds", response_class=HTMLResponse)
async def cds_page(request: Request):
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}

    return templates.TemplateResponse(
        request,
        "cds.html",
        {
            "request": request,
            "user_role": user.get("role", ""),
            "user_username": user.get("username", ""),
        },
    )


def build_cds_analytics(rows):
    from collections import Counter

    def g(r, *keys):
        for k in keys:
            v = r.get(k)
            if v:
                return str(v).strip()
        return ""

    by_status = Counter()
    by_type = Counter()
    by_municipality = Counter()
    by_day = Counter()

    for r in rows:
        if not isinstance(r, dict):
            continue

        by_status[g(r, "status", "Статус") or "Не указан"] += 1
        by_type[g(r, "type", "Тип заявки", "Тип") or "Не указан"] += 1

        address = g(r, "address", "Адрес")
        m = re.search(r"г\.?\s*о\.?\s*([А-ЯЁа-яё\- ]+?)(?:,|$)", address)
        if not m:
            m = re.search(r"\bг\.?\s*([А-ЯЁ][А-ЯЁа-яё\-]+)", address)
        by_municipality[m.group(1).strip() if m else "Не определён"] += 1

        d = g(r, "date", "Дата", "created_at", "Дата создания")
        dm = re.match(r"(\d{2})\.(\d{2})\.(\d{4})", d)
        if dm:
            by_day[dm.group(3) + "-" + dm.group(2) + "-" + dm.group(1)] += 1

    def top(counter, n=12):
        return [{"label": k, "value": v} for k, v in counter.most_common(n)]

    return {
        "by_status": top(by_status, 10),
        "by_type": top(by_type, 10),
        "by_municipality": top(by_municipality, 12),
        "by_day": [{"label": k, "value": v} for k, v in sorted(by_day.items())],
    }


@app.post("/api/cds/export")
async def api_cds_export(payload: dict):
    if run_status["cds"]["running"]:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": "Выгрузка уже выполняется."},
        )

    def to_ru_date(value: str) -> str:
        value = (value or "").strip()
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", value)
        if m:
            return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
        return value

    date_from = to_ru_date(payload.get("date_from", ""))
    date_to = to_ru_date(payload.get("date_to", ""))

    if not date_from or not date_to:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": "Не указан период (date_from, date_to)."},
        )

    run_status["cds"]["running"] = True
    run_status["cds"]["stage"] = "Выполняется"
    run_status["cds"]["message"] = f"Выгрузка за период {date_from} — {date_to}..."
    run_status["cds"]["last_error"] = ""

    try:
        from services.cds.scraper import scrape_cds_appeals

        result = await scrape_cds_appeals(date_from, date_to, headless=True)

        if not result.get("success"):
            error = result.get("error", "Неизвестная ошибка")
            run_status["cds"].update(
                running=False, stage="Ошибка", message=error, last_error=error
            )
            return JSONResponse(
                status_code=500,
                content={"success": False, "error": error},
            )

        rows = result.get("data", []) or []

        run_status["cds"].update(
            running=False,
            stage="Готово",
            message=f"Выгружено {len(rows)} обращений.",
        )

        norm_rows = []
        for r in rows:
            norm_rows.append({
                "date":         r.get("date") or r.get("Дата") or "",
                "department":   (r.get("department") or r.get("Подразделение") or "").strip() or "(не указано)",
                "number":       r.get("number") or r.get("Номер") or "",
                "status":       r.get("status") or r.get("Состояние обращения") or "",
                "address":      r.get("address") or r.get("Адрес обращения") or "",
                "reason":       r.get("reason") or r.get("Причина обращения") or r.get("type") or r.get("Тип обращения") or "",
                "deadline":     r.get("deadline") or r.get("Срок исполнения") or "",
                "type":         r.get("type") or r.get("Тип обращения") or "",
                "request_type": r.get("request_type") or r.get("Тип заявки") or "",
                "source":       r.get("source") or r.get("Источник поступления") or "",
                "applicant":    r.get("applicant") or r.get("Заявитель") or "",
                "executor":     r.get("executor") or r.get("Исполнитель") or "",
            })

        return {
            "success": True,
            "total": len(norm_rows),
            "rows": norm_rows,
            "download_url": "/data/cds/appeals.xlsx",
        }
    except Exception as e:
        import traceback
        run_status["cds"].update(
            running=False, stage="Ошибка",
            message="Ошибка: " + str(e), last_error=str(e),
        )
        print("[CDS] Export error:", traceback.format_exc())
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e)},
        )


@app.get("/api/cds/last-result")
async def api_cds_last_result():
    data = load_json_file(CDS_RESULT_FILE, default=None)

    if not data or not isinstance(data, dict) or not data.get("success"):
        return {"success": False, "rows": [], "timestamp": "", "date_from": "", "date_to": ""}

    rows = data.get("data", []) or []

    norm_rows = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        norm_rows.append({
            "date":         r.get("date") or r.get("Дата") or "",
            "department":   (r.get("department") or r.get("Подразделение") or "").strip() or "(не указано)",
            "number":       r.get("number") or r.get("Номер") or "",
            "status":       r.get("status") or r.get("Состояние обращения") or "",
            "address":      r.get("address") or r.get("Адрес обращения") or "",
            "reason":       r.get("reason") or r.get("Причина обращения") or r.get("type") or r.get("Тип обращения") or "",
            "deadline":     r.get("deadline") or r.get("Срок исполнения") or "",
            "type":         r.get("type") or r.get("Тип обращения") or "",
            "request_type": r.get("request_type") or r.get("Тип заявки") or "",
            "source":       r.get("source") or r.get("Источник поступления") or "",
            "applicant":    r.get("applicant") or r.get("Заявитель") or "",
            "executor":     r.get("executor") or r.get("Исполнитель") or "",
        })

    return {
        "success": True,
        "rows": norm_rows,
        "total": len(norm_rows),
        "timestamp": data.get("timestamp", ""),
        "date_from": data.get("date_from", ""),
        "date_to": data.get("date_to", ""),
    }


@app.get("/cds/run-status")
async def cds_run_status():
    return run_status["cds"]


# =========================
# APPEALS / ОБРАЩЕНИЯ
# =========================

@app.get("/appeals", response_class=HTMLResponse)
async def appeals_page(
    request: Request,
    status: str = "",
    message: str = "",
    error: str = "",
):
    appeals = list_appeals(status)
    stats = calculate_appeals_stats()

    return templates.TemplateResponse(
        request,
        "appeals.html",
        {
            "request": request,
            "appeals": appeals,
            "stats": stats,
            "current_status": status,
            "message": message,
            "error": error,
        },
    )


@app.post("/appeals/create")
async def appeals_create(request: Request):
    try:
        form = await request.form()

        subject = str(form.get("subject") or "").strip()
        sender_email = str(form.get("sender_email") or "manual@local").strip()
        text = str(form.get("text") or "").strip()

        uploaded_file = form.get("letter_file")
        uploaded_text = ""

        if uploaded_file is not None and getattr(uploaded_file, "filename", ""):
            file_bytes = await uploaded_file.read()
            uploaded_text = extract_text_from_file(uploaded_file.filename, file_bytes)

        original_text = "\n\n".join([part for part in [text, uploaded_text] if part.strip()]).strip()

        if not original_text:
            return RedirectResponse(
                url="/appeals?error=Нужно указать текст обращения или загрузить файл",
                status_code=303,
            )

        if not subject:
            subject = "Обращение без темы"

        item = create_appeal(
            subject=subject,
            original_text=original_text,
            sender_email=sender_email or "manual@local",
        )

        return RedirectResponse(
            url=f"/appeals/{item['request_id']}?message=Заявка создана",
            status_code=303,
        )

    except Exception as e:
        return RedirectResponse(
            url=f"/appeals?error=Ошибка создания обращения: {str(e)}",
            status_code=303,
        )


@app.get("/appeals/{request_id}", response_class=HTMLResponse)
async def appeal_detail_page(
    request: Request,
    request_id: str,
    message: str = "",
    error: str = "",
):
    item = get_appeal(request_id)

    if not item:
        return RedirectResponse(
            url="/appeals?error=Обращение не найдено",
            status_code=303,
        )

    return templates.TemplateResponse(
        request,
        "appeal_detail.html",
        {
            "request": request,
            "item": item,
            "default_template": DEFAULT_REPLY_TEMPLATE,
            "message": message,
            "error": error,
        },
    )


@app.post("/appeals/{request_id}/generate")
async def appeal_generate_draft(request: Request, request_id: str):
    item = get_appeal(request_id)

    if not item:
        return RedirectResponse(
            url="/appeals?error=Обращение не найдено",
            status_code=303,
        )

    form = await request.form()

    official_text = str(
        form.get("official_reply")
        or form.get("facts_for_reply")
        or ""
    ).strip()

    if not official_text:
        return RedirectResponse(
            url=f"/appeals/{request_id}?error=Нужно вставить официальный ответ",
            status_code=303,
        )

    item["facts_for_reply"] = official_text
    empathy_append_template = "{{ facts }}\n\n{{ empathy_block }}"

    draft = generate_reply_from_template(
        template_text=empathy_append_template,
        facts_text=official_text,
        item=item,
    )

    update_appeal(
        request_id,
        facts_for_reply=official_text,
        reply_template=empathy_append_template,
        draft=draft,
        status="awaiting_review",
        draft_created_at=datetime.now().isoformat(timespec="seconds"),
    )

    append_appeal_history(request_id, "draft_generated", {
        "official_reply": official_text,
        "draft": draft,
    })

    return RedirectResponse(
        url=f"/appeals/{request_id}?message=Эмпатичный блок добавлен к ответу",
        status_code=303,
    )


@app.post("/appeals/{request_id}/save-revision")
async def appeal_save_revision(request: Request, request_id: str):
    item = get_appeal(request_id)

    if not item:
        return RedirectResponse(
            url="/appeals?error=Обращение не найдено",
            status_code=303,
        )

    form = await request.form()
    manual_reply = str(form.get("manual_reply") or "").strip()

    if not manual_reply:
        return RedirectResponse(
            url=f"/appeals/{request_id}?error=Текст доработки не может быть пустым",
            status_code=303,
        )

    update_appeal(
        request_id,
        manual_reply=manual_reply,
        draft=manual_reply,
        status="awaiting_review",
    )

    append_appeal_history(request_id, "manual_revision_saved", {
        "manual_reply": manual_reply,
    })

    return RedirectResponse(
        url=f"/appeals/{request_id}?message=Доработка сохранена",
        status_code=303,
    )


@app.post("/appeals/{request_id}/approve")
async def appeal_approve(request_id: str):
    item = get_appeal(request_id)

    if not item:
        return RedirectResponse(
            url="/appeals?error=Обращение не найдено",
            status_code=303,
        )

    update_appeal(request_id, status="approved")
    append_appeal_history(request_id, "approved", {})

    return RedirectResponse(
        url=f"/appeals/{request_id}?message=Обращение утверждено",
        status_code=303,
    )


@app.post("/appeals/{request_id}/reject")
async def appeal_reject(request_id: str):
    item = get_appeal(request_id)

    if not item:
        return RedirectResponse(
            url="/appeals?error=Обращение не найдено",
            status_code=303,
        )

    update_appeal(request_id, status="rejected")
    append_appeal_history(request_id, "rejected", {})

    return RedirectResponse(
        url=f"/appeals/{request_id}?message=Обращение отклонено",
        status_code=303,
    )


@app.post("/appeals/{request_id}/mark-sent")
async def appeal_mark_sent(request_id: str):
    item = get_appeal(request_id)

    if not item:
        return RedirectResponse(
            url="/appeals?error=Обращение не найдено",
            status_code=303,
        )

    update_appeal(
        request_id,
        status="sent",
        sent_at=datetime.now().isoformat(timespec="seconds"),
    )
    append_appeal_history(request_id, "sent_marked", {})

    return RedirectResponse(
        url=f"/appeals/{request_id}?message=Обращение отмечено отправленным",
        status_code=303,
    )


@app.post("/appeals/{request_id}/delete")
async def delete_appeal(request_id: str):
    from services.appeals.storage import delete_appeal as delete_appeal_storage

    item = get_appeal(request_id)
    if not item:
        return RedirectResponse(
            url="/appeals?error=Обращение не найдено",
            status_code=303,
        )

    delete_appeal_storage(request_id)

    return RedirectResponse(
        url="/appeals?message=Обращение удалено",
        status_code=303,
    )


@app.get("/health")
async def health():
    return {
        "ok": True,
        "service": "Unified Dashboard",
    }




# =========================
# TOOLS / ПОЛЕЗНЫЕ ИНСТРУМЕНТЫ
# =========================
from services.tools import pptx_converter as pptx_conv


@app.get("/tools", response_class=HTMLResponse)
async def tools_page(request: Request):
    user = get_user_from_token(request.cookies.get("access_token")) or {}
    return templates.TemplateResponse(request, "tools.html", {
        "request": request,
        "templates": pptx_conv.list_templates(),
        "user_username": user.get("username", ""),
        "user_role": user.get("role", ""),
    })


@app.post("/tools/html2pptx/preview")
async def tools_preview(request: Request):
    form = await request.form()
    html = str(form.get("html") or "")
    file = form.get("html_file")
    if file is not None and getattr(file, "filename", ""):
        html = (await file.read()).decode("utf-8", errors="replace")
    if not html.strip():
        return JSONResponse(status_code=400, content={"ok": False, "message": "Нет HTML"})
    return {"ok": True, "slides": pptx_conv.parse_html_to_slides_pro(html)}


@app.post("/tools/html2pptx")
async def tools_convert(request: Request):
    form = await request.form()
    html = str(form.get("html") or "")
    file = form.get("html_file")
    if file is not None and getattr(file, "filename", ""):
        html = (await file.read()).decode("utf-8", errors="replace")
    template_name = str(form.get("template") or "") or None
    out_name = str(form.get("out_name") or "").strip() or "presentation.pptx"
    if not out_name.endswith(".pptx"):
        out_name += ".pptx"
    if not html.strip():
        return JSONResponse(status_code=400, content={"ok": False, "message": "Нет HTML"})
    mode = str(form.get("mode") or "smart")
    if mode == "shots":
        try:
            imgs = await asyncio.to_thread(pptx_conv.render_slides_images, html, Path(out_name).stem)
            if not imgs:
                return JSONResponse(status_code=500, content={"ok": False, "message": "Не удалось отрендерить слайды"})
            out = pptx_conv.build_pptx_from_images(imgs, out_name)
            rel = out.relative_to(BASE_DIR)
            return {"ok": True, "url": "/" + rel.as_posix(), "slides": len(imgs)}
        except Exception as e:
            import traceback
            traceback.print_exc()
            return JSONResponse(status_code=500, content={"ok": False, "message": str(e)})
    try:
        slides = pptx_conv.parse_html_to_slides_pro(html)
        out = pptx_conv.build_pptx(slides, template_name, out_name)
        rel = out.relative_to(BASE_DIR)
        return {"ok": True, "url": "/" + rel.as_posix(), "slides": len(slides)}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"ok": False, "message": str(e)})


@app.post("/tools/template")
async def tools_upload_template(request: Request):
    form = await request.form()
    file = form.get("template_file")
    if file is None or not getattr(file, "filename", ""):
        return JSONResponse(status_code=400, content={"ok": False, "message": "Нет файла"})
    if not file.filename.lower().endswith(".pptx"):
        return JSONResponse(status_code=400, content={"ok": False, "message": "Нужен .pptx"})
    pptx_conv.TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    (pptx_conv.TEMPLATES_DIR / file.filename).write_bytes(await file.read())
    return {"ok": True, "templates": pptx_conv.list_templates()}


@app.post("/tools/template/delete")
async def tools_delete_template(request: Request):
    form = await request.form()
    name = str(form.get("name") or "").strip()
    if not name or "/" in name or '\\' in name or ".." in name:
        return JSONResponse(status_code=400, content={"ok": False, "message": "Некорректное имя"})
    if not name.lower().endswith(".pptx"):
        return JSONResponse(status_code=400, content={"ok": False, "message": "Это не шаблон .pptx"})
    p = pptx_conv.TEMPLATES_DIR / name
    if not p.exists():
        return JSONResponse(status_code=404, content={"ok": False, "message": "Шаблон не найден"})
    p.unlink()
    print(f"[tools] шаблон удалён: {name}")
    return {"ok": True, "templates": pptx_conv.list_templates()}

@app.post("/tools/emblem")
async def tools_upload_emblem(request: Request):
    form = await request.form()
    file = form.get("emblem_file")
    if file is None or not getattr(file, "filename", ""):
        return JSONResponse(status_code=400, content={"ok": False, "message": "Нет файла"})
    low = file.filename.lower()
    if not (low.endswith(".png") or low.endswith(".jpg") or low.endswith(".jpeg")):
        return JSONResponse(status_code=400, content={"ok": False, "message": "Нужен .png или .jpg"})
    pptx_conv.TOOLS_DIR.mkdir(parents=True, exist_ok=True)
    dest = pptx_conv.TOOLS_DIR / ("emblem.png" if low.endswith(".png") else "emblem.jpg")
    dest.write_bytes(await file.read())
    print(f"[tools] герб загружен: {dest.name}")
    return {"ok": True}






# =========================
# SUMMARIZER / СУММАТОР
# =========================
from services.summarizer import engine as sum_engine
from services.summarizer import storage as sum_store


@app.get("/summarizer", response_class=HTMLResponse)
async def summarizer_page(request: Request):
    user = get_user_from_token(request.cookies.get("access_token")) or {}
    return templates.TemplateResponse(request, "summarizer.html", {
        "request": request,
        "user_username": user.get("username", ""),
        "reports": sum_store.list_reports(10),
    })


@app.post("/summarizer/api/summary")
async def summarizer_summary(request: Request):
    user = get_user_from_token(request.cookies.get("access_token")) or {}
    payload = await request.json()
    text = str(payload.get("text") or "").strip()
    if len(text) < 50:
        return JSONResponse(status_code=400,
                            content={"ok": False, "message": "Текст слишком короткий (мин. 50 символов)"})
    result = sum_engine.summarize(text)
    if not result.get("ok"):
        return JSONResponse(status_code=400, content=result)
    item = sum_store.create_report(text, result, user.get("username", "—"))
    return {"ok": True, "report": item}


@app.post("/summarizer/api/approve")
async def summarizer_approve(request: Request):
    user = get_user_from_token(request.cookies.get("access_token")) or {}
    payload = await request.json()
    it = sum_store.approve(str(payload.get("id") or ""), user.get("username", "—"))
    if not it:
        return JSONResponse(status_code=404, content={"ok": False, "message": "Отчёт не найден"})
    return {"ok": True, "report": it}


@app.post("/summarizer/api/reject")
async def summarizer_reject(request: Request):
    """BPMN: статус 'отклонить' не присваивается без комментария."""
    user = get_user_from_token(request.cookies.get("access_token")) or {}
    payload = await request.json()
    comment = str(payload.get("comment") or "").strip()
    if not comment:
        return JSONResponse(status_code=400,
                            content={"ok": False, "message": "Отклонение без комментария невозможно (правило BPMN)"})
    it = sum_store.reject(str(payload.get("id") or ""), user.get("username", "—"), comment)
    if not it:
        return JSONResponse(status_code=404, content={"ok": False, "message": "Отчёт не найден"})
    return {"ok": True, "report": it}


@app.post("/summarizer/api/regenerate")
async def summarizer_regenerate(request: Request):
    payload = await request.json()
    rid = str(payload.get("id") or "")
    it = sum_store.get_report(rid)
    if not it:
        return JSONResponse(status_code=404, content={"ok": False, "message": "Отчёт не найден"})
    src = it["source"]
    if it.get("revision_comment"):
        src += "\n\nКомментарий для корректировки: " + it["revision_comment"]
    result = sum_engine.summarize(src)
    if not result.get("ok"):
        return JSONResponse(status_code=400, content=result)
    it["result"] = result
    sum_store.to_pending(rid)
    return {"ok": True, "report": sum_store.get_report(rid)}


@app.get('/favicon.ico', include_in_schema=False)
async def favicon():
    return FileResponse('favicon.ico', media_type='image/png')


# ===============================
# MUNICIPALITY HTML/PDF REPORT
# ===============================
try:
    from services.municipality_report import router as municipality_report_router
    app.include_router(municipality_report_router)
except Exception as e:
    print(f"[municipality_report] router init error: {e}")


# ===============================
# SYSTEM CONTROL: RESTART & CACHE
# ===============================
try:
    from services.system_control import router as system_control_router
    app.include_router(system_control_router)
except Exception as e:
    print(f"[system_control] router init error: {e}")


# ===============================
# FEEDBACK / ПОМОЩЬ И ОБРАТНАЯ СВЯЗЬ
# ===============================
try:
    from services.feedback import router as feedback_router
    app.include_router(feedback_router)
except Exception as e:
    print(f"[feedback] router init error: {e}")


# ===============================
# SCHEDULER / АВТОЗАПУСК МОДУЛЕЙ
# ===============================
try:
    from services.scheduler import router as scheduler_router, start as start_scheduler
    from fastapi import Depends
    app.include_router(scheduler_router, dependencies=[Depends(require_admin_or_full)])
    start_scheduler()
    print("[scheduler] router connected (admin-only), loop started")
except Exception as e:
    print(f"[scheduler] init error: {e}")

@app.get("/scheduler", response_class=HTMLResponse)
async def scheduler_page(request: Request):
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}

    role = (user.get("role") or "").strip().lower()
    if role not in ADMIN_ROLES and role not in FULL_ACCESS_ROLES:
        return RedirectResponse(url="/", status_code=303)

    return templates.TemplateResponse(
        request,
        "scheduler.html",
        {
            "request": request,
            "user": user,
            "user_role": user.get("role", ""),
            "user_username": user.get("username", ""),
        },
    )


# ===== Модуль «Проверка задач по качеству воды» =====
try:
    from services.water_rm.proxy import router as water_rm_router
    app.include_router(water_rm_router)
except Exception as e:
    print(f"[water_rm] router init error: {e}")


@app.get("/water-rm", response_class=HTMLResponse)
async def water_rm_page(request: Request):
    return templates.TemplateResponse(request, "water_rm.html", {"request": request})


# ===============================
# ECUR / КОНТРОЛЬ ЖАЛОБ ДОБРОДЕЛА
# ===============================

ECUR_UI_CONFIG = {
    "source_name": "ДоброДел",
    "source_url": "https://admin.vmeste.mosreg.ru/",
}

run_status["ecur"] = {
    "running": False,
    "stage": "Ожидание запуска",
    "message": "Требуется авторизация в системе.",
    "last_error": "",
}


# ── СХЕМА ДЛЯ ВХОДА ──
class ECURLoginRequest(BaseModel):
    email: str
    password: str


def calculate_ecur_metrics(rows):
    """KPI по срокам: просрочено / сегодня / неделя / месяц / всего."""
    from datetime import datetime as _dt, timedelta

    if not rows or len(rows) < 2:
        return {"total": 0, "overdue": 0, "today": 0, "week": 0, "month": 0}

    header = rows[0]
    try:
        idx_deadline = header.index("Срок")
    except ValueError:
        idx_deadline = 11

    def parse_ru_date(s):
        if not s:
            return None
        s = str(s).strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
            try:
                return _dt.strptime(s.split()[0], fmt).date()
            except Exception:
                continue
        return None

    today_date = _dt.now().date()
    end_of_week = today_date + timedelta(days=(6 - today_date.weekday()))
    end_of_month = (today_date.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)

    overdue = today = week = month = 0
    total = len(rows) - 1

    for r in rows[1:]:
        if idx_deadline >= len(r):
            continue
        d = parse_ru_date(r[idx_deadline])
        if not d:
            continue
        if d < today_date:
            overdue += 1
        elif d == today_date:
            today += 1
        elif d <= end_of_week:
            week += 1
        elif d <= end_of_month:
            month += 1

    return {
        "total": total,
        "overdue": overdue,
        "today": today,
        "week": week,
        "month": month,
    }


# ── СТРАНИЦА ЕЦУР ──
@app.get("/ecur", response_class=HTMLResponse)
async def ecur_page(request: Request):
    from services.ecur.client import get_current_data

    data = get_current_data()
    rows = data.get("rows")
    meta = data.get("meta") or {}
    is_authed = data.get("is_authed", False)
    user_email = data.get("email", "")

    metrics = calculate_ecur_metrics(rows) if rows else {
        "total": 0, "overdue": 0, "today": 0, "week": 0, "month": 0
    }

    return templates.TemplateResponse(
        request,
        "ecur.html",
        {
            "request": request,
            "metrics": metrics,
            "has_data": bool(rows),
            "is_authed": is_authed,
            "user_email": user_email,
            "meta": meta,
            "config": ECUR_UI_CONFIG,
            "check_state": make_check_state("ecur"),
            "status": run_status["ecur"],
            "run_status": run_status["ecur"],
        },
    )


# ── СТАТУС ЗАПУСКА ──
@app.get("/ecur/run-status")
async def ecur_run_status():
    return run_status["ecur"]


# ── ВХОД ЧЕРЕЗ ИНТЕРФЕЙС (API) ──
@app.post("/ecur/api/login")
async def ecur_api_login(data: ECURLoginRequest):
    """Выполняет вход на портал под переданным email и паролем."""
    from services.ecur.client import authenticate_user

    if run_status["ecur"]["running"]:
        return JSONResponse(
            status_code=409,
            content={"ok": False, "error": "Обновление уже выполняется."},
        )

    run_status["ecur"].update(
        running=True,
        stage="Вход в ДоброДел",
        message="Проверка учётных данных…",
        last_error="",
        started_at=datetime.now().isoformat(timespec="seconds"),
        finished_at="",
    )

    ok, result = authenticate_user(data.email, data.password)

    if ok:
        run_status["ecur"].update(
            running=False,
            stage="Готово",
            message=f"Успешный вход. Загружено жалоб: {result}",
        )
        return {"ok": True, "count": result, "message": "Успешный вход."}
    else:
        run_status["ecur"].update(
            running=False,
            stage="Ошибка входа",
            message="Не удалось войти.",
            last_error=str(result),
        )
        return JSONResponse(
            status_code=400,
            content={"ok": False, "error": str(result)},
        )


# ── ВЫХОД / СБРОС СЕССИИ (API) ──
@app.post("/ecur/api/logout")
async def ecur_api_logout():
    """Сбрасывает текущую сессию и удаляет данные из памяти."""
    from services.ecur.client import clear_session

    clear_session()
    run_status["ecur"].update(
        running=False,
        stage="Ожидание запуска",
        message="Сессия завершена.",
        last_error="",
    )
    return {"ok": True}


# ── ОБНОВЛЕНИЕ СВОДА ──
@app.post("/ecur/api/refresh")
async def ecur_api_refresh():
    """Обновляет свод жалоб из ДоброДела под текущей сессией."""
    from services.ecur.client import refresh_data, get_current_data

    current_data = get_current_data()
    if not current_data.get("is_authed"):
        return JSONResponse(
            status_code=401,
            content={"ok": False, "error": "Сессия не активна. Авторизуйтесь заново."},
        )

    if run_status["ecur"]["running"]:
        return JSONResponse(
            status_code=409,
            content={"ok": False, "error": "Обновление уже выполняется."},
        )

    def worker():
        run_id = None
        try:
            run_id = run_history.record_start("ecur", user="—")
            run_status["ecur"].update(
                running=True,
                stage="Выгрузка свода",
                message="Загрузка свежих данных с портала…",
                last_error="",
                started_at=datetime.now().isoformat(timespec="seconds"),
                finished_at="",
            )
            save_run_time("ecur")

            ok, result = refresh_data()

            if ok:
                run_status["ecur"].update(
                    running=False,
                    stage="Готово",
                    message=f"Загружено жалоб: {result}",
                )
                if run_id:
                    run_history.record_finish(run_id, "success")
            else:
                run_status["ecur"].update(
                    running=False,
                    stage="Ошибка",
                    message="Не удалось получить свод.",
                    last_error=str(result),
                )
                if run_id:
                    run_history.record_finish(run_id, "error", str(result))
        except Exception as e:
            run_status["ecur"].update(
                running=False,
                stage="Ошибка",
                message="Внутренняя ошибка.",
                last_error=str(e),
            )
            if run_id:
                run_history.record_finish(run_id, "error", str(e))
        finally:
            run_status["ecur"]["finished_at"] = datetime.now().isoformat(timespec="seconds")
            save_run_time("ecur")

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    time.sleep(0.3)

    st = run_status["ecur"]
    if st["running"]:
        return {"ok": True, "message": "Обновление запущено.", "running": True}

    if st["last_error"]:
        return JSONResponse(
            status_code=500,
            content={"ok": False, "error": st["last_error"]},
        )
    return {"ok": True, "count": None, "message": st.get("message", "Готово")}


# ── ОТДАЧА data.js ──
@app.get("/ecur/data.js")
async def ecur_data_js():
    """Отдаёт data.js с данными для дашборда."""
    from services.ecur.client import get_current_data

    data = get_current_data()
    rows = data.get("rows")
    meta = data.get("meta")

    if rows:
        payload = (
            "/* сгенерировано из памяти */\n"
            "window.PRELOADED_META = " + json.dumps(meta, ensure_ascii=False) + ";\n"
            "window.PRELOADED_ROWS = " + json.dumps(rows, ensure_ascii=False) + ";\n"
        )
    else:
        payload = "window.PRELOADED_META=null;window.PRELOADED_ROWS=[];"

    return Response(
        content=payload,
        media_type="application/javascript; charset=utf-8",
    )



@app.post("/water-dashboard/run-check")
async def water_dashboard_run_check():
    command = [sys.executable, "-m", "services.water_dashboard.runner"]
    return start_background_service("water_dashboard", command)

@app.get("/water-dashboard/run-status")
async def water_dashboard_run_status():
    return run_status["water_dashboard"]

@app.get("/api/notifications")
async def list_notifications(request: Request):
    user = get_user_from_token(request.cookies.get("access_token"))
    if not user:
        return JSONResponse(status_code=401, content={"detail": "Требуется авторизация"})
    items = notif_store.list_all()
    return JSONResponse(content=items)
