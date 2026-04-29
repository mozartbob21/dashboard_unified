from pathlib import Path
import json
import subprocess
import sys
import threading

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"

EDO_DATA_DIR = DATA_DIR / "edo"
OVERDUE_DATA_DIR = DATA_DIR / "overdue"
WATERCONTROL_DATA_DIR = DATA_DIR / "watercontrol"

EDO_RESULT_FILE = EDO_DATA_DIR / "result.json"
OVERDUE_RESULT_FILE = OVERDUE_DATA_DIR / "final_result.json"
WATERCONTROL_RESULT_FILE = WATERCONTROL_DATA_DIR / "result.json"

app = FastAPI(title="Unified Dashboard")

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
app.mount("/data", StaticFiles(directory=str(DATA_DIR)), name="data")

templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


run_status = {
    "edo": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выполнению проверки EDO.",
        "last_error": ""
    },
    "overdue": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выполнению проверки просроченных задач.",
        "last_error": ""
    },
    "watercontrol": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к выполнению проверки WaterControl.",
        "last_error": ""
    }
}


def load_json_file(path: Path, default=None):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def save_json_file(path: Path, data: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8"
    )


def normalize_text(value):
    return (value or "").strip()


def normalize_key(municipality, organization):
    return (
        normalize_text(municipality).upper(),
        normalize_text(organization).lower()
    )


def to_int(value, default=0):
    try:
        if value is None:
            return default
        if isinstance(value, str):
            value = value.strip().replace(" ", "").replace("\u00A0", "").replace(",", ".")
            if not value:
                return default
        return int(float(value))
    except Exception:
        return default


def ensure_personal_message_flags(result, result_file: Path):
    if not result:
        return result

    personal_messages = result.get("personal_messages", []) or []
    changed = False

    for item in personal_messages:
        if "is_edited" not in item:
            item["is_edited"] = False
            changed = True

    if changed:
        try:
            save_json_file(result_file, result)
        except Exception as e:
            print("Ошибка при обновлении is_edited:", e)

    return result


def calculate_edo_metrics(result):
    rows = (result or {}).get("rows", []) or []

    total = len(rows)
    critical = sum(1 for row in rows if row.get("status") == "critical")
    risk = sum(1 for row in rows if row.get("status") == "risk")
    ok = sum(1 for row in rows if row.get("status") == "ok")

    return {
        "total": total,
        "critical": critical,
        "risk": risk,
        "ok": ok
    }


def calculate_overdue_metrics(raw_result):
    if not raw_result:
        return {
            "total": 0,
            "critical": 0,
            "risk": 0,
            "ok": 0,
        }

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


def calculate_watercontrol_metrics(result):
    rows = (result or {}).get("rows", []) or []

    total = len(rows)
    critical = sum(1 for row in rows if row.get("status") == "critical")
    risk = sum(1 for row in rows if row.get("status") == "risk")
    ok = sum(1 for row in rows if row.get("status") == "ok")

    return {
        "total": total,
        "critical": critical,
        "risk": risk,
        "ok": ok
    }


def transform_overdue_result_for_ui(raw):
    if not raw:
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

        rows.append({
            "municipality": item.get("municipality", ""),
            "organization": item.get("organization", item.get("municipality", "")),
            "responsible_name": item.get("responsible_name", ""),
            "responsible_phone": item.get("responsible_phone", ""),
            "status": status,
            "reason": reason,
            "overdue_count": overdue_count,
        })

    rows.sort(key=lambda x: (-to_int(x.get("overdue_count", 0)), x.get("municipality", "")))

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


def run_subprocess_worker(service_name: str, command: list[str], cwd: Path):
    status = run_status[service_name]

    try:
        status["running"] = True
        status["last_error"] = ""
        status["stage"] = "Запуск"
        status["message"] = f"Запущена проверка {service_name}."

        process = subprocess.Popen(
            command,
            cwd=str(cwd),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace"
        )

        if process.stdout:
            for line in process.stdout:
                text = line.strip()
                if not text:
                    continue

                print(f"[{service_name.upper()}]", text)

                if text.startswith("STAGE:"):
                    stage_name = text.replace("STAGE:", "").strip()
                    status["stage"] = stage_name
                    status["message"] = stage_name
                    continue

                lowered = text.lower()

                if "captcha" in lowered or "капча" in lowered:
                    status["stage"] = "Ожидание подтверждения"
                    status["message"] = text
                elif "screenshot" in lowered or "скриншот" in lowered:
                    status["stage"] = "Снятие скриншотов"
                    status["message"] = text
                elif "анализ" in lowered or "извлеч" in lowered:
                    status["stage"] = "Анализ данных"
                    status["message"] = text
                elif "сохранение" in lowered or "result saved" in lowered or "готово:" in lowered:
                    status["stage"] = "Сохранение результата"
                    status["message"] = text

        return_code = process.wait()

        if return_code != 0:
            status["running"] = False
            status["stage"] = "Ошибка"
            status["message"] = f"Проверка {service_name} завершилась с ошибкой."
            status["last_error"] = f"Процесс завершился с кодом {return_code}"
            return

        status["running"] = False
        status["stage"] = "Готово"
        status["message"] = f"Проверка {service_name} завершена успешно."

    except Exception as e:
        status["running"] = False
        status["stage"] = "Ошибка"
        status["message"] = "Во время выполнения произошла ошибка."
        status["last_error"] = str(e)


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        request,
        "home.html",
        {
            "request": request
        }
    )


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
            "metrics": metrics
        }
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
            "metrics": metrics
        }
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
            "metrics": metrics
        }
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


@app.post("/edo/run-check")
async def edo_run_check():
    if run_status["edo"]["running"]:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "message": "Проверка EDO уже выполняется."}
        )

    command = [sys.executable, "-m", "services.edo.runner"]
    thread = threading.Thread(
        target=run_subprocess_worker,
        args=("edo", command, BASE_DIR),
        daemon=True
    )
    thread.start()

    return {"ok": True}


@app.post("/overdue/run-check")
async def overdue_run_check():
    if run_status["overdue"]["running"]:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "message": "Проверка overdue уже выполняется."}
        )

    command = [sys.executable, "-m", "services.overdue.runner"]
    thread = threading.Thread(
        target=run_subprocess_worker,
        args=("overdue", command, BASE_DIR),
        daemon=True
    )
    thread.start()

    return {"ok": True}


@app.post("/watercontrol/run-check")
async def watercontrol_run_check():
    if run_status["watercontrol"]["running"]:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "message": "Проверка WaterControl уже выполняется."}
        )

    command = [sys.executable, "-m", "services.watercontrol.runner"]
    thread = threading.Thread(
        target=run_subprocess_worker,
        args=("watercontrol", command, BASE_DIR),
        daemon=True
    )
    thread.start()

    return {"ok": True}


@app.post("/edo/save-personal-message")
async def save_edo_personal_message(payload: dict):
    municipality = normalize_text(payload.get("municipality"))
    organization = normalize_text(payload.get("organization"))
    message = normalize_text(payload.get("message"))

    if not municipality or not organization or not message:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "message": "Не хватает данных для сохранения"}
        )

    data = load_json_file(EDO_RESULT_FILE)
    if not data:
        return JSONResponse(
            status_code=404,
            content={"ok": False, "message": "Файл результата EDO не найден"}
        )

    target_key = normalize_key(municipality, organization)
    updated = False

    for item in data.get("personal_messages", []):
        item_key = normalize_key(item.get("municipality", ""), item.get("organization", ""))
        if item_key == target_key:
            item["message"] = message
            item["is_edited"] = True
            updated = True
            break

    if not updated:
        return JSONResponse(
            status_code=404,
            content={"ok": False, "message": "Уведомление EDO не найдено"}
        )

    save_json_file(EDO_RESULT_FILE, data)
    return {"ok": True}


@app.post("/overdue/save-personal-message")
async def save_overdue_personal_message(payload: dict):
    municipality = normalize_text(payload.get("municipality"))
    organization = normalize_text(payload.get("organization"))
    message = normalize_text(payload.get("message"))

    if not municipality or not organization or not message:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "message": "Недостаточно данных для сохранения"}
        )

    data = load_json_file(OVERDUE_RESULT_FILE)
    if not data:
        return JSONResponse(
            status_code=404,
            content={"ok": False, "message": "Результаты проверки overdue ещё не сформированы"}
        )

    personal_messages = data.get("personal_messages", []) or []
    target_key = normalize_key(municipality, organization)
    updated = False

    for item in personal_messages:
        item_key = normalize_key(item.get("municipality", ""), item.get("organization", ""))
        if item_key == target_key:
            item["message"] = message
            item["is_edited"] = True
            updated = True
            break

    if not updated:
        personal_messages.append({
            "municipality": municipality,
            "organization": organization,
            "message": message,
            "is_edited": True,
            "status": "risk",
            "responsible_name": "",
            "responsible_phone": "",
        })

    data["personal_messages"] = personal_messages
    save_json_file(OVERDUE_RESULT_FILE, data)

    return {"ok": True, "message": "Сообщение сохранено."}


@app.post("/watercontrol/save-personal-message")
async def save_watercontrol_personal_message(payload: dict):
    municipality = normalize_text(payload.get("municipality"))
    organization = normalize_text(payload.get("organization"))
    message = normalize_text(payload.get("message"))

    if not municipality or not organization or not message:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "message": "Не хватает данных для сохранения"}
        )

    data = load_json_file(WATERCONTROL_RESULT_FILE)
    if not data:
        return JSONResponse(
            status_code=404,
            content={"ok": False, "message": "Файл результата WaterControl не найден"}
        )

    target_key = normalize_key(municipality, organization)
    updated = False

    for item in data.get("personal_messages", []):
        item_key = normalize_key(item.get("municipality", ""), item.get("organization", ""))
        if item_key == target_key:
            item["message"] = message
            item["is_edited"] = True
            updated = True
            break

    if not updated:
        data.setdefault("personal_messages", []).append({
            "municipality": municipality,
            "organization": organization,
            "message": message,
            "is_edited": True
        })

    save_json_file(WATERCONTROL_RESULT_FILE, data)
    return {"ok": True}