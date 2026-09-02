# -*- coding: utf-8 -*-
"""
routers/system.py

Системные роуты:
- health
- notifications
- assistant ask
- сводка по всем модулям
"""

import json
import re
from pathlib import Path
from datetime import datetime, timedelta
from typing import Any

from fastapi import APIRouter, Request
from fastapi.responses import JSONResponse

from services.auth.security import get_user_from_token
from services.notifications import store as notif_store
from services.appeals.storage import calculate_stats as calculate_appeals_stats


router = APIRouter()

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = BASE_DIR / "data"
GENERATED_DIR = BASE_DIR / "generated"

EDO_RESULT_FILE = DATA_DIR / "edo" / "result.json"
OVERDUE_RESULT_FILE = DATA_DIR / "overdue" / "final_result.json"
WATERCONTROL_RESULT_FILE = DATA_DIR / "watercontrol" / "result.json"
UTNKR_RESULT_FILE = DATA_DIR / "utnkr" / "violators.json"
MGKH_RM_RESULT_FILE = DATA_DIR / "mgkh_rm" / "result.json"
CAMERAS_STATE_FILE = DATA_DIR / "cameras" / "state" / "dashboard_state.json"
CDS_RESULT_FILE = DATA_DIR / "cds" / "result.json"
WATER_DASHBOARD_SNAPSHOT = DATA_DIR / "water_dashboard" / "snapshot.json"
GENERATED_PRESCRIPTIONS_DIR = GENERATED_DIR / "prescriptions"
ZIP_PUB_FILE = DATA_DIR / "zip_curator" / "published.json"


def load_json_file(path: Path, default=None):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def normalize_text(value):
    return (value or "").strip()


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


def calculate_ecur_metrics(rows):
    """KPI по срокам: просрочено / сегодня / неделя / месяц / всего."""
    if not rows or len(rows) < 2:
        return {"total": 0, "overdue": 0, "today": 0, "week": 0, "month": 0}

    header = rows[0]
    try:
        idx_deadline = header.index("Срок")
    except Exception:
        idx_deadline = 11

    def parse_ru_date(s):
        if not s:
            return None
        s = str(s).strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(s.split()[0], fmt).date()
            except Exception:
                continue
        return None

    today_date = datetime.now().date()
    end_of_week = today_date + timedelta(days=(6 - today_date.weekday()))
    end_of_month = (
        today_date.replace(day=1) + timedelta(days=32)
    ).replace(day=1) - timedelta(days=1)

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


def build_system_status_context() -> str:
    """Собирает актуальные метрики по всем блокам системы."""
    blocks = []

    # Камеры
    try:
        cameras_state = load_json_file(CAMERAS_STATE_FILE)
        if cameras_state:
            m = calculate_cameras_metrics(cameras_state)
            blocks.append(
                f"📹 Камеры: всего {m['total']}, работает {m['working']}, "
                f"не работает {m['not_working']}, не подключено {m['not_connected']}, "
                f"статус неизвестен {m['unknown']}."
            )
    except Exception:
        pass

    # ЭДО
    try:
        edo_result = load_json_file(EDO_RESULT_FILE)
        if edo_result:
            m = calculate_edo_metrics(edo_result)
            blocks.append(
                f"📄 ЭДО: всего {m['total']}, критичных {m['critical']}, "
                f"риск {m['risk']}, в норме {m['ok']}."
            )
    except Exception:
        pass

    # Просрочка
    try:
        overdue_result = load_json_file(OVERDUE_RESULT_FILE)
        if overdue_result:
            m = calculate_overdue_metrics(overdue_result)
            blocks.append(
                f"⏳ Просроченные задачи: всего {m['total']}, "
                f"критичных {m['critical']}, риск {m['risk']}, в норме {m['ok']}."
            )
    except Exception:
        pass

    # УТНКР
    try:
        utnkr_result = load_json_file(UTNKR_RESULT_FILE)
        if utnkr_result:
            m = calculate_utnkr_metrics(utnkr_result)
            blocks.append(
                f"🏗 Технадзор УТНКР: всего {m['total']}, "
                f"критичных {m['critical']}, риск {m['risk']}, в норме {m['ok']}."
            )
    except Exception:
        pass

    # WaterControl
    try:
        watercontrol_result = load_json_file(WATERCONTROL_RESULT_FILE)
        if watercontrol_result:
            m = calculate_watercontrol_metrics(watercontrol_result)
            blocks.append(
                f"💧 WaterControl: всего {m['total']}, критичных {m['critical']}, "
                f"риск {m['risk']}, в норме {m['ok']}."
            )
    except Exception:
        pass

    # МКХ Redmine
    try:
        mgkh_result = load_json_file(MGKH_RM_RESULT_FILE)
        if mgkh_result:
            m = calculate_mgkh_rm_metrics(mgkh_result)
            blocks.append(
                f"🛠 МКХ Redmine: всего {m['total']}, закрыть {m['close']}, "
                f"продлить {m['extend']}, переделать {m['rework']}."
            )
    except Exception:
        pass

    # ЦДС
    try:
        cds_result = load_json_file(CDS_RESULT_FILE)
        if cds_result and isinstance(cds_result, dict):
            cds_rows = cds_result.get("data", []) or cds_result.get("rows", []) or []
            blocks.append(
                f"📞 ЦДС: всего выгружено обращений {len(cds_rows)}."
            )
    except Exception:
        pass

    # Appeals
    try:
        stats = calculate_appeals_stats()
        if stats and isinstance(stats, dict):
            blocks.append(
                f"📬 Обращения граждан: всего {stats.get('total', 0)}, "
                f"на рассмотрении {stats.get('awaiting_review', 0)}, "
                f"согласовано {stats.get('approved', 0)}, "
                f"отправлено {stats.get('sent', 0)}."
            )
    except Exception:
        pass

    # ЕЦУР / ДоброДел
    try:
        from services.ecur.client import get_current_data

        ecur_data = get_current_data()
        ecur_rows = ecur_data.get("rows", [])
        if ecur_rows:
            m = calculate_ecur_metrics(ecur_rows)
            blocks.append(
                f"🏛 ДоброДел / ЕЦУР: всего {m['total']}, "
                f"просрочено {m['overdue']}, сегодня {m['today']}, "
                f"на неделе {m['week']}, в месяце {m['month']}."
            )
    except Exception:
        pass

    # Водный дашборд
    try:
        wd_snap = load_json_file(WATER_DASHBOARD_SNAPSHOT)
        if wd_snap and isinstance(wd_snap, dict):
            snap_date = wd_snap.get("snapshot_date", "неизвестно")
            updated_at = wd_snap.get("updated_at", "")
            table = wd_snap.get("table", []) or []
            blocks.append(
                f"🌊 Водный дашборд: строк в таблице {len(table)}, "
                f"снимок от {snap_date}, обновлено {updated_at or 'не указано'}."
            )
    except Exception:
        pass

    # Предписания по камерам
    try:
        gen_res = load_json_file(GENERATED_PRESCRIPTIONS_DIR / "generation_result.json")
        if gen_res and isinstance(gen_res, dict):
            items = gen_res.get("items", []) or []
            generated_at = gen_res.get("generated_at") or gen_res.get("created_at") or ""
            blocks.append(
                f"📑 Предписания по камерам: сформировано {len(items)}, "
                f"дата формирования {generated_at or 'не указана'}."
            )
    except Exception:
        pass

    # Куратор ЗиП
    try:
        from services.zip_curator import core as zc

        zc_state = zc.load_state()
        if zc_state and isinstance(zc_state, dict):
            pending = len(zc_state.get("pending", []))
            clean = len(zc_state.get("clean", []))
            blocks.append(
                f"📦 Куратор ЗиП РСО: на согласовании {pending}, согласовано {clean}."
            )
    except Exception:
        pass

    # Публикация ЗиП
    try:
        zip_pub = load_json_file(ZIP_PUB_FILE)
        if zip_pub and isinstance(zip_pub, dict):
            rows = zip_pub.get("rows", []) or []
            published_at = zip_pub.get("published_at")
            blocks.append(
                f"📦 Опубликованные остатки ЗиП: строк {max(len(rows) - 1, 0)}, "
                f"опубликовано {published_at or 'не указано'}."
            )
    except Exception:
        pass

    if not blocks:
        return "Данные по проверкам пока отсутствуют или файлы результатов пусты."

    return "Актуальная сводка по всем модулям системы:\n" + "\n".join(blocks)


@router.get("/health")
async def health():
    return {
        "ok": True,
        "service": "Unified Dashboard",
    }


@router.get("/api/notifications")
async def list_notifications(request: Request):
    user = get_user_from_token(request.cookies.get("access_token"))
    if not user:
        return JSONResponse(
            status_code=401,
            content={"detail": "Требуется авторизация"},
        )

    items = notif_store.list_all()
    return JSONResponse(content=items)


@router.post("/api/assistant/ask")
async def assistant_ask(payload: dict):
    """
    Принимает вопрос пользователя, подставляет актуальную сводку по всем блокам
    и просит ИИ ответить на вопрос на основе этих данных.
    """
    question = (payload.get("question") or "").strip()
    if not question:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "message": "Пустой вопрос"},
        )

    context = build_system_status_context()

    system_prompt = (
        "Ты — аналитик и голосовой помощник системы мониторинга ЖКХ. "
        "Отвечай на вопросы пользователя кратко, по существу и используя только предоставленные цифры. "
        "Если в данных нет ответа на вопрос, так и скажи. Не выдумывай факты.\n\n"
        f"Данные системы:\n{context}"
    )

    try:
        from services.summarizer.engine import _qwen_chat

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": question},
        ]

        answer = _qwen_chat(messages, max_tokens=300)

        return {
            "ok": True,
            "question": question,
            "answer": str(answer).strip(),
            "raw_context": context,
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "ok": False,
                "message": f"Ошибка ИИ: {str(e)}",
                "raw_context": context,
            },
        )
