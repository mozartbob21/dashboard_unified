"""
Модуль истории запусков проверок.

Хранит записи о всех запусках модулей (ЭДО, просрочки, камеры и т.д.)
в JSON-файле data/run_history.json.

Используется в app.py: запись стартует в run_subprocess_worker()
при старте процесса и обновляется по завершению.
"""

from __future__ import annotations

import json
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

BASE_DIR = Path(__file__).resolve().parent.parent
HISTORY_FILE = BASE_DIR / "data" / "run_history.json"
MAX_RECORDS = 200  # храним последние 200 запусков

_lock = threading.Lock()

# Человекочитаемые названия модулей для UI
MODULE_LABELS = {
    "edo": "ЭДО",
    "overdue": "Просроченные задачи",
    "watercontrol": "WaterControl",
    "utnkr": "УТНКР",
    "cameras": "Проверка камер",
    "camera_prescriptions": "Предписания по камерам",
}

MODULE_ICONS = {
    "edo": "📋",
    "overdue": "⏱️",
    "watercontrol": "💧",
    "utnkr": "🔍",
    "cameras": "📹",
    "camera_prescriptions": "📄",
}


def _load() -> list[dict]:
    if not HISTORY_FILE.exists():
        return []
    try:
        data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception as e:
        print(f"[run_history] Ошибка чтения {HISTORY_FILE}: {e}")
        return []


def _save(records: list[dict]) -> None:
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    # Обрезаем до MAX_RECORDS, оставляя самые свежие (в начале списка)
    if len(records) > MAX_RECORDS:
        records = records[:MAX_RECORDS]
    HISTORY_FILE.write_text(
        json.dumps(records, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def record_start(module_id: str, user: str = "—") -> str:
    """Записывает старт запуска. Возвращает run_id."""
    run_id = str(uuid.uuid4())
    record = {
        "run_id": run_id,
        "module_id": module_id,
        "module_label": MODULE_LABELS.get(module_id, module_id),
        "module_icon": MODULE_ICONS.get(module_id, "⚙️"),
        "user": user or "—",
        "started_at": _now_iso(),
        "finished_at": None,
        "duration_seconds": None,
        "status": "running",  # running | success | error
        "error_message": "",
    }
    with _lock:
        records = _load()
        records.insert(0, record)  # самые свежие — в начало
        _save(records)
    return run_id


def record_finish(
    run_id: str,
    status: str = "success",
    error_message: str = "",
) -> None:
    """Обновляет запись о завершении запуска."""
    if status not in ("success", "error"):
        status = "success"

    with _lock:
        records = _load()
        for rec in records:
            if rec.get("run_id") == run_id:
                finished_at = datetime.now(timezone.utc)
                rec["finished_at"] = finished_at.isoformat()
                rec["status"] = status
                rec["error_message"] = (error_message or "")[:500]

                started_str = rec.get("started_at")
                if started_str:
                    try:
                        started_at = datetime.fromisoformat(started_str)
                        rec["duration_seconds"] = round(
                            (finished_at - started_at).total_seconds(), 1
                        )
                    except Exception:
                        pass
                break
        _save(records)


def get_recent(limit: int = 5) -> list[dict]:
    """Возвращает последние N запусков с человекочитаемыми полями."""
    records = _load()[:limit]
    return [_enrich(r) for r in records]


def get_all(limit: int = MAX_RECORDS) -> list[dict]:
    """Все записи (с ограничением)."""
    records = _load()[:limit]
    return [_enrich(r) for r in records]


def _enrich(rec: dict) -> dict:
    """Добавляет человекочитаемые поля для UI."""
    out = dict(rec)
    out["started_at_human"] = _format_relative(rec.get("started_at"))
    out["duration_human"] = _format_duration(rec.get("duration_seconds"))
    return out


def _format_relative(iso_str: Optional[str]) -> str:
    if not iso_str:
        return "—"
    try:
        dt = datetime.fromisoformat(iso_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        now = datetime.now(timezone.utc)
        diff = now - dt
        seconds = int(diff.total_seconds())

        if seconds < 60:
            return "только что"
        if seconds < 3600:
            mins = seconds // 60
            return f"{mins} мин назад"
        if seconds < 86400:
            hours = seconds // 3600
            word = _plural(hours, "час", "часа", "часов")
            return f"{hours} {word} назад"
        if seconds < 86400 * 2:
            return f"вчера в {dt.astimezone().strftime('%H:%M')}"
        if seconds < 86400 * 7:
            days = seconds // 86400
            word = _plural(days, "день", "дня", "дней")
            return f"{days} {word} назад"
        return dt.astimezone().strftime("%d.%m.%Y %H:%M")
    except Exception:
        return iso_str or "—"


def _format_duration(seconds: Optional[float]) -> str:
    if seconds is None:
        return ""
    try:
        s = int(seconds)
        if s < 60:
            return f"{s} с"
        m, s = divmod(s, 60)
        if m < 60:
            return f"{m} мин {s:02d} с"
        h, m = divmod(m, 60)
        return f"{h} ч {m:02d} мин"
    except Exception:
        return ""


def _plural(n: int, one: str, few: str, many: str) -> str:
    n_abs = abs(n) % 100
    if 11 <= n_abs <= 14:
        return many
    n_mod = n_abs % 10
    if n_mod == 1:
        return one
    if 2 <= n_mod <= 4:
        return few
    return many
