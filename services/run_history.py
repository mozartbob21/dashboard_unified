"""
Модуль истории запусков проверок.

Хранит записи о всех запусках модулей (ЭДО, просрочки, камеры и т.д.)
в SQLite-таблице run_history (через utils/db.py).

Используется в app.py: запись стартует в run_subprocess_worker()
при старте процесса и обновляется по завершению.

Миграция: при первом импорте, если существует старый data/run_history.json,
данные переносятся в БД, а файл переименовывается в .json.bak
"""

from __future__ import annotations

import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from utils.db import get_db_connection

BASE_DIR = Path(__file__).resolve().parent.parent
HISTORY_FILE = BASE_DIR / "data" / "run_history.json"
MAX_RECORDS = 500  # в БД можно хранить больше, чем в JSON

# ─── Человекочитаемые названия модулей для UI ──────────────────────

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


# ─── Миграция JSON → SQLite (один раз) ─────────────────────────────

def _migrate_json_if_exists() -> None:
    """
    Если существует старый run_history.json — переносим все записи в БД,
    затем переименовываем файл в .json.bak чтобы не мигрировать повторно.
    """
    if not HISTORY_FILE.exists():
        return

    try:
        raw = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        if not isinstance(raw, list) or len(raw) == 0:
            # Пустой или битый файл — просто переименовываем
            HISTORY_FILE.rename(HISTORY_FILE.with_suffix(".json.bak"))
            return
    except Exception as e:
        print(f"[run_history] Ошибка чтения JSON для миграции: {e}")
        return

    migrated = 0
    with get_db_connection() as conn:
        for rec in raw:
            run_id = rec.get("run_id")
            if not run_id:
                continue
            # Проверяем, нет ли уже такой записи (идемпотентность)
            exists = conn.execute(
                "SELECT 1 FROM run_history WHERE run_id = ?", (run_id,)
            ).fetchone()
            if exists:
                continue

            conn.execute(
                """
                INSERT INTO run_history
                    (run_id, module_id, module_label, module_icon,
                     user, started_at, finished_at,
                     duration_seconds, status, error_message)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    rec.get("module_id", ""),
                    rec.get("module_label", ""),
                    rec.get("module_icon", "⚙️"),
                    rec.get("user", "—"),
                    rec.get("started_at", ""),
                    rec.get("finished_at"),
                    rec.get("duration_seconds"),
                    rec.get("status", "success"),
                    rec.get("error_message", ""),
                ),
            )
            migrated += 1

    # Переименовываем, чтобы больше не мигрировать
    backup_path = HISTORY_FILE.with_suffix(".json.bak")
    HISTORY_FILE.rename(backup_path)
    print(f"[run_history] Мигрировано {migrated} записей из JSON → SQLite")
    print(f"[run_history] Старый файл: {backup_path}")


# Выполняем миграцию при импорте модуля (один раз за процесс)
_migrate_json_if_exists()


# ─── Публичный API (сигнатуры НЕ изменились) ───────────────────────

def record_start(module_id: str, user: str = "—") -> str:
    """Записывает старт запуска. Возвращает run_id."""
    run_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()

    with get_db_connection() as conn:
        conn.execute(
            """
            INSERT INTO run_history
                (run_id, module_id, module_label, module_icon,
                 user, started_at, status, error_message)
            VALUES (?, ?, ?, ?, ?, ?, 'running', '')
            """,
            (
                run_id,
                module_id,
                MODULE_LABELS.get(module_id, module_id),
                MODULE_ICONS.get(module_id, "⚙️"),
                user or "—",
                now,
            ),
        )
    return run_id


def record_finish(
    run_id: str,
    status: str = "success",
    error_message: str = "",
) -> None:
    """Обновляет запись о завершении запуска."""
    if status not in ("success", "error"):
        status = "success"

    finished_at = datetime.now(timezone.utc)

    with get_db_connection() as conn:
        # Читаем started_at для вычисления duration
        row = conn.execute(
            "SELECT started_at FROM run_history WHERE run_id = ?",
            (run_id,),
        ).fetchone()

        duration = None
        if row and row["started_at"]:
            try:
                started = datetime.fromisoformat(row["started_at"])
                duration = round((finished_at - started).total_seconds(), 1)
            except Exception:
                pass

        conn.execute(
            """
            UPDATE run_history
               SET finished_at = ?,
                   duration_seconds = ?,
                   status = ?,
                   error_message = ?
             WHERE run_id = ?
            """,
            (
                finished_at.isoformat(),
                duration,
                status,
                (error_message or "")[:500],
                run_id,
            ),
        )


def get_recent(limit: int = 5) -> list[dict]:
    """Возвращает последние N запусков с человекочитаемыми полями."""
    with get_db_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM run_history ORDER BY started_at DESC LIMIT ?",
            (limit,),
        ).fetchall()
    return [_enrich(_row_to_dict(r)) for r in rows]


def get_all(limit: int = MAX_RECORDS) -> list[dict]:
    """Все записи (с ограничением)."""
    with get_db_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM run_history ORDER BY started_at DESC LIMIT ?",
            (limit,),
        ).fetchall()
    return [_enrich(_row_to_dict(r)) for r in rows]


# ─── Внутренние helpers ─────────────────────────────────────────────

def _row_to_dict(row) -> dict:
    """sqlite3.Row → dict с теми же ключами, что были в JSON."""
    return {
        "run_id": row["run_id"],
        "module_id": row["module_id"],
        "module_label": row["module_label"] or "",
        "module_icon": row["module_icon"] or "⚙️",
        "user": row["user"] or "—",
        "started_at": row["started_at"],
        "finished_at": row["finished_at"],
        "duration_seconds": row["duration_seconds"],
        "status": row["status"] or "success",
        "error_message": row["error_message"] or "",
    }


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