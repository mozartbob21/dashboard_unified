"""
Планировщик автозапуска модулей.

- Хранит расписание в SQLite (scheduler_jobs)
- Фоновый поток каждые 30 сек проверяет, пора ли запускать
- Запускает модули через subprocess (как кнопка «Запустить»)
- REST API для управления расписанием из UI
"""

from __future__ import annotations

import sys
import threading
import time
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse

from utils.db import get_db_connection
from services.run_history import MODULE_LABELS, MODULE_ICONS

# ─── Конфигурация ───────────────────────────────────────────────────

TICK_INTERVAL = 30  # секунд между проверками расписания

# Какие модули можно запускать по расписанию
# module_id → команда запуска (аналогично app.py)
MODULE_COMMANDS: dict[str, list[str]] = {
    "edo":                  [sys.executable, "-m", "services.edo.runner"],
    "overdue":              [sys.executable, "-m", "services.overdue.runner"],
    "watercontrol":         [sys.executable, "-m", "services.watercontrol.runner"],
    "utnkr":                [sys.executable, "-m", "services.utnkr.scanner"],
    "cameras":              [sys.executable, "-m", "services.cameras.camera_checker"],
    "camera_prescriptions": [sys.executable, "-m", "services.cameras.prescription_generator"],
    "cds":                  [sys.executable, "-m", "services.cds.runner"],
}

# Интервалы по умолчанию (минуты) для первичного заполнения
DEFAULT_INTERVALS: dict[str, int] = {
    "edo": 120,
    "overdue": 120,
    "watercontrol": 120,
    "utnkr": 180,
    "cameras": 240,
    "camera_prescriptions": 360,
    "cds": 1440,
}

# ─── Инициализация таблицы (seed) ───────────────────────────────────

_SEEDED = False
_SEED_LOCK = threading.Lock()


def _seed_jobs() -> None:
    """
    При первом запуске заполняет scheduler_jobs записями
    для всех известных модулей (enabled=0 — выключены).
    Существующие записи НЕ перезаписываются.
    """
    global _SEEDED
    if _SEEDED:
        return
    with _SEED_LOCK:
        if _SEEDED:
            return

        with get_db_connection() as conn:
            for module_id, command in MODULE_COMMANDS.items():
                label = MODULE_LABELS.get(module_id, module_id)
                icon = MODULE_ICONS.get(module_id, "⚙️")

                exists = conn.execute(
                    "SELECT 1 FROM scheduler_jobs WHERE module_id = ?",
                    (module_id,),
                ).fetchone()
                if exists:
                    # Синхронизируем название/иконку из кода (переименования)
                    conn.execute(
                        """
                        UPDATE scheduler_jobs
                           SET module_label = ?, module_icon = ?
                         WHERE module_id = ?
                        """,
                        (label, icon, module_id),
                    )
                    continue

                conn.execute(
                    """
                    INSERT INTO scheduler_jobs
                        (module_id, module_label, module_icon,
                         command, enabled, interval_minutes)
                    VALUES (?, ?, ?, ?, 0, ?)
                    """,
                    (
                        module_id,
                        label,
                        icon,
                        " ".join(command),
                        DEFAULT_INTERVALS.get(module_id, 120),
                    ),
                )

        _SEEDED = True
        print("[scheduler] Jobs seeded")


# ─── Утилиты ────────────────────────────────────────────────────────

def _now() -> datetime:
    return datetime.now(timezone.utc)


def _now_iso() -> str:
    return _now().isoformat()


def _row_to_dict(row) -> dict:
    return {
        "module_id": row["module_id"],
        "module_label": row["module_label"],
        "module_icon": row["module_icon"] or "⚙️",
        "command": row["command"],
        "enabled": bool(row["enabled"]),
        "interval_minutes": row["interval_minutes"],
        "last_run_at": row["last_run_at"],
        "next_run_at": row["next_run_at"],
        "last_status": row["last_status"],
        "last_error": row["last_error"] or "",
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


# ─── Публичный API (для использования из других модулей) ────────────

def get_all_jobs() -> list[dict]:
    _seed_jobs()
    with get_db_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM scheduler_jobs ORDER BY module_id"
        ).fetchall()
    return [_row_to_dict(r) for r in rows]


def get_job(module_id: str) -> Optional[dict]:
    _seed_jobs()
    with get_db_connection() as conn:
        row = conn.execute(
            "SELECT * FROM scheduler_jobs WHERE module_id = ?",
            (module_id,),
        ).fetchone()
    return _row_to_dict(row) if row else None


def set_enabled(module_id: str, enabled: bool) -> dict:
    """Включает/выключает автозапуск. При включении ставит next_run_at."""
    _seed_jobs()
    now = _now()

    with get_db_connection() as conn:
        row = conn.execute(
            "SELECT * FROM scheduler_jobs WHERE module_id = ?",
            (module_id,),
        ).fetchone()
        if not row:
            raise ValueError(f"Unknown module: {module_id}")

        if enabled:
            interval = row["interval_minutes"] or 60
            next_run = (now + timedelta(minutes=interval)).isoformat()
        else:
            next_run = None

        conn.execute(
            """
            UPDATE scheduler_jobs
               SET enabled = ?,
                   next_run_at = ?,
                   updated_at = ?
             WHERE module_id = ?
            """,
            (int(enabled), next_run, now.isoformat(), module_id),
        )

    return get_job(module_id)


def set_interval(module_id: str, interval_minutes: int) -> dict:
    """Изменяет интервал. Пересчитывает next_run_at если включён."""
    _seed_jobs()
    if interval_minutes < 5:
        raise ValueError("Минимальный интервал — 5 минут")
    if interval_minutes > 1440:
        raise ValueError("Максимальный интервал — 1440 минут (24 часа)")

    now = _now()

    with get_db_connection() as conn:
        row = conn.execute(
            "SELECT * FROM scheduler_jobs WHERE module_id = ?",
            (module_id,),
        ).fetchone()
        if not row:
            raise ValueError(f"Unknown module: {module_id}")

        # Пересчитываем next_run если задача включена
        next_run = row["next_run_at"]
        if row["enabled"]:
            last = row["last_run_at"]
            if last:
                try:
                    base = datetime.fromisoformat(last)
                    next_run = (base + timedelta(minutes=interval_minutes)).isoformat()
                except Exception:
                    next_run = (now + timedelta(minutes=interval_minutes)).isoformat()
            else:
                next_run = (now + timedelta(minutes=interval_minutes)).isoformat()

        conn.execute(
            """
            UPDATE scheduler_jobs
               SET interval_minutes = ?,
                   next_run_at = ?,
                   updated_at = ?
             WHERE module_id = ?
            """,
            (interval_minutes, next_run, now.isoformat(), module_id),
        )

    return get_job(module_id)


def mark_run_started(module_id: str) -> None:
    """Отмечает что модуль запущен планировщиком."""
    now = _now()
    with get_db_connection() as conn:
        row = conn.execute(
            "SELECT interval_minutes FROM scheduler_jobs WHERE module_id = ?",
            (module_id,),
        ).fetchone()
        interval = (row["interval_minutes"] if row else 60) or 60
        next_run = (now + timedelta(minutes=interval)).isoformat()

        conn.execute(
            """
            UPDATE scheduler_jobs
               SET last_run_at = ?,
                   next_run_at = ?,
                   last_status = 'running',
                   last_error = '',
                   updated_at = ?
             WHERE module_id = ?
            """,
            (now.isoformat(), next_run, now.isoformat(), module_id),
        )


def mark_run_finished(module_id: str, status: str, error: str = "") -> None:
    """Отмечает результат запуска."""
    with get_db_connection() as conn:
        conn.execute(
            """
            UPDATE scheduler_jobs
               SET last_status = ?,
                   last_error = ?,
                   updated_at = ?
             WHERE module_id = ?
            """,
            (status, (error or "")[:500], _now_iso(), module_id),
        )


# ─── Фоновый поток ──────────────────────────────────────────────────

_thread: Optional[threading.Thread] = None
_stop_event = threading.Event()


def _get_run_status_dict() -> dict:
    """
    Безопасно получает run_status из app.py.
    Импортируем лениво, чтобы избежать циклических импортов.
    """
    try:
        from app import run_status
        return run_status
    except ImportError:
        return {}


def _is_module_running(module_id: str) -> bool:
    """Проверяет, не запущен ли модуль уже (через run_status в app.py)."""
    rs = _get_run_status_dict()
    status = rs.get(module_id, {})
    return bool(status.get("running", False))


def _launch_module(module_id: str) -> None:
    """Запускает модуль через subprocess (тот же путь, что и кнопка в UI)."""
    if module_id not in MODULE_COMMANDS:
        print(f"[scheduler] Unknown module: {module_id}")
        return

    if _is_module_running(module_id):
        print(f"[scheduler] {module_id} already running, skip")
        return

    command = MODULE_COMMANDS[module_id]
    print(f"[scheduler] Launching {module_id}: {' '.join(command)}")

    mark_run_started(module_id)

    try:
        from app import start_background_service
        result = start_background_service(module_id, command)
        if isinstance(result, JSONResponse):
            # Уже запущен — не ошибка
            mark_run_finished(module_id, "skipped", "already running")
        else:
            # start_background_service возвращает {"ok": True}
            # Реальный результат придёт когда subprocess завершится
            # и run_subprocess_worker вызовет run_history.record_finish
            pass
    except Exception as e:
        print(f"[scheduler] Error launching {module_id}: {e}")
        mark_run_finished(module_id, "error", str(e))


def _tick() -> None:
    """Один цикл проверки расписания."""
    now = _now()

    with get_db_connection() as conn:
        rows = conn.execute(
            """
            SELECT module_id, next_run_at
              FROM scheduler_jobs
             WHERE enabled = 1
               AND next_run_at IS NOT NULL
            """
        ).fetchall()

    for row in rows:
        module_id = row["module_id"]
        next_run_str = row["next_run_at"]

        if not next_run_str:
            continue

        try:
            next_run = datetime.fromisoformat(next_run_str)
            if next_run.tzinfo is None:
                next_run = next_run.replace(tzinfo=timezone.utc)
        except Exception:
            continue

        if now >= next_run:
            _launch_module(module_id)


def _scheduler_loop() -> None:
    """Основной цикл фонового потока."""
    print(f"[scheduler] Background thread started (tick every {TICK_INTERVAL}s)")
    _seed_jobs()

    while not _stop_event.is_set():
        try:
            _tick()
        except Exception as e:
            print(f"[scheduler] Tick error: {e}")

        _stop_event.wait(TICK_INTERVAL)

    print("[scheduler] Background thread stopped")


def start() -> None:
    """Запускает фоновый поток планировщика."""
    global _thread
    if _thread is not None and _thread.is_alive():
        print("[scheduler] Already running")
        return

    _stop_event.clear()
    _thread = threading.Thread(
        target=_scheduler_loop,
        name="scheduler",
        daemon=True,
    )
    _thread.start()


def stop() -> None:
    """Останавливает фоновый поток."""
    _stop_event.set()
    if _thread is not None:
        _thread.join(timeout=5)
    print("[scheduler] Stopped")


def is_running() -> bool:
    return _thread is not None and _thread.is_alive()


# ─── FastAPI Router ─────────────────────────────────────────────────

router = APIRouter(prefix="/api/scheduler", tags=["scheduler"])


@router.get("/jobs")
async def api_get_jobs():
    """Список всех задач планировщика."""
    return {
        "ok": True,
        "scheduler_running": is_running(),
        "jobs": get_all_jobs(),
    }


@router.get("/jobs/{module_id}")
async def api_get_job(module_id: str):
    """Одна задача по module_id."""
    job = get_job(module_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"Job not found: {module_id}")
    return {"ok": True, "job": job}


@router.post("/jobs/{module_id}/enable")
async def api_enable_job(module_id: str):
    """Включить автозапуск модуля."""
    try:
        job = set_enabled(module_id, True)
        return {"ok": True, "message": f"{module_id} включён", "job": job}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/jobs/{module_id}/disable")
async def api_disable_job(module_id: str):
    """Выключить автозапуск модуля."""
    try:
        job = set_enabled(module_id, False)
        return {"ok": True, "message": f"{module_id} выключен", "job": job}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/jobs/{module_id}/interval")
async def api_set_interval(module_id: str, payload: dict):
    """Изменить интервал: {"interval_minutes": 90}."""
    interval = payload.get("interval_minutes")
    if interval is None:
        raise HTTPException(status_code=400, detail="interval_minutes required")
    try:
        interval = int(interval)
        job = set_interval(module_id, interval)
        return {"ok": True, "message": f"Интервал {module_id}: {interval} мин", "job": job}
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/jobs/{module_id}/run-now")
async def api_run_now(module_id: str):
    """Запустить модуль немедленно (вне расписания)."""
    if module_id not in MODULE_COMMANDS:
        raise HTTPException(status_code=404, detail=f"Unknown module: {module_id}")

    if _is_module_running(module_id):
        return JSONResponse(
            status_code=409,
            content={"ok": False, "message": f"{module_id} уже выполняется"},
        )

    _launch_module(module_id)
    return {"ok": True, "message": f"{module_id} запущен"}


@router.get("/status")
async def api_scheduler_status():
    """Общий статус планировщика."""
    return {
        "ok": True,
        "running": is_running(),
        "tick_interval": TICK_INTERVAL,
        "total_jobs": len(MODULE_COMMANDS),
        "enabled_jobs": sum(1 for j in get_all_jobs() if j["enabled"]),
    }