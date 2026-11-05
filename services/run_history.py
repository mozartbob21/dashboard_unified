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
import re
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
    "presentations": "Еженедельные презентации",
    "omsu-rating": "Рейтинг ОМСУ",
    "mingkh-water-map": "МИНЖКХ — карта жалоб",
    "edds": "ЕДДС — контроль заявок",
    "edo":                  "Заполненность данных",
    "overdue":              "Просроченные задачи",
    "watercontrol":         "Контроль воды",
    "utnkr":                "Технадзор УТНКР",
    "cameras":              "Проверка камер",
    "camera_prescriptions": "Предписания по камерам",
    "cds":                  "Обращения 1С",
    "mgkh_rm":              "Задачи МИНЖКХ",
    "water_rm":             "Проверка задач по воде",
    "ecur":                 "ЕЦУР — жалобы",
    "appeals":              "Эмпатичные ответы",
    "municipality-report":  "Отчёт по округу",
    "water-dashboard":      "Качество водоснабжения",
    "mingkh":               "Дашборд по жалобам МинЖКХ",
    "collective":           "Коллективные жалобы",
    "tools":                "Полезные инструменты",
    "summarizer":           "Сумматор",
    "zips":                 "Остатки ЗиП РСО",
    "telegram":             "Телеграм — выбранные чаты",
    "aichat":               "ИИ-чат",
}

MODULE_ICONS = {
    "presentations": "📄",
    "omsu-rating": "📊",
    "edds":                 "🚨",
    "mingkh-water-map": "🗺️",
    "edo":                  "📋",
    "overdue":              "⏱️",
    "watercontrol":         "💧",
    "utnkr":                "🛡️",
    "cameras":              "📹",
    "camera_prescriptions": "📄",
    "cds":                  "🏠",
    "mgkh_rm":              "📅",
    "water_rm":             "💧",
    "ecur":                 "📮",
    "appeals":              "✍️",
    "municipality-report":  "📊",
    "water-dashboard":      "🌊",
    "mingkh":               "📊",
    "collective":           "👥",
    "tools":                "🧰",
    "summarizer":           "📝",
    "zips":                 "📦",
    "telegram":             "💬",
    "aichat":               "🤖",
}

# Старые и служебные имена остаются в записи для фильтрации прав;
# подпись всегда берётся из текущего русскоязычного каталога.
_HISTORY_LABEL_ALIASES = {
    "mgkh-rm": "mgkh_rm",
    "water_dashboard": "water-dashboard",
}
DEFAULT_MODULE_LABEL = "Модуль платформы"


def _module_presentation(module_id: str) -> tuple[str, str]:
    canonical_id = _HISTORY_LABEL_ALIASES.get(module_id, module_id)
    return (
        MODULE_LABELS.get(canonical_id, DEFAULT_MODULE_LABEL),
        MODULE_ICONS.get(canonical_id, "⚙️"),
    )

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
        if not isinstance(raw, list) or not raw or any(not isinstance(rec, dict) for rec in raw):
            # An unfamiliar archive must stay intact for the operator.
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
                    "Проверка завершилась ошибкой." if rec.get("status") == "error" else "",
                ),
            )
            migrated += 1

    # Переименовываем, чтобы больше не мигрировать
    backup_path = HISTORY_FILE.with_suffix(".json.bak")
    if backup_path.exists():
        return  # Preserve any previous backup; replay is idempotent by run ID.
    try:
        HISTORY_FILE.rename(backup_path)
    except OSError:
        return  # Committed rows are durable; keep the source if it cannot move.
    print(f"[run_history] Мигрировано {migrated} записей из JSON → SQLite")
    print(f"[run_history] Старый файл: {backup_path}")


# Выполняем миграцию при импорте модуля (один раз за процесс)
_migrate_json_if_exists()


# ─── Публичный API (сигнатуры НЕ изменились) ───────────────────────

def record_start(module_id: str, user: str = "—", *, params=None, private_owner=None) -> str:
    """Записывает старт запуска. Возвращает run_id."""
    run_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc).isoformat()
    module_label, module_icon = _module_presentation(module_id)

    from services import check_history
    check_history.claim(run_id)
    try:
        with get_db_connection() as conn:
            check_history.initialize(conn)
            conn.execute('BEGIN IMMEDIATE')
            conn.execute(
                """INSERT INTO run_history
                   (run_id,module_id,module_label,module_icon,user,started_at,status,error_message)
                   VALUES (?,?,?,?,?,?,'running','')""",
                (run_id,module_id,module_label,module_icon,
                 user if isinstance(user,str) and (user == '—' or re.fullmatch(r'[\w.-]{1,64}',user))
                 else 'Пользователь',now))
            check_history.begin(conn,run_id,module_id,params,private_owner)
    except BaseException:
        check_history.release(run_id)
        raise
    return run_id


def record_finish(run_id: str, status: str = "success", error_message: str = "",
                  *, metrics=None, result_code=None, result_params=None) -> None:
    """Finalize once. The journal stores fixed diagnostics instead of remote text."""
    from services import check_history
    if status not in ("success", "partial", "error", "skipped", "interrupted"):
        status = "error"
    finished_at = datetime.now(timezone.utc)
    try:
        with get_db_connection() as conn:
            check_history.initialize(conn)
            conn.execute('BEGIN IMMEDIATE')
            row = conn.execute("SELECT started_at,status FROM run_history WHERE run_id=?", (run_id,)).fetchone()
            if not row or row['status'] != 'running':
                return
            status = check_history.finish(conn,run_id,status,metrics,result_code,result_params)
            reason = {'partial':'Данные подтверждены частично.', 'error':'Проверка завершилась ошибкой.','interrupted':'Результат не подтверждён.',
                      'skipped':'Повторная проверка не выполнялась.'}.get(status,'')
            if status == 'error' and error_message:
                from services.water_dashboard.diagnostics import MESSAGES
                if error_message in MESSAGES.values():
                    reason = error_message
            try:
                duration = round((finished_at - datetime.fromisoformat(row['started_at'])).total_seconds(), 1)
            except (TypeError,ValueError):
                duration = None
            conn.execute("""UPDATE run_history SET finished_at=?,duration_seconds=?,status=?,error_message=?
                             WHERE run_id=?""", (finished_at.isoformat(),duration,status,reason,run_id))
    finally:
        check_history.release(run_id)


def get_recent(limit: int = 5, modules: list[str] | None = None) -> list[dict]:
    """Возвращает последние N запусков с человекочитаемыми полями."""
    if modules is not None and not modules:
        return []
    from services import check_history
    check_history.recover()
    where = " WHERE c.private_subject IS NULL"
    if modules is not None:
        where += " AND r.module_id IN (" + ",".join("?" for _ in modules) + ")"
    params = [] if modules is None else list(modules)
    with get_db_connection() as conn:
        rows = conn.execute(
            "SELECT r.* FROM run_history r LEFT JOIN check_results c USING(run_id)" + where + " ORDER BY r.started_at DESC LIMIT ?",
            [*params, max(1, min(int(limit), MAX_RECORDS))],
        ).fetchall()
    return [_enrich(_row_to_dict(r)) for r in rows]


def get_all(limit: int = MAX_RECORDS, modules: list[str] | None = None) -> list[dict]:
    """Все записи (с ограничением)."""
    return get_recent(limit, modules=modules)


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
    out["module_label"], out["module_icon"] = _module_presentation(rec.get("module_id", ""))
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
