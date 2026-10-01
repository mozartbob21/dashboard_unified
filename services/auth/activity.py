"""Aggregate account usage without keeping request data or individual event logs.

A visit starts after 30 minutes without a qualifying event in that scope.
Only HTML opens and HTTP 2xx action requests count, not polling or scheduled jobs.
"""
from datetime import datetime, timedelta, timezone
import logging

from fastapi import HTTPException
from core.roles import MODULE_ALIASES, MODULE_NAMES
from utils.db import get_db_connection

VISIT_GAP_MINUTES = 30
PLATFORM = "__platform__"
SECTION_NAMES = {
    **MODULE_NAMES,
    "aichat": "ИИ-чат и библиотека промптов",
    "scheduler": "Автозапуск",
    "voice": "Голосовые заметки",
    "users": "Управление пользователями",
}
SHARED_PATHS = {
    "/aichat": "aichat", "/prompthub": "aichat", "/api/assistant": "aichat",
    "/scheduler": "scheduler", "/api/scheduler": "scheduler",
    "/voice": "voice", "/api/voice": "voice",
    "/users": "users", "/api/users": "users",
}


def _event(request, response, module_paths):
    """Classify the response; never inspect bodies, query strings or credentials."""
    if not 200 <= response.status_code < 300:
        return None
    if "prefetch" in (request.headers.get("purpose", "") + request.headers.get("sec-purpose", "")).lower():
        return None
    path = request.url.path.rstrip("/") or "/"
    module_id = None
    for prefix, candidate in sorted({**module_paths, **SHARED_PATHS}.items(), key=lambda pair: -len(pair[0])):
        if path == prefix or path.startswith(prefix + "/"):
            module_id = candidate
            break
    if module_id not in SECTION_NAMES and path != "/":
        return None
    if request.method == "GET":
        if (not response.headers.get("content-type", "").lower().startswith("text/html")
                or "attachment" in response.headers.get("content-disposition", "").lower()
                or "/api/" in path or path.startswith("/api/")):
            return None
        return module_id, "page_view"
    if request.method in {"POST", "PUT", "PATCH", "DELETE"} and module_id:
        if path == "/api/users/notifications/read":
            return None
        # Manual scheduler operations belong to the affected module. Timer jobs
        # never enter this middleware and are not attributed to any account.
        if path.startswith("/api/scheduler/jobs/"):
            candidate = path.split("/")[4]
            module_id = MODULE_ALIASES.get(candidate, candidate)
            if module_id not in MODULE_NAMES:
                return None
        return module_id, "action"
    return None


def record_activity(user, module_id, event, *, now=None):
    """Atomic counters keep simultaneous requests from creating extra visits."""
    if not user or user.get("kc_sub") or type(user.get("id")) is not int:
        return
    if event not in {"page_view", "action"} or (module_id is not None and module_id not in SECTION_NAMES):
        return
    instant = now or datetime.now(timezone.utc)
    stamp = instant.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    cutoff = (instant - timedelta(minutes=VISIT_GAP_MINUTES)).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    with get_db_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        if not conn.execute(
            "SELECT 1 FROM users WHERE id=? AND is_active=1 "
            "AND NOT EXISTS (SELECT 1 FROM account_archives WHERE user_id=users.id)",
            (user["id"],),
        ).fetchone():
            return
        for scope in [PLATFORM] + ([module_id] if module_id else []):
            conn.execute(
                """INSERT INTO user_activity
                   (user_id,module_id,visits,page_views,actions,first_seen_at,last_seen_at)
                   VALUES (?,?,1,?,?,?,?)
                   ON CONFLICT(user_id,module_id) DO UPDATE SET
                     visits = visits + CASE WHEN last_seen_at <= ? THEN 1 ELSE 0 END,
                     page_views = page_views + excluded.page_views,
                     actions = actions + excluded.actions,
                     first_seen_at = min(first_seen_at,excluded.first_seen_at),
                     last_seen_at = max(last_seen_at,excluded.last_seen_at)""",
                (user["id"], scope, int(event == "page_view"), int(event == "action"), stamp, stamp, cutoff),
            )


def track_response(request, response, module_paths):
    """Unavailable statistics must never break an otherwise successful operation."""
    try:
        event = _event(request, response, module_paths)
        if event is not None:
            record_activity(getattr(request.state, "user", None), *event)
    except Exception:
        # Exception messages and URLs can contain user data; keep both out.
        logging.getLogger(__name__).warning("Activity counters could not be updated")


def _empty():
    return {"visits": 0, "page_views": 0, "actions": 0, "first_seen_at": None, "last_seen_at": None}


def account_summaries():
    with get_db_connection() as conn:
        rows = conn.execute("SELECT * FROM user_activity WHERE module_id=?", (PLATFORM,)).fetchall()
    return {row["user_id"]: {key: row[key] for key in _empty()} for row in rows}


def account_activity(user_id):
    with get_db_connection() as conn:
        if not conn.execute("SELECT 1 FROM users WHERE id=?", (user_id,)).fetchone():
            raise HTTPException(404, "Пользователь не найден")
        rows = conn.execute("SELECT * FROM user_activity WHERE user_id=?", (user_id,)).fetchall()
        started_at = conn.execute("SELECT started_at FROM user_activity_collection WHERE id=1").fetchone()[0]
    counts = {row["module_id"]: {key: row[key] for key in _empty()} for row in rows}
    return {
        "user_id": user_id, "tracking_since": started_at,
        "visit_gap_minutes": VISIT_GAP_MINUTES,
        "platform": counts.get(PLATFORM, _empty()),
        "modules": [
            {"module_id": module, "label": label, **counts.get(module, _empty())}
            for module, label in SECTION_NAMES.items()
        ],
    }
