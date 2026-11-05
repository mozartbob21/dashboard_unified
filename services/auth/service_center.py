"""Small cached snapshot; no hostnames, paths, environment, user data or error details."""
import shutil
import sqlite3
import threading
import time
from datetime import datetime, timezone

from utils import db

try:
    import psutil
except ImportError:
    psutil = None

_STARTED = time.monotonic()
_lock = threading.Lock()
_cached = None
_cached_at = 0.0
CACHE_SECONDS = 10


def _hardware():
    result = {"cpu_percent": None, "ram_used": None, "ram_total": None,
              "disk_used": None, "disk_total": None}
    if psutil:
        try:
            result["cpu_percent"] = psutil.cpu_percent(interval=0.12)
        except (OSError, psutil.Error):
            pass
        try:
            ram = psutil.virtual_memory()
            result.update(ram_used=ram.total - ram.available, ram_total=ram.total)
        except (OSError, psutil.Error):
            pass
    try:
        disk = shutil.disk_usage(db.DB_FILE.parent)
        result.update(disk_used=disk.used, disk_total=disk.total)
    except OSError:
        pass
    return result


def _database():
    result = {"database": "Недоступна", "chat_jobs": None}
    try:
        with db.get_db_connection() as conn:
            conn.execute("SELECT 1").fetchone()
            result["database"] = "Доступна"
            if conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='aichat_jobs'").fetchone():
                row = conn.execute("""SELECT count(CASE WHEN status='queued' THEN 1 END),
                    count(CASE WHEN status='running' THEN 1 END) FROM aichat_jobs""").fetchone()
                result["chat_jobs"] = {"queued": row[0], "running": row[1]}
    except (OSError, sqlite3.Error):
        pass
    return result


def snapshot():
    global _cached, _cached_at
    with _lock:
        now = time.monotonic()
        if _cached is None or now - _cached_at >= CACHE_SECONDS:
            _cached = {**_hardware(), **_database(), "application": "Отвечает",
                       "uptime_seconds": int(now - _STARTED),
                       "updated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                       "scope": "Процессор и память — показатели ОС, видимые процессу. В контейнере это "
                                "могут быть показатели всего узла, без учёта лимитов контейнера. "
                                "Диск — файловая система базы данных. Время работы — текущий процесс приложения."}
            _cached_at = now
        return _cached
