"""
Локальный модуль работы с SQLite.
- Единая точка получения соединения к dashboard.db
- Контекстный менеджер для автоматического закрытия и коммита
- Row factory: sqlite3.Row (чтобы обращаться к колонкам по имени)
- foreign_keys = ON
- Автоинициализация схемы при первом подключении
"""
from __future__ import annotations


import sqlite3
import threading
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator


from utils.common import DATA_DIR, ensure_dir




# Единый путь к базе для всего проекта
DB_FILE: Path = DATA_DIR / "dashboard.db"


# Флаг "схема уже проинициализирована в этом процессе"
# Чтобы CREATE TABLE IF NOT EXISTS вызывался только один раз за запуск процесса.
_SCHEMA_READY = False
_SCHEMA_LOCK = threading.Lock()




# Полная схема БД. CREATE TABLE IF NOT EXISTS — безопасно для существующих БД.
_SCHEMA_SQL = """
-- Пользователи системы
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL,
    modules TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);

-- История запусков модулей
CREATE TABLE IF NOT EXISTS run_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id TEXT UNIQUE NOT NULL,
    module_id TEXT NOT NULL,
    module_label TEXT,
    module_icon TEXT,
    user TEXT,
    started_at TEXT,
    finished_at TEXT,
    duration_seconds REAL,
    status TEXT,
    error_message TEXT
);
CREATE INDEX IF NOT EXISTS idx_run_history_module ON run_history(module_id);
CREATE INDEX IF NOT EXISTS idx_run_history_started ON run_history(started_at);

-- Справочник ответственных по муниципалитетам
CREATE TABLE IF NOT EXISTS responsibles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    municipality TEXT UNIQUE NOT NULL,
    name TEXT,
    phone TEXT
);

-- Отчёты (главная таблица модулей)
CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    module_id TEXT NOT NULL,
    created_at TEXT NOT NULL,
    summary_message TEXT,
    public_chat_message TEXT,
    public_message TEXT,
    report_text TEXT,
    source_url TEXT,
    redmine_url TEXT,
    extraction_note TEXT,
    extra_data TEXT,
    is_latest INTEGER DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_reports_module ON reports(module_id);
CREATE INDEX IF NOT EXISTS idx_reports_created ON reports(created_at);
CREATE INDEX IF NOT EXISTS idx_reports_latest ON reports(module_id, is_latest);

-- Строки данных отчёта
CREATE TABLE IF NOT EXISTS report_rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_id INTEGER NOT NULL,
    municipality TEXT,
    organization TEXT,
    responsible_name TEXT,
    responsible_phone TEXT,
    status TEXT,
    reason TEXT,
    extra_data TEXT,
    FOREIGN KEY (report_id) REFERENCES reports(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_rows_report ON report_rows(report_id);
CREATE INDEX IF NOT EXISTS idx_rows_municipality ON report_rows(municipality);

-- Персональные сообщения
CREATE TABLE IF NOT EXISTS report_personal_messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_id INTEGER NOT NULL,
    responsible_name TEXT,
    responsible_phone TEXT,
    municipality TEXT,
    organization TEXT,
    metric_name TEXT,
    status TEXT,
    message TEXT NOT NULL,
    is_edited INTEGER DEFAULT 0,
    FOREIGN KEY (report_id) REFERENCES reports(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_pm_report ON report_personal_messages(report_id);

-- Проблемы с данными
CREATE TABLE IF NOT EXISTS report_missing_data_issues (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_id INTEGER NOT NULL,
    row_index INTEGER,
    organization TEXT,
    municipality TEXT,
    responsible_name TEXT,
    responsible_phone TEXT,
    empty_fields TEXT,
    message TEXT,
    FOREIGN KEY (report_id) REFERENCES reports(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_mdi_report ON report_missing_data_issues(report_id);

-- Скриншоты
CREATE TABLE IF NOT EXISTS report_screenshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    report_id INTEGER NOT NULL,
    path TEXT NOT NULL,
    is_primary INTEGER DEFAULT 0,
    FOREIGN KEY (report_id) REFERENCES reports(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_ss_report ON report_screenshots(report_id);
"""




def _init_schema(conn: sqlite3.Connection) -> None:
    """
    Создает все таблицы и индексы, если их нет.
    Безопасно вызывать многократно благодаря IF NOT EXISTS.
    """
    conn.executescript(_SCHEMA_SQL)
    conn.commit()




def _ensure_schema_ready(conn: sqlite3.Connection) -> None:
    """
    Гарантирует, что схема инициализирована (один раз за процесс).
    Потокобезопасно благодаря _SCHEMA_LOCK.
    """
    global _SCHEMA_READY
    if _SCHEMA_READY:
        return
    with _SCHEMA_LOCK:
        if _SCHEMA_READY:
            return
        _init_schema(conn)
        _SCHEMA_READY = True




def _make_connection() -> sqlite3.Connection:
    """
    Создает соединение с БД с настройками по умолчанию:
    - row_factory = Row (доступ к колонкам по имени)
    - foreign_keys = ON
    - check_same_thread = False (FastAPI работает в нескольких потоках)
    """
    ensure_dir(DB_FILE.parent)
    conn = sqlite3.connect(
        DB_FILE,
        check_same_thread=False,
        detect_types=sqlite3.PARSE_DECLTYPES,
    )
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn




@contextmanager
def get_db_connection() -> Iterator[sqlite3.Connection]:
    """
    Контекстный менеджер для работы с БД.

    Использование:
        with get_db_connection() as conn:
            rows = conn.execute("SELECT * FROM users").fetchall()

    При выходе из блока:
    - При успехе  -> commit
    - При ошибке  -> rollback
    - В любом случае -> close

    При первом вызове за процесс — автоматически создаёт схему БД.
    """
    conn = _make_connection()
    try:
        _ensure_schema_ready(conn)
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()