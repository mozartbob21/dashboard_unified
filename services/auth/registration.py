# services/auth/registration.py
"""
Регистрация пользователей без подтверждения по почте
и персональная «память аккаунта» (user_settings).
"""
from __future__ import annotations

import json
import re
import sqlite3
from datetime import datetime, timezone

from utils.db import get_db_connection

DEFAULT_ROLE = "Пользователь"
DEFAULT_MODULES = []  # Блоки назначает управляющий учётными записями.

USERNAME_RE = re.compile(r"^[a-zA-Zа-яА-ЯёЁ0-9_.\-]{3,32}$")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _now():
    return datetime.now(timezone.utc)


def _now_iso() -> str:
    return _now().isoformat(timespec="seconds")


def ensure_registration_tables() -> None:
    with get_db_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS user_settings (
                user_id    INTEGER PRIMARY KEY,
                data       TEXT NOT NULL DEFAULT '{}',
                updated_at TEXT NOT NULL
            );
            """
        )
        
        # Миграция: добавляем email в users, если его нет
        try:
            conn.execute("SELECT email FROM users LIMIT 1")
        except sqlite3.OperationalError:
            conn.execute("ALTER TABLE users ADD COLUMN email TEXT")
            print("[registration] Добавлена колонка email в таблицу users")


# создаём таблицы при импорте
ensure_registration_tables()


def _username_taken(conn, username: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM users WHERE lower(username) = lower(?)", (username,)
    ).fetchone() is not None


def _email_taken(conn, email: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM users WHERE lower(email) = lower(?)", (email,)
    ).fetchone() is not None


def register_user(username: str, email: str, password_hash: str):
    """Сразу создаёт аккаунт без доступа к блокам: их назначает управляющий."""
    username = (username or "").strip()
    email = (email or "").strip().lower()

    if not USERNAME_RE.fullmatch(username):
        return False, "Логин: 3–32 символа (буквы, цифры, _ . -)"
    if not EMAIL_RE.fullmatch(email):
        return False, "Некорректный e-mail"

    with get_db_connection() as conn:
        # Проверка уникальности и создание аккаунта — одна транзакция, включая
        # email, для которого в старой схеме нет уникального индекса.
        conn.execute("BEGIN IMMEDIATE")
        if _username_taken(conn, username):
            return False, "Такой логин уже занят"
        if _email_taken(conn, email):
            return False, "На этот e-mail уже есть аккаунт"

        cur = conn.execute(
            """INSERT INTO users (username, email, password_hash, role, modules, is_active, created_at)
               VALUES (?, ?, ?, ?, ?, 1, ?)""",
            (
                username,
                email,
                password_hash,
                DEFAULT_ROLE,
                json.dumps(DEFAULT_MODULES, ensure_ascii=False),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        user_id = cur.lastrowid
        conn.execute(
            "INSERT INTO user_settings (user_id, data, updated_at) VALUES (?, '{}', ?)",
            (user_id, _now_iso()),
        )

    return True, {"id": user_id, "username": username, "role": DEFAULT_ROLE}


# ─── Память аккаунта ────────────────────────────────────────────────

def get_settings(user_id: int) -> dict:
    with get_db_connection() as conn:
        row = conn.execute(
            "SELECT data FROM user_settings WHERE user_id = ?", (user_id,)
        ).fetchone()
    if not row:
        return {}
    try:
        return json.loads(row["data"] or "{}")
    except json.JSONDecodeError:
        return {}


def save_settings(user_id: int, patch: dict) -> dict:
    """Сливает новые настройки с текущими и сохраняет."""
    current = get_settings(user_id)
    current.update(patch or {})
    with get_db_connection() as conn:
        conn.execute(
            """INSERT INTO user_settings (user_id, data, updated_at) VALUES (?, ?, ?)
               ON CONFLICT(user_id) DO UPDATE SET
                   data = excluded.data,
                   updated_at = excluded.updated_at""",
            (user_id, json.dumps(current, ensure_ascii=False), _now_iso()),
        )
    return current
