# services/auth/registration.py
"""
Регистрация пользователей с подтверждением по почте
и персональная «память аккаунта» (user_settings).
"""
from __future__ import annotations

import hashlib
import json
import re
import secrets
import sqlite3
from datetime import datetime, timedelta, timezone

from utils.db import get_db_connection

CODE_TTL_MINUTES = 10
MAX_ATTEMPTS = 5
RESEND_COOLDOWN_SECONDS = 60

DEFAULT_ROLE = "Пользователь"
DEFAULT_MODULES = ["edo", "overdue", "watercontrol"]

USERNAME_RE = re.compile(r"^[a-zA-Zа-яА-ЯёЁ0-9_.\-]{3,32}$")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _now():
    return datetime.now(timezone.utc)


def _now_iso() -> str:
    return _now().isoformat(timespec="seconds")


def _hash_code(code: str) -> str:
    return hashlib.sha256(code.encode()).hexdigest()


def _ensure_dt(value: str) -> datetime:
    dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def ensure_registration_tables() -> None:
    with get_db_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS pending_registrations (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                username      TEXT NOT NULL,
                email         TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                code_hash     TEXT NOT NULL,
                attempts      INTEGER NOT NULL DEFAULT 0,
                created_at    TEXT NOT NULL,
                expires_at    TEXT NOT NULL,
                last_send_at  TEXT
            );
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


def _find_pending(conn, identifier: str):
    return conn.execute(
        """SELECT * FROM pending_registrations
           WHERE lower(username) = lower(?) OR lower(email) = lower(?)""",
        (identifier, identifier),
    ).fetchone()


def start_registration(username: str, email: str, password_hash: str):
    """Создаёт заявку и возвращает (ok, dict с code/email) или (False, сообщение)."""
    username = (username or "").strip()
    email = (email or "").strip().lower()

    if not USERNAME_RE.match(username):
        return False, "Логин: 3–32 символа (буквы, цифры, _ . -)"
    if not EMAIL_RE.match(email):
        return False, "Некорректный e-mail"

    with get_db_connection() as conn:
        if _username_taken(conn, username):
            return False, "Такой логин уже занят"
        if _email_taken(conn, email):
            return False, "На этот e-mail уже есть аккаунт"

        conn.execute(
            "DELETE FROM pending_registrations WHERE lower(username)=lower(?) OR lower(email)=lower(?)",
            (username, email),
        )

        code = f"{secrets.randbelow(1_000_000):06d}"
        conn.execute(
            """INSERT INTO pending_registrations
               (username, email, password_hash, code_hash, attempts, created_at, expires_at, last_send_at)
               VALUES (?, ?, ?, ?, 0, ?, ?, ?)""",
            (
                username, email, password_hash, _hash_code(code),
                _now_iso(),
                (_now() + timedelta(minutes=CODE_TTL_MINUTES)).isoformat(timespec="seconds"),
                _now_iso(),
            ),
        )
    return True, {"username": username, "email": email, "code": code}


def resend_code(identifier: str):
    identifier = (identifier or "").strip()
    with get_db_connection() as conn:
        row = _find_pending(conn, identifier)
        if not row:
            return False, "Заявка на регистрацию не найдена"

        if row["last_send_at"]:
            delta = (_now() - _ensure_dt(row["last_send_at"])).total_seconds()
            if delta < RESEND_COOLDOWN_SECONDS:
                return False, "Письмо уже отправлено недавно — подождите минуту"

        code = f"{secrets.randbelow(1_000_000):06d}"
        conn.execute(
            """UPDATE pending_registrations
               SET code_hash = ?, attempts = 0, expires_at = ?, last_send_at = ?
               WHERE id = ?""",
            (
                _hash_code(code),
                (_now() + timedelta(minutes=CODE_TTL_MINUTES)).isoformat(timespec="seconds"),
                _now_iso(),
                row["id"],
            ),
        )
    return True, {"email": row["email"], "code": code}


def verify_registration(identifier: str, code: str):
    """Проверяет код и создаёт пользователя. Возвращает (ok, user dict / сообщение)."""
    identifier = (identifier or "").strip()
    code = (code or "").strip()

    with get_db_connection() as conn:
        row = _find_pending(conn, identifier)
        if not row:
            return False, "Заявка на регистрацию не найдена"

        if _now() > _ensure_dt(row["expires_at"]):
            conn.execute("DELETE FROM pending_registrations WHERE id = ?", (row["id"],))
            return False, "Код просрочен. Запросите новый."

        if row["attempts"] >= MAX_ATTEMPTS:
            conn.execute("DELETE FROM pending_registrations WHERE id = ?", (row["id"],))
            return False, "Слишком много попыток. Запросите новый код."

        if _hash_code(code) != row["code_hash"]:
            conn.execute(
                "UPDATE pending_registrations SET attempts = attempts + 1 WHERE id = ?",
                (row["id"],),
            )
            return False, "Неверный код"

        if _username_taken(conn, row["username"]):
            conn.execute("DELETE FROM pending_registrations WHERE id = ?", (row["id"],))
            return False, "Логин уже занят"

        cur = conn.execute(
            """INSERT INTO users (username, email, password_hash, role, modules, is_active, created_at)
               VALUES (?, ?, ?, ?, ?, 1, ?)""",
            (
                row["username"],
                row["email"],
                row["password_hash"],
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
        conn.execute("DELETE FROM pending_registrations WHERE id = ?", (row["id"],))

    return True, {"id": user_id, "username": row["username"], "role": DEFAULT_ROLE}


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