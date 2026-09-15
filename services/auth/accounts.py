"""Локальные учётные записи: родительский и назначенные управляющие."""
import json
import re
import secrets
import sqlite3

from fastapi import HTTPException, Request
from core.roles import ALL_MODULE_IDS
from services.auth.security import hash_password, get_user_from_token
from utils.db import get_db_connection


def initialize_access_control():
    """Однократно переводит прежний полный доступ в явный список блоков."""
    with get_db_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        row = conn.execute("SELECT grants_migrated FROM account_control WHERE id=1").fetchone()
        if not row["grants_migrated"]:
            for user in conn.execute("SELECT id,role FROM users").fetchall():
                if user["role"].strip().lower() in {"admin", "администратор", "руководитель", "пользователь"}:
                    conn.execute("UPDATE users SET modules=? WHERE id=?", (json.dumps(ALL_MODULE_IDS), user["id"]))
            conn.execute("UPDATE account_control SET grants_migrated=1 WHERE id=1")


def is_parent_manager(user):
    if not user or user.get("kc_sub"):
        return False
    with get_db_connection() as conn:
        return conn.execute(
            "SELECT 1 FROM account_control c JOIN users u ON u.id=c.manager_user_id "
            "WHERE c.id=1 AND u.username=? AND u.is_active=1",
            (user.get("username"),),
        ).fetchone() is not None


def is_account_manager(user):
    if not user or user.get("kc_sub"):
        return False
    with get_db_connection() as conn:
        return conn.execute(
            "SELECT 1 FROM users u CROSS JOIN account_control c "
            "WHERE c.id=1 AND u.username=? AND u.is_active=1 "
            "AND (u.id=c.manager_user_id OR EXISTS (SELECT 1 FROM account_managers m WHERE m.user_id=u.id))",
            (user.get("username"),),
        ).fetchone() is not None


def require_account_manager(request: Request):
    user = getattr(request.state, "user", None)
    if not user:
        user = get_user_from_token(request.cookies.get("access_token"))
    if not is_account_manager(user):
        raise HTTPException(403, "Требуется право «Управление пользователями»")
    return user


def _password_hash(password):
    if not isinstance(password, str) or len(password) < 12 or len(password.encode("utf-8")) > 72:
        raise HTTPException(400, "Пароль: минимум 12 символов, не более 72 байт UTF-8")
    return hash_password(password)


def provision_manager(username="user_manager"):
    """Создаёт управляющего один раз; никогда не сбрасывает существующий пароль."""
    initialize_access_control()
    password = secrets.token_urlsafe(20)
    password_hash = _password_hash(password)
    with get_db_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        existing = conn.execute(
            "SELECT u.username FROM account_control c JOIN users u ON u.id=c.manager_user_id WHERE c.id=1"
        ).fetchone()
        if existing:
            return {"created": False, "username": existing["username"]}
        if conn.execute("SELECT 1 FROM users WHERE lower(username)=lower(?)", (username,)).fetchone():
            raise ValueError("Логин управляющего уже занят; существующая учётная запись не изменена")
        cur = conn.execute(
            "INSERT INTO users (username,password_hash,role,modules,is_active) VALUES (?, ?, ?, '[]', 1)",
            (username, password_hash, "Управление пользователями"),
        )
        conn.execute("UPDATE account_control SET manager_user_id=? WHERE id=1", (cur.lastrowid,))
    return {"created": True, "username": username, "password": password}


def list_accounts():
    with get_db_connection() as conn:
        rows = conn.execute(
            "SELECT u.id,u.username,u.email,u.role,u.modules,u.is_active,u.created_at, "
            "(u.id=c.manager_user_id) AS is_manager, "
            "(u.id=c.manager_user_id OR EXISTS (SELECT 1 FROM account_managers m WHERE m.user_id=u.id)) AS can_manage_users "
            "FROM users u CROSS JOIN account_control c "
            "WHERE c.id=1 ORDER BY u.username COLLATE NOCASE"
        ).fetchall()
    result = []
    for row in rows:
        item = dict(row)
        item["modules"] = json.loads(item["modules"] or "[]")
        item["is_active"] = bool(item["is_active"])
        item["is_manager"] = bool(item["is_manager"])
        item["can_manage_users"] = bool(item["can_manage_users"])
        result.append(item)
    return result


def notify_manager(message):
    with get_db_connection() as conn:
        conn.execute("INSERT INTO account_notifications(message) VALUES (?)", (message,))


def list_notifications():
    with get_db_connection() as conn:
        rows = conn.execute("SELECT * FROM account_notifications ORDER BY id DESC LIMIT 100").fetchall()
        unread = conn.execute("SELECT count(*) FROM account_notifications WHERE is_read=0").fetchone()[0]
    return {"notifications": [dict(row) for row in rows], "unread": unread}


def read_notifications(through_id):
    with get_db_connection() as conn:
        conn.execute("UPDATE account_notifications SET is_read=1 WHERE id<=?", (through_id,))


def save_account(payload, user_id=None, *, actor):
    modules = payload.modules
    if any(module not in ALL_MODULE_IDS for module in modules):
        raise HTTPException(400, "Неизвестный блок")
    modules_json = json.dumps(list(dict.fromkeys(modules)), ensure_ascii=False)
    username = payload.username.strip().lower()
    if not re.fullmatch(r"[a-zа-яё0-9_.-]{3,32}", username):
        raise HTTPException(400, "Логин: 3–32 символа, буквы, цифры, _ . -")
    email = payload.email.strip().lower() or None
    if email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
        raise HTTPException(400, "Некорректная почта")
    password_hash = _password_hash(payload.password) if payload.password else None
    try:
        with get_db_connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            parent_id = conn.execute("SELECT manager_user_id FROM account_control WHERE id=1").fetchone()[0]
            actor_row = conn.execute("SELECT id FROM users WHERE username=? AND is_active=1", (actor.get("username"),)).fetchone()
            if not actor_row or actor.get("kc_sub"):
                raise HTTPException(403, "Требуется право «Управление пользователями»")
            parent = actor_row["id"] == parent_id
            delegated = conn.execute("SELECT 1 FROM account_managers WHERE user_id=?", (actor_row["id"],)).fetchone()
            if not parent and not delegated:
                raise HTTPException(403, "Право управления пользователями отозвано")
            target_manager = user_id == parent_id or conn.execute("SELECT 1 FROM account_managers WHERE user_id=?", (user_id,)).fetchone() is not None
            if not parent and (target_manager or payload.can_manage_users is True):
                raise HTTPException(403, "Только родительская учётная запись может изменять управляющих и выдавать право управления")
            if user_id is None:
                if not password_hash:
                    raise HTTPException(400, "Укажите пароль для новой учётной записи")
                if conn.execute("SELECT 1 FROM users WHERE lower(username)=?", (username,)).fetchone():
                    raise HTTPException(409, "Такой логин уже существует")
                cur = conn.execute(
                    "INSERT INTO users (username,email,password_hash,role,modules,is_active) VALUES (?,?,?,'Пользователь',?,?)",
                    (username, email, password_hash, modules_json, int(payload.is_active)),
                )
                user_id = cur.lastrowid
            else:
                row = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
                if not row:
                    raise HTTPException(404, "Пользователь не найден")
                if username != row["username"].lower():
                    raise HTTPException(400, "Логин существующего пользователя менять нельзя")
                manager = conn.execute("SELECT manager_user_id FROM account_control WHERE id=1").fetchone()[0]
                if user_id == manager and (not payload.is_active or modules or payload.can_manage_users is False):
                    raise HTTPException(400, "Учётная запись управления должна оставаться активной без рабочих блоков")
                conn.execute(
                    "UPDATE users SET email=?, modules=?, is_active=?, password_hash=COALESCE(?,password_hash) WHERE id=?",
                    (email, modules_json, int(payload.is_active), password_hash, user_id),
                )
            if parent and user_id != parent_id and payload.can_manage_users is not None:
                before = conn.execute("SELECT 1 FROM account_managers WHERE user_id=?", (user_id,)).fetchone() is not None
                if payload.can_manage_users:
                    conn.execute("INSERT OR IGNORE INTO account_managers(user_id) VALUES (?)", (user_id,))
                else:
                    conn.execute("DELETE FROM account_managers WHERE user_id=?", (user_id,))
                if before != payload.can_manage_users:
                    action = "выдано" if payload.can_manage_users else "отозвано"
                    conn.execute("INSERT INTO account_notifications(message) VALUES (?)", (f"Учётная запись «{username}»: {action} право «Управление пользователями».",))
    except sqlite3.IntegrityError:
        raise HTTPException(409, "Учётная запись с такими данными уже существует")
