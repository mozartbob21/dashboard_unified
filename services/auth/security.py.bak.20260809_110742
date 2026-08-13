from fastapi import Request
"""
Локальный модуль безопасности.
- bcrypt напрямую для хеширования паролей
- python-jose для подписи JWT-токенов
Все работает локально, без внешних сервисов.
"""
import os
import json
import secrets
from pathlib import Path
from datetime import datetime, timedelta, timezone
from typing import Any

import bcrypt
from jose import jwt, JWTError


BASE_DIR = Path(__file__).resolve().parents[2]
AUTH_DATA_DIR = BASE_DIR / "data" / "auth"
USERS_FILE = AUTH_DATA_DIR / "users.json"
SECRET_FILE = AUTH_DATA_DIR / "jwt_secret.key"

AUTH_DATA_DIR.mkdir(parents=True, exist_ok=True)


JWT_ALGORITHM = "HS256"
JWT_EXPIRE_HOURS = int(os.getenv("JWT_EXPIRE_HOURS", "8"))

# bcrypt принимает максимум 72 байта в пароле
BCRYPT_MAX_BYTES = 72


def _generate_secret_key() -> str:
    return secrets.token_urlsafe(64)


def get_jwt_secret() -> str:
    """
    Получает локальный секретный ключ для JWT.
    При первом запуске генерирует и сохраняет в файл.
    """
    env_secret = os.getenv("JWT_SECRET_KEY", "").strip()
    if env_secret:
        return env_secret

    if SECRET_FILE.exists():
        secret = SECRET_FILE.read_text(encoding="utf-8").strip()
        if secret:
            return secret

    new_secret = _generate_secret_key()
    SECRET_FILE.write_text(new_secret, encoding="utf-8")
    print("[auth] Сгенерирован новый JWT_SECRET_KEY:", str(SECRET_FILE))
    return new_secret


def _truncate_password(password: str) -> bytes:
    """
    bcrypt не принимает пароли длиннее 72 байт.
    Безопасно обрезаем до этого лимита.
    """
    encoded = password.encode("utf-8")
    if len(encoded) > BCRYPT_MAX_BYTES:
        encoded = encoded[:BCRYPT_MAX_BYTES]
    return encoded


def hash_password(password: str) -> str:
    """Хеширует пароль через bcrypt."""
    pwd_bytes = _truncate_password(password)
    salt = bcrypt.gensalt(rounds=12)
    hashed = bcrypt.hashpw(pwd_bytes, salt)
    return hashed.decode("utf-8")


def verify_password(plain_password: str, hashed_password: str) -> bool:
    """Проверяет пароль против bcrypt-хеша."""
    if not plain_password or not hashed_password:
        return False
    try:
        pwd_bytes = _truncate_password(plain_password)
        hash_bytes = hashed_password.encode("utf-8")
        return bcrypt.checkpw(pwd_bytes, hash_bytes)
    except Exception as e:
        print("[auth] Ошибка проверки пароля:", repr(e))
        return False


def create_access_token(payload: dict[str, Any], expires_hours: int | None = None) -> str:
    to_encode = dict(payload)
    expire = datetime.now(timezone.utc) + timedelta(
        hours=expires_hours if expires_hours is not None else JWT_EXPIRE_HOURS
    )
    to_encode.update({"exp": expire, "iat": datetime.now(timezone.utc)})
    return jwt.encode(to_encode, get_jwt_secret(), algorithm=JWT_ALGORITHM)


def decode_access_token(token: str) -> dict[str, Any] | None:
    if not token:
        return None
    try:
        return jwt.decode(token, get_jwt_secret(), algorithms=[JWT_ALGORITHM])
    except JWTError as e:
        print("[auth] JWT отклонен:", repr(e))
        return None




# ===========================
# ОБФУСКАЦИЯ ДЕФОЛТНЫХ ПАРОЛЕЙ
# ===========================
# Пароли в коде хранятся в закодированном виде (XOR + base64),
# чтобы не светить их в репозитории. Это НЕ настоящее шифрование
# (ключ-соль лежит рядом), а защита от случайного просмотра.
# Реальная защита паролей - bcrypt-хеши в data/auth/users.json.

import base64 as _b64

_OBFUSCATION_SALT = "unified-dashboard-2024"

def _decode_password(encoded: str) -> str:
    """Декодирует пароль, зашитый в коде."""
    try:
        salt_bytes = _OBFUSCATION_SALT.encode("utf-8")
        data = _b64.b64decode(encoded.encode("ascii"))
        result = bytes(b ^ salt_bytes[i % len(salt_bytes)] for i, b in enumerate(data))
        return result.decode("utf-8")
    except Exception:
        return ""


# ===========================
# ЛОКАЛЬНОЕ ХРАНИЛИЩЕ ПОЛЬЗОВАТЕЛЕЙ (SQLite)
# ===========================
# Пользователи хранятся в таблице users в data/dashboard.db.
# Старый файл data/auth/users.json больше не используется
# (оставлен на диске только для архивных целей).

from utils.db import get_db_connection


DEFAULT_USERS = [
    {
        "username": "admin",
        "password": _decode_password("FAoEDwdUVh4="),
        "role": "Администратор",
        "modules": ["edo", "overdue", "watercontrol", "utnkr", "cameras", "appeals", "municipality-report","cds","mgkh_rm"],
    },
    {
        "username": "data",
        "password": _decode_password("EQ8dB1hXVw=="),
        "role": "Контроль данных",
        "modules": ["edo", "overdue", "watercontrol"],
    },
    {
        "username": "utnkr",
        "password": _decode_password("ABoHDRtUVh4="),
        "role": "УТНКР",
        "modules": ["utnkr"],
    },
    {
        "username": "cameras",
        "password": _decode_password("Fg8EAxsEFxxWUg=="),
        "role": "Камеры",
        "modules": ["cameras"],
    },
    {
        "username": "appeals",
        "password": _decode_password("FB4ZAwgJFxxWUg=="),
        "role": "Эмпатичные ответы",
        "modules": ["appeals"],
    },
]


def _row_to_user(row) -> dict[str, Any]:
    """
    Превращает sqlite3.Row в словарь в том же формате,
    в котором раньше пользователи лежали в users.json.
    """
    modules_raw = row["modules"] or "[]"
    try:
        modules = json.loads(modules_raw)
    except Exception:
        modules = []

    return {
        "username": row["username"],
        "password_hash": row["password_hash"],
        "role": row["role"],
        "modules": modules,
        "is_active": bool(row["is_active"]),
    }


def _ensure_default_users() -> None:
    """
    Если таблица users пуста — заполняет ее дефолтными пользователями.
    Вызывается при первом обращении.
    """
    with get_db_connection() as conn:
        count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if count > 0:
            return

        for user in DEFAULT_USERS:
            conn.execute(
                """
                INSERT INTO users (username, password_hash, role, modules, is_active)
                VALUES (?, ?, ?, ?, 1)
                """,
                (
                    user["username"],
                    hash_password(user["password"]),
                    user["role"],
                    json.dumps(user["modules"], ensure_ascii=False),
                ),
            )

        print("[auth] Таблица users была пустой - созданы дефолтные пользователи:")
        for user in DEFAULT_USERS:
            print(f"  - {user['username']}  ({user['role']})")


def load_users() -> list[dict[str, Any]]:
    """Загружает всех пользователей из БД."""
    _ensure_default_users()
    try:
        with get_db_connection() as conn:
            rows = conn.execute(
                "SELECT username, password_hash, role, modules, is_active FROM users"
            ).fetchall()
            return [_row_to_user(row) for row in rows]
    except Exception as e:
        print("[auth] Ошибка чтения users из БД:", repr(e))
        return []


def save_users(users: list[dict[str, Any]]) -> None:
    """
    Сохраняет список пользователей в БД (UPSERT по username).
    Используется существующим кодом, которому проще передать весь список.
    """
    with get_db_connection() as conn:
        for user in users:
            conn.execute(
                """
                INSERT INTO users (username, password_hash, role, modules, is_active)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(username) DO UPDATE SET
                    password_hash = excluded.password_hash,
                    role          = excluded.role,
                    modules       = excluded.modules,
                    is_active     = excluded.is_active
                """,
                (
                    user.get("username", "").strip().lower(),
                    user.get("password_hash", ""),
                    user.get("role", ""),
                    json.dumps(user.get("modules", []), ensure_ascii=False),
                    1 if user.get("is_active", True) else 0,
                ),
            )


def find_user_by_username(username: str) -> dict[str, Any] | None:
    """Ищет одного пользователя по username (без учета регистра)."""
    if not username:
        return None
    _ensure_default_users()
    username = username.strip().lower()
    try:
        with get_db_connection() as conn:
            row = conn.execute(
                """
                SELECT username, password_hash, role, modules, is_active
                FROM users
                WHERE LOWER(username) = ?
                """,
                (username,),
            ).fetchone()
            return _row_to_user(row) if row else None
    except Exception as e:
        print("[auth] Ошибка поиска пользователя в БД:", repr(e))
        return None


def authenticate_user(username: str, password: str) -> dict[str, Any] | None:
    user = find_user_by_username(username)
    if not user:
        return None
    if not user.get("is_active", True):
        return None
    if not verify_password(password, user.get("password_hash", "")):
        return None
    return user


def get_user_from_token(token: str | None) -> dict[str, Any] | None:
    if not token:
        return None
    payload = decode_access_token(token)
    if not payload:
        return None
    username = payload.get("sub")
    if not username:
        return None
    user = find_user_by_username(username)
    if not user or not user.get("is_active", True):
        return None
    return user


def has_module_access(user: dict[str, Any] | None, module_id: str) -> bool:
    if not user:
        return False
    return module_id in user.get("modules", [])


ADMIN_ROLES = {"admin", "администратор"}

def require_admin_user(request: Request) -> dict:
    """Возвращает админа из cookie access_token или кидает 403."""
    from fastapi import HTTPException
    token = request.cookies.get("access_token")
    user = get_user_from_token(token) or {}
    if (user.get("role") or "").strip().lower() not in ADMIN_ROLES:
        raise HTTPException(status_code=403, detail="Доступ только для администраторов")
    return user

