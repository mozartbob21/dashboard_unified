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
# ЛОКАЛЬНОЕ ХРАНИЛИЩЕ ПОЛЬЗОВАТЕЛЕЙ
# ===========================

DEFAULT_USERS = [
    {
        "username": "admin",
        "password": _decode_password("FAoEDwdUVh4="),
        "role": "Администратор",
        "modules": ["edo", "overdue", "watercontrol", "utnkr", "cameras", "appeals", "municipality-report"],
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


def _create_default_users_file() -> None:
    users = []
    for user in DEFAULT_USERS:
        users.append({
            "username": user["username"],
            "password_hash": hash_password(user["password"]),
            "role": user["role"],
            "modules": user["modules"],
            "is_active": True,
        })

    USERS_FILE.write_text(
        json.dumps(users, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print("[auth] Создан users.json с дефолтными пользователями:", str(USERS_FILE))
    print("[auth] Логины пользователей по умолчанию:")
    for user in DEFAULT_USERS:
        print(f"  - {user['username']}  ({user['role']})")


def load_users() -> list[dict[str, Any]]:
    if not USERS_FILE.exists():
        _create_default_users_file()
    try:
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        print("[auth] Ошибка чтения users.json:", repr(e))
        return []


def save_users(users: list[dict[str, Any]]) -> None:
    USERS_FILE.write_text(
        json.dumps(users, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def find_user_by_username(username: str) -> dict[str, Any] | None:
    if not username:
        return None
    username = username.strip().lower()
    for user in load_users():
        if user.get("username", "").lower() == username:
            return user
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
