# -*- coding: utf-8 -*-
"""Роли и права: общие для app.py и роутеров."""
from fastapi import HTTPException, Request
from services.auth.security import (
    get_user_from_token,
    has_module_access,
    ADMIN_ROLES,
)

# Роли, которые видят ВСЕ модули и админ-интерфейс.
FULL_ACCESS_ROLES = {"администратор", "руководитель", "пользователь"}

# Полный список модулей платформы
ALL_MODULE_IDS = [
    "edo", "overdue", "watercontrol", "utnkr", "cameras", "appeals",
    "cds", "mgkh_rm", "ecur", "municipality-report", "water-dashboard",
    "water_rm", "tools",
]

# Маппинг ролей Keycloak -> модули платформы
KC_ROLE_TO_MODULE = {
    "admin":             "__full_access__",
    "администратор":     "__full_access__",
    "руководитель":      "__full_access__",
    "edo":               "edo",
    "overdue":           "overdue",
    "watercontrol":      "watercontrol",
    "utnkr":             "utnkr",
    "cameras":           "cameras",
    "appeals":           "appeals",
    "cds":               "cds",
    "mgkh_rm":           "mgkh_rm",
    "ecur":              "ecur",
    "municipality-report": "municipality-report",
    "water-dashboard":   "water-dashboard",
    "water_rm":          "water_rm",
    "tools":             "tools",
}

def _kc_user_modules(user: dict) -> list:
    """Собирает список модулей по ролям Keycloak."""
    roles = user.get("roles", []) or []
    role = (user.get("role") or "").strip().lower()
    if role in {"admin", "администратор", "руководитель"}:
        return list(ALL_MODULE_IDS)
    modules = []
    for r in roles:
        mapped = KC_ROLE_TO_MODULE.get(r.lower())
        if mapped and mapped != "__full_access__" and mapped not in modules:
            modules.append(mapped)
    return modules

def is_full_access(user) -> bool:
    role = (user.get("role") or "").strip().lower()
    if role in FULL_ACCESS_ROLES:
        return True
    if user.get("kc_sub") and role in {"admin", "администратор", "руководитель"}:
        return True
    return False

def check_module_access(user, module_id) -> bool:
    """Полные роли проходят на любой модуль, остальные — стандартная проверка."""
    if is_full_access(user):
        return True
    return has_module_access(user, module_id)

def effective_modules(user) -> list:
    """Модули для главной. Поддерживает локальных и Keycloak-пользователей."""
    if is_full_access(user):
        return list(ALL_MODULE_IDS)
    if user.get("kc_sub"):
        return _kc_user_modules(user)
    return user.get("modules", [])

async def require_admin_or_full(request: Request):
    """Dependency для роутеров, доступных админу и полным ролям."""
    user = getattr(request.state, "user", None)
    if not user:
        user = get_user_from_token(request.cookies.get("access_token"))
    if not user:
        raise HTTPException(status_code=401, detail="Требуется авторизация")
    role = (user.get("role") or "").strip().lower()
    if role in ADMIN_ROLES or role in FULL_ACCESS_ROLES:
        return user
    raise HTTPException(status_code=403, detail="Недостаточно прав")
