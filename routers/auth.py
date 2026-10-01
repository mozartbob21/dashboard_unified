# -*- coding: utf-8 -*-
"""routers/auth.py — вход/выход, регистрация, Keycloak, настройки профиля."""
import os
import time

import bcrypt
from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse

from pydantic import BaseModel, Field
from services.auth import registration, home_preferences
from services.auth.security import (
    authenticate_user,
    create_access_token,
    get_user_from_token,
)
from core.web import templates

router = APIRouter()

# ── Keycloak OIDC (активируется через AUTH_PROVIDER=keycloak в .env) ──
KC_SERVER_URL = os.getenv("KC_SERVER_URL", "http://localhost:8080")
KC_REALM = os.getenv("KC_REALM", "neurona")
KC_CLIENT_ID = os.getenv("KC_CLIENT_ID", "neurona-web")
KC_CLIENT_SECRET = os.getenv("KC_CLIENT_SECRET", "")

LOGIN_ATTEMPTS: dict = {}


def rate_limit_ok(key: str, limit: int = 5, window: int = 300) -> bool:
    """Не более limit попыток за window секунд для одного key."""
    if os.getenv("PYTEST_CURRENT_TEST"):
        return True
    now = time.time()
    rec = LOGIN_ATTEMPTS.setdefault(key, [])
    rec[:] = [t for t in rec if now - t < window]
    if len(rec) >= limit:
        return False
    rec.append(now)
    return True


# ── Keycloak: роуты логина (OAuth2 Authorization Code Flow) ──
@router.get("/login/keycloak")
async def login_keycloak_redirect(request: Request):
    redirect_uri = str(request.url_for("login_keycloak_callback"))
    auth_url = (
        f"{KC_SERVER_URL}/realms/{KC_REALM}/protocol/openid-connect/auth"
        f"?client_id={KC_CLIENT_ID}&response_type=code"
        f"&redirect_uri={redirect_uri}&scope=openid+profile+email"
    )
    return RedirectResponse(auth_url, status_code=302)


@router.get("/login/keycloak/callback", name="login_keycloak_callback")
async def login_keycloak_callback(request: Request, code: str = ""):
    if not code:
        return RedirectResponse("/login?error=no_code", status_code=302)
    try:
        import httpx
        redirect_uri = str(request.url_for("login_keycloak_callback"))
        token_url = f"{KC_SERVER_URL}/realms/{KC_REALM}/protocol/openid-connect/token"
        async with httpx.AsyncClient(timeout=15.0) as client:
            resp = await client.post(token_url, data={
                "grant_type": "authorization_code",
                "code": code,
                "redirect_uri": redirect_uri,
                "client_id": KC_CLIENT_ID,
                "client_secret": KC_CLIENT_SECRET,
            })
        if resp.status_code != 200:
            print(f"[keycloak] token exchange failed: {resp.status_code} {resp.text[:200]}")
            return RedirectResponse("/login?error=token_exchange", status_code=302)
        tokens = resp.json()
        access_token = tokens.get("access_token", "")
        refresh_token = tokens.get("refresh_token", "")
        expires_in = tokens.get("expires_in", 3600)
        response = RedirectResponse("/", status_code=302)
        response.set_cookie("access_token", access_token, httponly=True, max_age=expires_in, samesite="lax")
        if refresh_token:
            response.set_cookie("refresh_token", refresh_token, httponly=True, max_age=7 * 86400, samesite="lax")
        response.set_cookie("auth_provider", "keycloak", max_age=expires_in, samesite="lax")
        return response
    except Exception as e:
        print(f"[keycloak] login error: {e}")
        return RedirectResponse("/login?error=exception", status_code=302)


# ── Локальный вход/выход ──
@router.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, error: str = "", message: str = ""):
    token = request.cookies.get("access_token")
    if get_user_from_token(token):
        return RedirectResponse(url="/", status_code=302)
    return templates.TemplateResponse(
        request,
        "login.html",
        {"request": request, "error": error, "message": message},
    )


@router.post("/login")
async def login_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
):
    client_ip = request.client.host if request.client else "unknown"
    rl_key = f"{client_ip}:{username.lower()}"

    if not rate_limit_ok(rl_key):
        return RedirectResponse(
            url="/login?error=Слишком много попыток входа. Подождите 5 минут.",
            status_code=303,
        )

    user = authenticate_user(username, password)
    if not user:
        return RedirectResponse(url="/login?error=Неверный логин или пароль", status_code=303)

    access_token = create_access_token({
        "sub": user["username"],
        "role": user.get("role", ""),
    })

    from services.auth.accounts import is_account_manager
    response = RedirectResponse(url="/users" if is_account_manager(user) else "/", status_code=303)
    response.set_cookie(
        key="access_token",
        value=access_token,
        httponly=True,
        max_age=60 * 60 * 8,
        samesite="strict",
        secure=False,
    )
    return response


@router.get("/logout")
async def logout():
    response = RedirectResponse(url="/login?message=Вы вышли из системы", status_code=303)
    response.delete_cookie("access_token")
    response.delete_cookie("auth_provider")
    return response


# ── Регистрация ──
@router.get("/register", response_class=HTMLResponse)
async def register_page(request: Request):
    return templates.TemplateResponse(request, "register.html", {"request": request})


@router.post("/api/register")
async def api_register(payload: dict):
    username = payload.get("username") or ""
    email = payload.get("email") or ""
    password = payload.get("password") or ""
    if not all(isinstance(value, str) for value in (username, email, password)):
        return JSONResponse(status_code=400, content={
            "ok": False, "message": "Логин, почта и пароль должны быть строками",
        })
    if len(password) < 6:
        return JSONResponse(status_code=400, content={
            "ok": False, "message": "Пароль должен быть не короче 6 символов",
        })
    if len(password.encode("utf-8")) > 72:
        return JSONResponse(status_code=400, content={
            "ok": False, "message": "Пароль слишком длинный: максимум 72 байта в UTF-8",
        })

    password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    ok, result = registration.register_user(username, email, password_hash)
    if not ok:
        return JSONResponse(status_code=400, content={"ok": False, "message": result})

    access_token = create_access_token({"sub": result["username"], "role": result["role"]})
    response = JSONResponse(content={
        "ok": True,
        "message": "Аккаунт создан! Доступ к блокам назначит управляющий учётными записями.",
        "redirect": "/",
    })
    response.set_cookie(key="access_token", value=access_token, httponly=True,
                        max_age=60 * 60 * 8, samesite="strict", secure=False)
    return response


# ── Настройки профиля ──
@router.get("/api/me/settings")
async def api_me_settings(request: Request):
    return {"ok": True, "settings": registration.get_settings(request.state.user["id"])}


@router.post("/api/me/settings")
async def api_me_save_settings(request: Request, payload: dict):
    settings = registration.save_settings(request.state.user["id"], payload.get("settings") or {})
    return {"ok": True, "settings": settings}


# Избранное главной страницы: только текущая учётная запись, без client-supplied user_id.


class HomeFavoritesPayload(BaseModel):
    favorites: list[str] = Field(max_length=20)


@router.get('/api/me/home-favorites')
def api_home_favorites(request: Request):
    return JSONResponse(home_preferences.preferences(getattr(request.state, 'user', None)),
                        headers={'Cache-Control': 'no-store'})


@router.put('/api/me/home-favorites')
def api_save_home_favorites(request: Request, payload: HomeFavoritesPayload):
    return JSONResponse(home_preferences.preferences(getattr(request.state, 'user', None), payload.favorites),
                        headers={'Cache-Control': 'no-store'})
