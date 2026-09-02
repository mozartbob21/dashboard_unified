# -*- coding: utf-8 -*-
"""routers/auth.py — вход/выход, регистрация, Keycloak, настройки профиля."""
import os
import time

import bcrypt
from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse

from services.auth import mailer, registration
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

    response = RedirectResponse(url="/", status_code=303)
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
async def api_register_start(payload: dict):
    try:
        from core.web import templates as _t  # noqa: F401 (не используется, просто для совместимости)
        username = (payload.get("username") or "").strip()
        email = (payload.get("email") or "").strip()
        password = str(payload.get("password") or "")

        if len(password) < 6:
            return JSONResponse(status_code=400,
                                content={"ok": False, "message": "Пароль должен быть не короче 6 символов"})

        password_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        print(f"[register] Старт регистрации: username={username}, email={email}")

        ok, result = registration.start_registration(username, email, password_hash)
        if not ok:
            print(f"[register] Ошибка валидации: {result}")
            return JSONResponse(status_code=400, content={"ok": False, "message": result})

        print(f"[register] Заявка создана, отправляю код на {result['email']}")
        sent, err = mailer.send_verification_code(result["email"], result["code"])
        if not sent:
            print(f"[register] Ошибка отправки письма: {err}")
            if os.getenv("REGISTRATION_DEBUG_CODE") == "1":
                return {"ok": True, "debug_code": result["code"],
                        "message": "SMTP не настроен (dev-режим). Код: " + result["code"]}
            return JSONResponse(status_code=500, content={
                "ok": False,
                "message": f"Не удалось отправить письмо. Проверьте YANDEX_SMTP_USER / YANDEX_SMTP_PASSWORD. ({err})",
            })

        print("[register] Письмо отправлено успешно")
        return {"ok": True, "message": f"Код отправлен на {result['email']}"}
    except Exception as e:
        import traceback
        print("[register] НЕПРЕДВИДЕННАЯ ОШИБКА:")
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"ok": False, "message": f"Ошибка сервера: {str(e)}"})


@router.post("/api/register/resend")
async def api_register_resend(payload: dict):
    ident = (payload.get("username") or payload.get("email") or "").strip()
    ok, result = registration.resend_code(ident)
    if not ok:
        return JSONResponse(status_code=400, content={"ok": False, "message": result})

    sent, err = mailer.send_verification_code(result["email"], result["code"])
    if not sent:
        if os.getenv("REGISTRATION_DEBUG_CODE") == "1":
            return {"ok": True, "debug_code": result["code"], "message": "Dev-режим. Код: " + result["code"]}
        return JSONResponse(status_code=500, content={"ok": False, "message": f"Ошибка отправки: {err}"})

    return {"ok": True, "message": f"Письмо отправлено на {result['email']}"}


@router.post("/api/register/verify")
async def api_register_verify(payload: dict):
    ident = (payload.get("username") or payload.get("email") or "").strip()
    code = (payload.get("code") or "").strip()

    ok, result = registration.verify_registration(ident, code)
    if not ok:
        return JSONResponse(status_code=400, content={"ok": False, "message": result})

    access_token = create_access_token({"sub": result["username"], "role": result["role"]})
    response = JSONResponse(content={"ok": True, "message": "Аккаунт создан! Добро пожаловать.", "redirect": "/"})
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
