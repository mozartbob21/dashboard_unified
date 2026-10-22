"""Authenticated local account linking. This is not a public bot webhook."""
from datetime import datetime, timezone
from functools import lru_cache
import hashlib
import hmac
import ipaddress
import json
import re

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.templating import Jinja2Templates

from services.max_bot.config import ROOT, Settings
from services.max_bot.store import Store, StoreError

router = APIRouter(prefix="/max-bot", tags=["max-bot"])
templates = Jinja2Templates(directory=str(ROOT / "templates"))


def settings():
    return Settings.from_env()


@lru_cache(maxsize=2)
def _store(config):
    try:
        return Store(config)
    except OSError:
        raise StoreError("SYSTEM_UNAVAILABLE", "Система временно недоступна. Повторите позже.", 503) from None


def get_store(config=Depends(settings)):
    if not config.enabled or not config.bot_token:
        raise HTTPException(503, {"code": "MAX_BOT_DISABLED", "message": "Бот пока не настроен."})
    try:
        return _store(config)
    except StoreError as exc:
        raise HTTPException(exc.status_code, {"code": exc.code, "message": exc.message}) from None


def require_user(request: Request):
    from services.auth.security import get_user_from_token
    user = getattr(request.state, "user", None) or get_user_from_token(request.cookies.get("access_token"))
    if not user or not user.get("is_active", True):
        raise HTTPException(401, {"code": "AUTH_REQUIRED"})
    if user.get("kc_sub") or type(user.get("id")) is not int:
        raise HTTPException(403, {"code": "LOCAL_ACCOUNT_REQUIRED", "message": "Для локальной проверки нужна локальная учётная запись Нейроны."})
    if "zips" not in user.get("modules", []):
        raise HTTPException(403, {"code": "MODULE_FORBIDDEN"})
    return user


def secure_transport(request: Request):
    if request.url.scheme == "https":
        return
    try:
        peer_local = ipaddress.ip_address(request.client.host).is_loopback
        host = request.url.hostname
        host_local = host == "localhost" or ipaddress.ip_address(host).is_loopback
    except (ValueError, AttributeError):
        peer_local = host_local = False
    if not (peer_local and host_local):
        raise HTTPException(403, {"code": "MAX_BOT_HTTPS_REQUIRED"})


def csrf_token(request):
    from services.auth.security import get_jwt_secret
    session = request.cookies.get("access_token", "")
    if not session:
        raise HTTPException(401, {"code": "AUTH_REQUIRED"})
    return hmac.new(get_jwt_secret().encode(), ("max-bot-csrf:" + session).encode(), hashlib.sha256).hexdigest()


async def mutation(request: Request, user=Depends(require_user), _=Depends(secure_transport)):
    expected_origin = f"{request.url.scheme}://{request.url.netloc}"
    header = request.headers.get("x-csrf-token", "")
    if request.headers.get("origin") != expected_origin or not re.fullmatch(r"[0-9a-f]{64}", header) or not hmac.compare_digest(header, csrf_token(request)):
        raise HTTPException(403, {"code": "CSRF_FAILED"})
    if request.headers.get("content-type", "").split(";")[0].strip() != "application/json":
        raise HTTPException(415, {"code": "JSON_REQUIRED"})
    data = bytearray()
    async for chunk in request.stream():
        data.extend(chunk)
        if len(data) > 1024:
            raise HTTPException(413, {"code": "BODY_TOO_LARGE"})
    try:
        def pairs(items):
            result = {}
            for key, value in items:
                if key in result:
                    raise ValueError("duplicate key")
                result[key] = value
            return result
        def constant(_):
            raise ValueError("nonfinite JSON")
        payload = json.loads(data, object_pairs_hook=pairs, parse_constant=constant)
    except (ValueError, UnicodeDecodeError):
        raise HTTPException(400, {"code": "INVALID_REQUEST"}) from None
    if not isinstance(payload, dict):
        raise HTTPException(400, {"code": "INVALID_REQUEST"})
    return payload


def call(function, *args):
    try:
        return function(*args)
    except StoreError as exc:
        raise HTTPException(exc.status_code, {"code": exc.code, "message": exc.message}) from None


def iso(timestamp):
    return datetime.fromtimestamp(timestamp, tz=timezone.utc).isoformat()


def response(data):
    return JSONResponse(data, headers={"Cache-Control": "private, no-store", "Referrer-Policy": "no-referrer"})


@router.get("", dependencies=[Depends(secure_transport)])
def page(request: Request, user=Depends(require_user)):
    return templates.TemplateResponse(request, "max_bot.html", {
        "request": request, "user": user, "user_username": user.get("username", ""),
        "user_role": user.get("role", ""), "max_bot_csrf": csrf_token(request),
    }, headers={"Cache-Control": "private, no-store", "Referrer-Policy": "no-referrer"})


@router.get("/api/status", dependencies=[Depends(secure_transport)])
def status(user=Depends(require_user), config=Depends(settings)):
    if not config.enabled or not config.bot_token:
        return response({"enabled": False, "active": False, "pending": []})
    state = call(call(_store, config).binding_status, user["id"])
    return response({"enabled": True, "active": state["active"],
        "pending": [{"id": item["request_id"], "label": "MAX " + item["max_user_mask"],
                     "expires_at": iso(item["expires_at"])} for item in state["pending"]],
        "active_label": "MAX " + state.get("binding", {}).get("max_user_mask", "") if state["active"] else "",
        "bot_public_name": config.bot_public_name})


@router.post("/api/link")
def link(payload=Depends(mutation), user=Depends(require_user), store=Depends(get_store)):
    if payload:
        raise HTTPException(400, {"code": "INVALID_REQUEST"})
    result = call(store.create_link, user["id"])
    store.audit("link_create", "ok")
    return response({"code": result["code"], "expires_at": iso(result["expires_at"])})


@router.post("/api/confirm")
def confirm(payload=Depends(mutation), user=Depends(require_user), store=Depends(get_store)):
    if set(payload) != {"request_id"} or not isinstance(payload["request_id"], str) or len(payload["request_id"]) != 32:
        raise HTTPException(400, {"code": "INVALID_REQUEST"})
    call(store.confirm_link, user["id"], payload["request_id"])
    store.audit("link_confirm", "ok")
    return response({"ok": True})


@router.post("/api/unlink")
def unlink(payload=Depends(mutation), user=Depends(require_user), store=Depends(get_store)):
    if payload:
        raise HTTPException(400, {"code": "INVALID_REQUEST"})
    call(store.unlink, user["id"])
    store.audit("unlink", "ok")
    return response({"ok": True})
