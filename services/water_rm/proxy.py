# -*- coding: utf-8 -*-
"""
Прокси для модуля «Проверка задач по качеству воды».
- /water-rm/rmapi/*  — проксирует запросы на Redmine (mgkh.rm.mosreg.ru)
- /water-rm/yadisk   — скачивает публичный файл с Яндекс.Диска
Ключ Redmine берётся из переменной окружения MGKH_API_KEY (.env).
"""
import os
import urllib.request
import urllib.error
import urllib.parse
import json
import pathlib

DATA_DIR = pathlib.Path(__file__).resolve().parents[2] / "data" / "water_rm"
DATA_DIR.mkdir(parents=True, exist_ok=True)
HISTORY_FILE = DATA_DIR / "last_check.json"

from fastapi import APIRouter, Request, HTTPException
from fastapi.responses import Response

router = APIRouter(prefix="/water-rm", tags=["water_rm"])

RM_BASE = "https://mgkh.rm.mosreg.ru"

# Пытаемся подгрузить .env, если доступен python-dotenv (без падения, если его нет)
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


def _get_api_key() -> str:
    return (os.getenv("MGKH_API_KEY") or "").strip()


async def _proxy_redmine(request: Request, subpath: str):
    """Универсальный проксировщик к Redmine."""
    query = request.url.query
    url = f"{RM_BASE}/{subpath}"
    if query:
        url += f"?{query}"

    body = await request.body()
    req = urllib.request.Request(
        url,
        data=body if body else None,
        method=request.method,
    )

    # Ключ: приоритет — реальный заголовок от фронта, иначе из .env.
    # Заглушка "__ENV__" означает «использовать ключ из окружения».
    hdr_key = request.headers.get("X-Redmine-API-Key", "").strip()
    if hdr_key and hdr_key != "__ENV__":
        api_key = hdr_key
    else:
        api_key = _get_api_key()
    if api_key:
        req.add_header("X-Redmine-API-Key", api_key)

    # Пробрасываем Content-Type
    ctype = request.headers.get("Content-Type")
    if ctype:
        req.add_header("Content-Type", ctype)

    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            content = r.read()
            status_code = r.status
            content_type = r.headers.get("Content-Type", "application/json")
    except urllib.error.HTTPError as e:
        content = e.read()
        status_code = e.code
        content_type = e.headers.get("Content-Type", "text/plain") if e.headers else "text/plain"
    except Exception as e:
        content = str(e).encode("utf-8")
        status_code = 502
        content_type = "text/plain; charset=utf-8"

    return Response(content=content, status_code=status_code, media_type=content_type)


@router.api_route("/rmapi/{subpath:path}", methods=["GET", "POST", "PUT", "DELETE"])
async def rm_proxy(request: Request, subpath: str):
    return await _proxy_redmine(request, subpath)


@router.get("/has-key")
async def has_key():
    """Сообщает фронту, задан ли ключ в .env (сам ключ НЕ отдаём)."""
    return {"has_key": bool(_get_api_key())}

@router.get("/history")
async def get_history():
    """Отдаёт результат последней проверки (или пусто)."""
    if HISTORY_FILE.exists():
        return Response(
            content=HISTORY_FILE.read_bytes(),
            status_code=200,
            media_type="application/json; charset=utf-8",
        )
    return Response(content=b'{"empty": true}', status_code=200,
                    media_type="application/json; charset=utf-8")


@router.post("/history")
async def save_history(request: Request):
    """Перезаписывает результат последней проверки."""
    body = await request.body()
    HISTORY_FILE.write_bytes(body)
    return {"ok": True}

@router.get("/yadisk")
async def yadisk_download(public_key: str = ""):
    """Скачивает публичный файл с Яндекс.Диска."""
    if not public_key:
        raise HTTPException(status_code=400, detail="Не передана ссылка на таблицу (public_key)")

    try:
        api_url = (
            "https://cloud-api.yandex.net/v1/disk/public/resources/download"
            "?public_key=" + urllib.parse.quote(public_key, safe="")
        )
        with urllib.request.urlopen(api_url, timeout=60) as r:
            href = json.loads(r.read().decode("utf-8"))["href"]

        with urllib.request.urlopen(href, timeout=120) as r:
            content = r.read()

        return Response(
            content=content,
            status_code=200,
            media_type="application/octet-stream",
        )
    except urllib.error.HTTPError as e:
        msg = f"Яндекс.Диск ответил HTTP {e.code} — проверьте, что ссылка публичная."
        return Response(content=msg.encode("utf-8"), status_code=502, media_type="text/plain; charset=utf-8")
    except Exception as e:
        return Response(content=str(e).encode("utf-8"), status_code=502, media_type="text/plain; charset=utf-8")

