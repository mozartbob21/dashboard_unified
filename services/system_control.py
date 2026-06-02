"""
Управление системой: перезапуск приложения и очистка кешей.
"""

from pathlib import Path
from datetime import datetime
import os

from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse

router = APIRouter(prefix="/api/system", tags=["system"])

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RELOAD_TARGET = PROJECT_ROOT / "app.py"


def clear_server_caches():
    cleared = []
    try:
        from services import municipality_report
        if hasattr(municipality_report, "_MUNICIPALITY_REGISTRY_CACHE"):
            municipality_report._MUNICIPALITY_REGISTRY_CACHE = None
            cleared.append("municipality_registry")
    except Exception:
        pass
    return cleared


@router.get("/health")
async def health_check():
    return {
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
    }


@router.post("/clear-cache")
async def clear_cache():
    cleared = clear_server_caches()
    return {
        "status": "ok",
        "cleared_caches": cleared,
        "message": "Серверные кеши очищены",
    }


@router.post("/restart")
async def restart_app():
    if not RELOAD_TARGET.exists():
        raise HTTPException(
            status_code=500,
            detail=f"Не найден файл для перезапуска: {RELOAD_TARGET}",
        )

    cleared = clear_server_caches()

    try:
        now = datetime.now().timestamp()
        os.utime(str(RELOAD_TARGET), (now, now))
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Не удалось инициировать перезапуск: {str(e)}",
        )

    return JSONResponse(
        content={
            "status": "ok",
            "cleared_caches": cleared,
            "reload_target": str(RELOAD_TARGET.relative_to(PROJECT_ROOT)),
            "message": (
                "Запрос на перезапуск отправлен. "
                "Сервер перезапустится через 1-2 секунды. "
                "Если сервер запущен без --reload, перезапуск не сработает."
            ),
        }
    )
