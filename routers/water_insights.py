"""Authenticated polling API; only POST starts a bounded model job."""
import json
from functools import lru_cache
from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse
from core.roles import check_module_access
from services.auth.security import get_user_from_token
from services.water_dashboard.config import SOURCES, SNAPSHOT_FILE
from services.water_dashboard.insights import InsightService

router=APIRouter(prefix='/water-dashboard/insights',tags=['water-dashboard'])


def require_water_user(request:Request):
    # Re-read the authenticated account, rather than trusting cached state grants.
    user=get_user_from_token(request.cookies.get('access_token'))
    if not user or not user.get('is_active',True): raise HTTPException(401,'Необходимо войти в систему')
    if not check_module_access(user,'water-dashboard'): raise HTTPException(403,'Нет доступа к сводному дашборду')
    return user


@lru_cache(maxsize=1)
def service():
    return InsightService()


def _source(source_id):
    if source_id not in {s['id'] for s in SOURCES}: raise HTTPException(404,'Источник не найден')
    try:
        snapshot=json.loads(SNAPSHOT_FILE.read_text(encoding='utf-8'))
        return snapshot.get('sources',{}).get(source_id,{})
    except (OSError,ValueError,AttributeError): return {}


def _response(source_id,start):
    result=service().request(source_id,_source(source_id),start=start)
    return JSONResponse(result,headers={'Cache-Control':'no-store'})


@router.get('/{source_id}')
def status(source_id:str, user=Depends(require_water_user)):
    return _response(source_id,False)


@router.post('/{source_id}')
def start(source_id:str, request:Request, user=Depends(require_water_user)):
    origin=request.headers.get('origin')
    if request.headers.get('sec-fetch-site')=='cross-site' or (origin and origin.rstrip('/')!=str(request.base_url).rstrip('/')):
        raise HTTPException(403,'Недопустимый источник запроса')
    return _response(source_id,True)
