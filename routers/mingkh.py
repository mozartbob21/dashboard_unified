import asyncio
from datetime import date, timedelta
from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse
from core.web import templates
from core.roles import check_module_access
from services.auth.integrations import credentials
from services.mingkh import dashboard
from services.mingkh.client import Pentaho, PortalError


def require_mingkh(request: Request):
    if not check_module_access(getattr(request.state,'user',None),'mingkh'):
        raise HTTPException(403,'Нет доступа к дашборду МИНЖКХ')

router=APIRouter(prefix='/mingkh',dependencies=[Depends(require_mingkh)])


def access():
    value=credentials('mingkh')
    if not value:
        raise HTTPException(503,'Администратору нужно сохранить логин и пароль МИНЖКХ в разделе «Пользователи → Логины и пароли».')
    return value


@router.get('')
async def page(request:Request):
    today=date.today();start=(today.weekday()-3)%7
    return templates.TemplateResponse(request,'mingkh.html',{'request':request,'presets':dashboard.PRESETS,
        'initial_dates':[(today-timedelta(days=d)).isoformat() for d in (start,0,start+7,7)]})


@router.get('/api/dataset')
async def dataset(request:Request):
    try:
        result=await asyncio.to_thread(dashboard.get_dataset,access(),dict(request.query_params))
        return JSONResponse(result,headers={'Cache-Control':'no-store'})
    except ValueError as exc:return JSONResponse({'error':str(exc)},status_code=400)
    except PortalError as exc:return JSONResponse({'error':str(exc)},status_code=502)
    except HTTPException as exc:return JSONResponse({'error':exc.detail},status_code=exc.status_code)


@router.get('/api/updated')
async def updated():
    try:
        value,_=await asyncio.to_thread(Pentaho(**access()).last_updated)
        return JSONResponse({'updated':value},headers={'Cache-Control':'no-store'})
    except PortalError as exc:return JSONResponse({'error':str(exc)},status_code=502)
    except HTTPException as exc:return JSONResponse({'error':exc.detail},status_code=exc.status_code)
