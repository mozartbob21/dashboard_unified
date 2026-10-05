import asyncio
from datetime import date, timedelta
from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse
from core.web import templates
from core.roles import check_module_access
from core.map_tiles import browser_config
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


@router.get('/water-map')
async def water_map_page(request: Request):
    from services.auth.accounts import is_account_manager
    return templates.TemplateResponse(request, 'mingkh-water-map.html', {
        'request': request, 'map_config': browser_config(),
        'can_import': is_account_manager(getattr(request.state, 'user', None))})


@router.get('/api/water-map')
async def water_map_dataset(request: Request):
    import gzip
    from fastapi.responses import FileResponse, Response
    from services.mingkh import water_map
    await asyncio.to_thread(water_map.ensure_seed)
    if not water_map.DATA_FILE.exists():
        return JSONResponse({'error': 'Архив карты ещё не загружен. Администратор может загрузить ZIP или water_points.json на этой странице.'}, status_code=404)
    headers = {'Cache-Control': 'private, no-store', 'Vary': 'Accept-Encoding'}
    # The private file is stored and transferred compressed; browsers decompress JSON.
    encodings = request.headers.get('accept-encoding', '').lower()
    if any(item.strip().split(';')[0] == 'gzip' and 'q=0' not in item for item in encodings.split(',')):
        headers['Content-Encoding'] = 'gzip'
        return FileResponse(water_map.DATA_FILE, media_type='application/json', headers=headers)
    content = await asyncio.to_thread(lambda: gzip.decompress(water_map.DATA_FILE.read_bytes()))
    return Response(content, media_type='application/json', headers=headers)


def require_water_map_editor(request: Request):
    from services.auth.accounts import is_account_manager
    if not is_account_manager(getattr(request.state, 'user', None)):
        raise HTTPException(403, 'Загрузка и обновление карты доступны администратору.')


@router.post('/api/water-map/import', dependencies=[Depends(require_water_map_editor)])
async def water_map_import(request: Request):
    from services.mingkh import water_map
    body = bytearray()
    async for chunk in request.stream():
        body.extend(chunk)
        if len(body) > water_map.MAX_BYTES:
            return JSONResponse({'error': 'Файл больше 100 МБ.'}, status_code=413)
    try:
        result = await asyncio.to_thread(water_map.import_upload, body)
        return JSONResponse({'ok': True, **result}, headers={'Cache-Control': 'no-store'})
    except water_map.StoreBusy as exc:
        return JSONResponse({'error': str(exc)}, status_code=409)
    except ValueError as exc:
        return JSONResponse({'error': str(exc)}, status_code=400)


@router.post('/api/water-map/refresh', dependencies=[Depends(require_water_map_editor)])
async def water_map_refresh():
    from services.mingkh import water_map
    value = credentials('edds')
    if not value:
        raise HTTPException(503, 'Администратору нужно сохранить доступ «Добродел» в разделе «Пользователи → Логины и пароли».')
    try:
        return JSONResponse(water_map.start_refresh(value), status_code=202)
    except water_map.StoreBusy as exc:
        return JSONResponse({'error': str(exc)}, status_code=409)
    except ValueError as exc:
        return JSONResponse({'error': str(exc)}, status_code=400)


@router.get('/api/water-map/status')
async def water_map_status():
    from services.mingkh import water_map
    return JSONResponse(water_map.status(), headers={'Cache-Control': 'no-store'})
