"""Offline treatment-plant complaints snapshot in the ministry dashboard."""
import asyncio

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse

from core.map_tiles import browser_config
from core.roles import check_module_access
from core.web import templates
from routers.mingkh import require_mingkh
from services.auth.accounts import is_account_manager
from services.mingkh import treatment
from services import run_history


router = APIRouter(prefix='/mingkh/treatment', dependencies=[Depends(require_mingkh)])
PRIVATE = {'Cache-Control': 'private, no-store'}


@router.get('')
async def page(request: Request):
    user = request.state.user
    return templates.TemplateResponse(request, 'mingkh-treatment.html', {
        'can_import': is_account_manager(user), 'can_edds': check_module_access(user, 'edds'),
        'map_config': browser_config(),
    }, headers=PRIVATE)


@router.get('/api/data')
async def dataset(request: Request):
    data = await asyncio.to_thread(treatment.overview, check_module_access(request.state.user, 'edds'))
    if data is None:
        return JSONResponse({'error': 'Снимок ещё не загружен.'}, status_code=404, headers=PRIVATE)
    return JSONResponse(data, headers=PRIVATE)


@router.get('/api/complaints/{ident}')
async def complaint(ident: int):
    row = await asyncio.to_thread(treatment.detail, 'complaints', str(ident))
    if row is None:
        raise HTTPException(404, 'Жалоба не найдена в текущем снимке.')
    return JSONResponse(row, headers=PRIVATE)


@router.get('/api/events/{ident}')
async def event(ident: str, request: Request):
    if not check_module_access(request.state.user, 'edds'):
        raise HTTPException(403, 'Нет доступа к событиям ЕДДС.')
    row = await asyncio.to_thread(treatment.detail, 'edds', ident)
    if row is None:
        raise HTTPException(404, 'Событие не найдено в текущем снимке.')
    return JSONResponse(row, headers=PRIVATE)


@router.post('/api/import')
async def import_snapshot(request: Request):
    if not is_account_manager(request.state.user):
        raise HTTPException(403, 'Загружать снимок может только управляющий.')
    run_id = await asyncio.to_thread(run_history.record_start, 'mingkh', params={'operation': 'treatment'})
    try:
        body = bytearray()
        async for chunk in request.stream():
            if len(body) + len(chunk) > treatment.MAX_UPLOAD:
                raise HTTPException(413, 'Архив больше 24 МБ.')
            body.extend(chunk)
        result = await asyncio.to_thread(treatment.import_archive, body)
    except (HTTPException, treatment.SnapshotError) as exc:
        await asyncio.to_thread(run_history.record_finish, run_id, 'error', 'Снимок не прошёл проверку.')
        if isinstance(exc, HTTPException):
            raise
        raise HTTPException(400, str(exc)) from None
    except Exception:
        await asyncio.to_thread(run_history.record_finish, run_id, 'error', 'Не удалось сохранить снимок.')
        raise
    await asyncio.to_thread(run_history.record_finish, run_id, 'success', metrics={
        'appeals': result['complaints'], 'facilities': result['objects'],
    })
    result.pop('incidents', None)
    return JSONResponse(result, headers=PRIVATE)
