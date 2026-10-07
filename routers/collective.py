"""Collective dashboard uses the app session and the administrator's MINGKH integration."""
import asyncio
import hashlib
import json
from datetime import datetime
from urllib.parse import quote

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse, Response
from core.web import templates
from core.roles import check_module_access
from services.auth.integrations import credentials
from services.collective.client import CollectivePortal, HEADERS, PortalError, period
from services.collective import presentation

def require_collective(request: Request):
    if not check_module_access(getattr(request.state, 'user', None), 'collective'):
        raise HTTPException(403, 'Нет доступа к коллективным жалобам')


router = APIRouter(prefix='/mingkh/collective', dependencies=[Depends(require_collective)])
PRIVATE = {'Cache-Control': 'private, no-store'}
MAX_EXPORT_BYTES = 20 * 1024 * 1024


@router.get('')
async def page(request: Request):
    user = getattr(request.state, 'user', None) or {}
    owner = str(user.get('id') or user.get('username') or '')
    return templates.TemplateResponse(request, 'collective.html', {
        'request': request, 'owner_key': hashlib.sha256(owner.encode()).hexdigest()[:24],
        'can_view_mingkh': check_module_access(user, 'mingkh'),
    }, headers=PRIVATE)


@router.get('/api/data')
async def dataset(start: str = '', end: str = ''):
    try:
        start_date, end_date = period(start, end)
        access = credentials('mingkh')
        if not access:
            raise HTTPException(503, 'Администратору нужно сохранить логин и пароль МИНЖКХ в разделе «Пользователи → Логины и пароли».')
        kurators, rows = await asyncio.to_thread(CollectivePortal(**access).fetch, start_date, end_date)
        return JSONResponse({
            'head': HEADERS, 'rows': rows,
            'meta': f'МинЖКХ МО и объединения ({len(kurators)}) · {start_date:%d.%m.%Y}–{end_date:%d.%m.%Y}',
            'updated': datetime.now().strftime('%d.%m.%Y %H:%M'),
            'start': start, 'end': end,
        }, headers=PRIVATE)
    except ValueError as exc:
        return JSONResponse({'error': str(exc)}, status_code=400, headers=PRIVATE)
    except PortalError as exc:
        return JSONResponse({'error': str(exc)}, status_code=502, headers=PRIVATE)
    except HTTPException as exc:
        return JSONResponse({'error': exc.detail}, status_code=exc.status_code, headers=PRIVATE)


@router.post('/api/pptx')
async def export_pptx(request: Request):
    payload = bytearray()
    async for chunk in request.stream():
        payload.extend(chunk)
        if len(payload) > MAX_EXPORT_BYTES:
            return JSONResponse({'error': 'Слишком много данных для презентации. Уточните фильтры.'}, status_code=413, headers=PRIVATE)
    try:
        body = json.loads(payload)
        if not isinstance(body, dict):
            raise ValueError('Некорректные данные презентации.')
        start, end = period(body.get('start'), body.get('end'))
        rows = presentation.validate_rows(body.get('rows'))
        data = await asyncio.to_thread(presentation.build, rows, start, end)
    except (ValueError, TypeError) as exc:
        return JSONResponse({'error': str(exc) or 'Некорректные данные презентации.'}, status_code=400, headers=PRIVATE)
    name = f'Коллективные обращения {start:%d.%m}–{end:%d.%m.%Y}.pptx'
    return Response(data, media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    headers={**PRIVATE, 'Content-Disposition': "attachment; filename*=UTF-8''" + quote(name)})
