"""Shared collective complaints workspace, protected by the existing module grant."""
import asyncio
import hashlib
import json
from urllib.parse import quote

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse, Response
from core.web import templates
from core.roles import check_module_access
from services.auth.accounts import is_account_manager
from services.collective.client import HEADERS, period
from services.collective import presentation, store, refresh


def require_collective(request: Request):
    if not check_module_access(getattr(request.state, 'user', None), 'collective'):
        raise HTTPException(403, 'Нет доступа к коллективным жалобам')


def require_manager(request):
    if not is_account_manager(getattr(request.state, 'user', None)):
        raise HTTPException(403, 'Управление ответственными и перенос данных доступны администратору.')


def actor(request):
    user = getattr(request.state, 'user', None) or {}
    if user.get('kc_sub') is not None:
        subject = user['kc_sub']
        if isinstance(subject, str) and 0 < len(subject) <= 200 and not any(ord(c) < 32 for c in subject):
            return 'keycloak:' + subject
        raise HTTPException(401, 'Войдите в систему ещё раз.')
    if type(user.get('id')) is int and user['id'] > 0:
        return 'local:' + str(user['id'])
    raise HTTPException(401, 'Войдите в систему ещё раз.')


router = APIRouter(prefix='/mingkh/collective', dependencies=[Depends(require_collective)])
PRIVATE = {'Cache-Control': 'private, no-store'}
MAX_EXPORT_BYTES = 20 * 1024 * 1024


def result(data, status=200):
    return JSONResponse(data, status_code=status, headers=PRIVATE)


async def body_json(request, limit=100_000):
    payload = bytearray()
    async for chunk in request.stream():
        payload.extend(chunk)
        if len(payload) > limit:
            raise HTTPException(413, 'Слишком большой объём данных.')
    try:
        body = json.loads(payload, parse_constant=lambda _: (_ for _ in ()).throw(ValueError()))
        if not isinstance(body, dict):
            raise ValueError
        return body
    except (ValueError, UnicodeError, RecursionError):
        raise HTTPException(400, 'Некорректный формат данных.') from None


def failure(exc):
    if isinstance(exc, store.ConflictError):
        return result({'error': 'Карточку уже изменили. Проверьте сохранённые данные и повторите правку.', 'workflow': exc.current}, 409)
    if isinstance(exc, store.NotFoundError):
        return result({'error': 'Запись не найдена. Обновите данные страницы.'}, 404)
    if isinstance(exc, refresh.BusyError):
        return result({'error': str(exc)}, 409)
    return result({'error': str(exc)}, 400)


@router.get('')
async def page(request: Request):
    user = getattr(request.state, 'user', None) or {}
    owner = str(user.get('id') or user.get('kc_sub') or user.get('username') or '')
    return templates.TemplateResponse(request, 'collective.html', {
        'request': request, 'owner_key': hashlib.sha256(owner.encode()).hexdigest()[:24],
        'can_view_mingkh': check_module_access(user, 'mingkh'),
        'can_manage': is_account_manager(user),
    }, headers=PRIVATE)


@router.get('/api/data')
async def dataset(start: str = '', end: str = ''):
    try:
        period(start, end)
        saved = await asyncio.to_thread(store.get_snapshot, start, end)
        if saved is None:
            return result({'error': 'За этот период ещё нет сохранённых данных. Обновите из источника.'}, 404)
        return result(saved)
    except ValueError as exc:
        return failure(exc)


@router.post('/api/refresh')
async def refresh_data(request: Request):
    actor(request)
    body = await body_json(request, 1000)
    try:
        if set(body) != {'start', 'end'}:
            raise ValueError('Укажите начало и конец периода.')
        ticket = await asyncio.to_thread(refresh.submit, body['start'], body['end'], request.headers.get('Idempotency-Key'))
        return result(ticket, 202)
    except ValueError as exc:
        return failure(exc)


@router.get('/api/jobs/{ident}')
async def refresh_status(ident: str):
    job = await asyncio.to_thread(refresh.status, ident)
    return result(job) if job else result({'error': 'Обновление не найдено. Запустите его повторно.'}, 404)


@router.patch('/api/appeals/{key}')
async def edit_appeal(key: str, request: Request):
    who = actor(request)
    body = await body_json(request)
    try:
        version = body.pop('version', None)
        if not body or set(body) - {'report', 'assignee_id', 'theme'}:
            raise ValueError('Переданы неизвестные поля карточки.')
        workflow = await asyncio.to_thread(store.update_workflow, key, version, body, who)
        return result({'workflow': workflow})
    except (ValueError, store.NotFoundError, store.ConflictError) as exc:
        return failure(exc)


@router.get('/api/assignees')
async def assignees():
    return result({'assignees': await asyncio.to_thread(store.list_assignees)})


@router.post('/api/assignees')
async def add_assignee(request: Request):
    require_manager(request)
    who = actor(request)
    body = await body_json(request, 2000)
    try:
        if set(body) != {'name'}:
            raise ValueError('Укажите имя ответственного.')
        item = await asyncio.to_thread(store.create_assignee, body['name'], who)
        return result({'assignee': item, 'assignees': await asyncio.to_thread(store.list_assignees)}, 201)
    except (ValueError, store.ConflictError) as exc:
        return failure(exc)


@router.patch('/api/assignees/{ident}')
async def edit_assignee(ident: int, request: Request):
    require_manager(request)
    who = actor(request)
    body = await body_json(request, 2000)
    try:
        if set(body) != {'version', 'active'} or type(body['active']) is not bool:
            raise ValueError('Укажите состояние ответственного.')
        item = await asyncio.to_thread(store.update_assignee, ident, body['version'], {'archived': not body['active']}, who)
        return result({'assignee': item, 'assignees': await asyncio.to_thread(store.list_assignees)})
    except (ValueError, store.NotFoundError, store.ConflictError) as exc:
        return failure(exc)


@router.post('/api/import-distribution')
async def import_distribution(request: Request):
    require_manager(request)
    who = actor(request)
    body = await body_json(request, 20 * 1024 * 1024)
    try:
        counts = await asyncio.to_thread(store.import_legacy, body, who)
        return result({'imported': counts, 'assignees': await asyncio.to_thread(store.list_assignees)})
    except (ValueError, store.ConflictError) as exc:
        return failure(exc)


@router.post('/api/pptx')
async def export_pptx(request: Request):
    try:
        body = await body_json(request, MAX_EXPORT_BYTES)
        start, end = period(body.get('start'), body.get('end'))
        rows = body.get('rows')
        if not isinstance(rows, list) or not 1 <= len(rows) <= 2000:
            raise ValueError('Выберите от 1 до 2000 обращений для презентации.')
        if any(not isinstance(r, dict) or not isinstance(r.get('_key'), str) for r in rows):
            raise ValueError('Обновите страницу: презентация формируется из сохранённых обращений.')
        if rows:
            saved = await asyncio.to_thread(store.get_snapshot, start.isoformat(), end.isoformat())
            if not saved:
                raise ValueError('Сохранённые данные периода не найдены. Обновите страницу.')
            source = {key: dict(zip(HEADERS, row)) for key, row in zip(saved['keys'], saved['rows'])}
            resolved, seen = [], set()
            for row in rows:
                key = row.get('_key') if isinstance(row, dict) else None
                if not isinstance(key, str) or key not in source or key in seen:
                    raise ValueError('В презентации есть неизвестные или повторные обращения. Обновите страницу.')
                seen.add(key)
                workflow = saved['workflows'][key]
                resolved.append({**source[key], '_theme': workflow.get('theme') or row.get('_theme'),
                                 '_report': workflow.get('report', ''), '_assignee': workflow.get('assignee_name', '')})
            rows = resolved
        rows = presentation.validate_rows(rows)
        data = await asyncio.to_thread(presentation.build, rows, start, end)
    except (ValueError, TypeError) as exc:
        return result({'error': str(exc) or 'Некорректные данные презентации.'}, 400)
    name = f'Коллективные обращения {start:%d.%m}–{end:%d.%m.%Y}.pptx'
    return Response(data, media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    headers={**PRIVATE, 'Content-Disposition': "attachment; filename*=UTF-8''" + quote(name)})
