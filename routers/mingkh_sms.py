"""Private SMS workspace sharing the complaints dashboard's access grant."""
import asyncio
import json
import uuid

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse
from core.roles import check_module_access
from core.web import templates
from services.auth.accounts import is_account_manager
from services.mingkh.sms import context, jobs, operations, rez_data, sys_data


def require_access(request: Request):
    user = getattr(request.state, 'user', None)
    if not user:
        raise HTTPException(401, 'Войдите в Нейрону для продолжения.')
    if not check_module_access(user, 'mingkh'):
        raise HTTPException(403, 'Нет доступа к дашборду по жалобам МинЖКХ.')


router = APIRouter(prefix='/mingkh/sms', dependencies=[Depends(require_access)])


def reply(value, status=200):
    return JSONResponse(value, status_code=status, headers={'Cache-Control': 'private, no-store'})


async def body(request, limit=256 * 1024):
    data = bytearray()
    async for chunk in request.stream():
        data.extend(chunk)
        if len(data) > limit:
            raise HTTPException(413, 'Файл или запрос слишком большой.')
    return bytes(data)


async def json_body(request):
    try:
        value = json.loads(await body(request), parse_constant=lambda _: (_ for _ in ()).throw(ValueError()))
        if not isinstance(value, dict):
            raise ValueError
        return value
    except (ValueError, UnicodeDecodeError, RecursionError):
        raise HTTPException(422, 'Не удалось прочитать параметры запроса.') from None


@router.get('')
@router.get('/')
async def overview(request: Request):
    return await page(request, 'overview')


@router.get('/summary')
@router.get('/resonance')
@router.get('/system')
@router.get('/solutions')
async def section(request: Request):
    return await page(request, request.url.path.rsplit('/', 1)[-1])


async def page(request, name):
    return templates.TemplateResponse(request, 'mingkh-sms-' + name + '.html', {
        'request': request, 'sms_page': name, 'embed': request.query_params.get('embed') == '1',
        'can_manage_import': is_account_manager(request.state.user),
    }, headers={'Cache-Control': 'private, no-store'})


@router.get('/api/options')
async def options():
    return reply(await asyncio.to_thread(operations.options))


@router.get('/api/jobs/{ident}')
async def job_status(request: Request, ident: str):
    value = await asyncio.to_thread(jobs.status, jobs.owner_key(request.state.user), ident)
    if value is None:
        raise HTTPException(404, 'Расчёт не найден. Запустите его повторно.')
    return reply(value)


@router.get('/api/rez/found')
async def found():
    try:
        return reply(await asyncio.to_thread(rez_data.found_load))
    except ValueError as exc:
        return reply({'error': str(exc)}, 422)


@router.post('/api/rez/found')
async def save_found(request: Request):
    value = await json_body(request)
    try:
        start, end = context.date_range(value.get('from'), value.get('to'), max_days=366)
        result = await asyncio.to_thread(rez_data.found_upsert, start.isoformat(), end.isoformat(), value.get('cases'))
        return reply(result)
    except ValueError as exc:
        return reply({'error': str(exc)}, 422)


@router.get('/api/sys/state')
async def registry_state():
    dates = await asyncio.to_thread(sys_data.registries)
    return reply({'registries': [value.isoformat() for value in dates]})


@router.post('/api/import')
async def import_archive(request: Request):
    if not is_account_manager(request.state.user):
        raise HTTPException(403, 'Перенос реестров доступен только администратору.')
    from services.mingkh.sms.imports import import_archive as run_import
    payload = await body(request, 20 * 1024 * 1024)
    try:
        return reply(await asyncio.to_thread(run_import, payload))
    except ValueError as exc:
        return reply({'error': str(exc)}, 422)
    except (OSError, RuntimeError):
        return reply({'error': 'Не удалось сохранить архив. Администратору нужно проверить доступ к диску и сохранённые резервные копии.'}, 503)


@router.post('/api/{action:path}')
async def collect(request: Request, action: str):
    if action not in operations.HEAVY:
        raise HTTPException(404, 'Раздел не найден.')
    value = await json_body(request)
    try:
        payload = operations.validate(action, value)
        result = await asyncio.to_thread(jobs.submit, jobs.owner_key(request.state.user), action, payload,
                                        request.headers.get('Idempotency-Key') or uuid.uuid4().hex)
        return reply(result, 202)
    except jobs.JobConflict as exc:
        return reply({'error': str(exc)}, 409)
    except ValueError as exc:
        return reply({'error': str(exc)}, 422)
