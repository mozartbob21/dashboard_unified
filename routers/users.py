from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse
from pydantic import BaseModel, ConfigDict, Field, StrictBool, StrictStr
from typing import Literal

from core.web import templates
from core.roles import MODULE_NAMES
from services.auth.accounts import require_account_manager, list_accounts, save_account, save_own_modules
from services.auth.accounts import list_notifications, read_notifications
from services.auth.accounts import is_parent_manager
from services.auth.accounts import archive_account
from services.auth.activity import account_activity, account_summaries, module_popularity

router = APIRouter(dependencies=[Depends(require_account_manager)])


class AccountPayload(BaseModel):
    username: StrictStr = Field(min_length=3, max_length=32)
    email: StrictStr = Field(default="", max_length=254)
    password: StrictStr = Field(default="", max_length=72)
    modules: list[StrictStr] = Field(default_factory=list, max_length=50)
    is_active: StrictBool = True
    can_manage_users: StrictBool | None = None


class OwnModulesPayload(BaseModel):
    model_config = ConfigDict(extra="forbid")
    modules: list[StrictStr] = Field(max_length=50)


class ReadNotificationsPayload(BaseModel):
    through_id: int = Field(ge=0)


async def _token_payload(request):
    import json
    raw = await request.body()
    if len(raw) > 16384:
        raise HTTPException(400, 'Слишком большой запрос токена')
    try:
        payload = json.loads(raw)
    except (ValueError, UnicodeError):
        raise HTTPException(400, 'Некорректный запрос токена') from None
    if not isinstance(payload, dict) or set(payload) != {'token'}:
        raise HTTPException(400, 'Укажите только поле token')
    return payload['token']


@router.get('/api/users/source-api-keys/{source}')
def source_key_status(source: str):
    from services.auth.source_api_keys import status
    return status(source)


@router.put('/api/users/source-api-keys/{source}')
async def source_key_save(source: str, request: Request):
    from services.auth.source_api_keys import save
    return save(source, await _token_payload(request))


@router.delete('/api/users/source-api-keys/{source}')
def source_key_remove(source: str):
    from services.auth.source_api_keys import remove
    return remove(source)


@router.get('/api/users/ai-tokens')
def ai_tokens_status():
    from services.auth.ai_credentials import statuses
    return statuses()


@router.put('/api/users/ai-tokens/{module_id}')
async def ai_token_save(module_id: str, request: Request):
    # Manual validation keeps submitted tokens out of FastAPI validation errors.
    from services.auth.ai_credentials import save
    return save(module_id, await _token_payload(request))


@router.delete('/api/users/ai-tokens/{module_id}')
def ai_token_remove(module_id: str):
    from services.auth.ai_credentials import remove
    return remove(module_id)


class IntegrationPayload(BaseModel):
    username: StrictStr = Field(min_length=1,max_length=254)
    password: StrictStr = Field(default='',max_length=1024)
    source_url: StrictStr | None = Field(default=None, max_length=2048)
    folder_path: StrictStr | None = Field(default=None, max_length=1024)


@router.get('/api/users/integrations/{service}')
async def integration_status(service: Literal['edds','edds_arm','mingkh','zip_cloud','presentations']):
    from services.auth.integrations import credential_status
    return credential_status(service)


@router.put('/api/users/integrations/{service}')
async def integration_save(payload: IntegrationPayload, service: Literal['edds','edds_arm','mingkh','zip_cloud','presentations']):
    from services.auth.integrations import save_credentials
    save_credentials(payload.username,payload.password,service,
                     source_url=payload.source_url,folder_path=payload.folder_path)
    return {'ok':True}


@router.post('/api/users/integrations/zip_cloud/check')
def integration_zip_cloud_check():
    from services.zip_curator.source import check_cloud_access
    from services.zip_curator.yandex_disk import DiskSyncError
    try:
        return check_cloud_access()
    except HTTPException:
        raise
    except DiskSyncError as exc:
        return JSONResponse({'ok': False, 'message': str(exc)}, status_code=502)
    except Exception:
        return JSONResponse({'ok': False, 'message': 'Не удалось проверить папку Облака МО. Проверьте подключение сервера и повторите попытку.'}, status_code=502)


@router.post('/api/users/integrations/presentations/check')
def integration_presentations_check():
    """Read-only source check: no files/messages/baseline changes."""
    from services.max_bot.presentation_config import source_factory
    from services.max_bot.presentation_source import SourceError
    try:
        with source_factory() as source:
            items = source.list_attachments()
        return {'ok': True, 'message': f'Доступ подтверждён. Файлов в источнике: {len(items)}. Проверка ничего не отправляет в MAX.'}
    except SourceError as exc:
        return JSONResponse({'ok': False, 'message': exc.message}, status_code=502)
    except Exception:
        return JSONResponse({'ok': False, 'message': 'Не удалось проверить источник презентаций. Повторите позже.'}, status_code=502)


@router.post('/api/users/integrations/edds/check')
def integration_edds_check():
    """Проверяет вход и одну страницу API отчёта без сохранения жалоб."""
    from services.auth.integrations import credentials
    from services.edds.dobrodel import DobrodelClient, DobrodelError

    logged_in = False
    try:
        try:
            account = credentials('edds')
        except HTTPException:
            raise DobrodelError('config') from None
        if not isinstance(account, dict) or not all(
            isinstance(account.get(key), str) and account[key] for key in ('username', 'password')
        ):
            raise DobrodelError('credentials')
        with DobrodelClient(account['username'], account['password']) as client:
            import time
            client.deadline = time.monotonic() + 120
            client.login()
            logged_in = True
            from datetime import date, timedelta
            from services.edds.collector import CURATOR, STATUSES
            today = date.today()
            client.probe_report({'filters.curators': CURATOR, 'filters.statuses': STATUSES,
                'filters.createdAfter': (today - timedelta(days=1)).isoformat(),
                'filters.createdBefore': today.isoformat()})
    except DobrodelError as error:
        status = 400 if error.code in {'credentials', 'login', 'access'} else 502
        message = ('Вход подтверждён, но проверка выгрузки не прошла. ' if logged_in else '') + str(error)
        return JSONResponse(status_code=status, content={'ok': False, 'message': message})
    return {'ok': True, 'message': 'Вход и получение данных отчёта Добродела подтверждены. Можно обновить свод жалоб.'}


@router.post('/api/users/{user_id}/archive')
async def users_archive(user_id: int, request: Request):
    archive_account(user_id, request.state.user)
    return {'ok': True}


@router.post('/api/users/{user_id}/restore')
async def users_restore(user_id: int, request: Request):
    archive_account(user_id, request.state.user, restore=True)
    return {'ok': True}


@router.get("/api/users/notifications")
async def notifications_list():
    return list_notifications()


@router.post("/api/users/notifications/read")
async def notifications_read(payload: ReadNotificationsPayload):
    read_notifications(payload.through_id)
    return {"ok": True}


@router.get("/users", response_class=HTMLResponse)
async def users_page(request: Request):
    return templates.TemplateResponse(request, "users.html", {"request": request, "module_names": MODULE_NAMES,
                                     "initial_admin_section": "users"}, headers={"Cache-Control": "no-store"})


@router.get("/api/users")
async def users_list(request: Request):
    summaries = account_summaries()
    users = [{**user, "activity": summaries.get(user["id"])} for user in list_accounts()]
    return {"users": users, "modules": MODULE_NAMES,
            "is_parent_manager": is_parent_manager(request.state.user),
            "current_user_id": next((user["id"] for user in users
                                     if user["username"] == request.state.user.get("username")), None)}


@router.get("/api/users/module-popularity")
def users_module_popularity():
    return JSONResponse(module_popularity(), headers={"Cache-Control": "no-store"})


@router.get("/api/users/{user_id}/activity")
async def users_activity(user_id: int):
    return account_activity(user_id)


@router.post("/api/users", status_code=201)
async def users_create(payload: AccountPayload, request: Request):
    save_account(payload, actor=request.state.user)
    return {"ok": True}


@router.put("/api/users/me/modules")
async def users_own_modules(payload: OwnModulesPayload, actor=Depends(require_account_manager)):
    save_own_modules(payload.modules, actor=actor)
    return {"ok": True}


@router.put("/api/users/{user_id}")
async def users_update(user_id: int, payload: AccountPayload, request: Request):
    save_account(payload, user_id, actor=request.state.user)
    return {"ok": True}
