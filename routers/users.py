from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field, StrictBool, StrictStr
from typing import Literal

from core.web import templates
from core.roles import MODULE_NAMES
from services.auth.accounts import require_account_manager, list_accounts, save_account
from services.auth.accounts import list_notifications, read_notifications
from services.auth.accounts import is_parent_manager
from services.auth.accounts import archive_account

router = APIRouter(dependencies=[Depends(require_account_manager)])


class AccountPayload(BaseModel):
    username: StrictStr = Field(min_length=3, max_length=32)
    email: StrictStr = Field(default="", max_length=254)
    password: StrictStr = Field(default="", max_length=72)
    modules: list[StrictStr] = Field(default_factory=list, max_length=50)
    is_active: StrictBool = True
    can_manage_users: StrictBool | None = None


class ReadNotificationsPayload(BaseModel):
    through_id: int = Field(ge=0)


class IntegrationPayload(BaseModel):
    username: StrictStr = Field(min_length=1,max_length=254)
    password: StrictStr = Field(default='',max_length=1024)


@router.get('/api/users/integrations/{service}')
async def integration_status(service: Literal['edds','edds_arm']):
    from services.auth.integrations import credential_status
    return credential_status(service)


@router.put('/api/users/integrations/{service}')
async def integration_save(payload: IntegrationPayload, service: Literal['edds','edds_arm']):
    from services.auth.integrations import save_credentials
    save_credentials(payload.username,payload.password,service)
    return {'ok':True}


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
    return templates.TemplateResponse(request, "users.html", {"request": request, "module_names": MODULE_NAMES})


@router.get("/api/users")
async def users_list(request: Request):
    return {"users": list_accounts(), "modules": MODULE_NAMES,
            "is_parent_manager": is_parent_manager(request.state.user)}


@router.post("/api/users", status_code=201)
async def users_create(payload: AccountPayload, request: Request):
    save_account(payload, actor=request.state.user)
    return {"ok": True}


@router.put("/api/users/{user_id}")
async def users_update(user_id: int, payload: AccountPayload, request: Request):
    save_account(payload, user_id, actor=request.state.user)
    return {"ok": True}
