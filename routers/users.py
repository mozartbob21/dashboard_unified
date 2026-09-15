from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse
from pydantic import BaseModel, Field, StrictBool, StrictStr

from core.web import templates
from core.roles import MODULE_NAMES
from services.auth.accounts import require_account_manager, list_accounts, save_account
from services.auth.accounts import list_notifications, read_notifications

router = APIRouter(dependencies=[Depends(require_account_manager)])


class AccountPayload(BaseModel):
    username: StrictStr = Field(min_length=3, max_length=32)
    email: StrictStr = Field(default="", max_length=254)
    password: StrictStr = Field(default="", max_length=72)
    modules: list[StrictStr] = Field(default_factory=list, max_length=50)
    is_active: StrictBool = True


class ReadNotificationsPayload(BaseModel):
    through_id: int = Field(ge=0)


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
async def users_list():
    return {"users": list_accounts(), "modules": MODULE_NAMES}


@router.post("/api/users", status_code=201)
async def users_create(payload: AccountPayload):
    save_account(payload)
    return {"ok": True}


@router.put("/api/users/{user_id}")
async def users_update(user_id: int, payload: AccountPayload):
    save_account(payload, user_id)
    return {"ok": True}
