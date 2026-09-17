"""Authenticated UI/API. Owner identity always comes from Neurona, never the request body."""
import hashlib
import os
import re
from typing import Annotated
from urllib.parse import quote, urlsplit
from uuid import UUID

from fastapi import APIRouter, Depends, HTTPException, Query, Request, Response
from pydantic import BaseModel, ConfigDict, Field, field_validator

from core.roles import check_module_access
from core.web import templates
from services.telegram import client as tg
from services.telegram.store import Busy
from utils.db import get_db_connection


def require_user(request: Request):
    user = getattr(request.state, 'user', None)
    if not user:
        raise HTTPException(401, 'Войдите в «Нейрону».')
    if not check_module_access(user, 'telegram'):
        raise HTTPException(403, 'Нет доступа к модулю Telegram.')
    if user.get('kc_sub'):
        identity = f"keycloak:{os.getenv('KC_SERVER_URL', 'http://localhost:8080')}:{os.getenv('KC_REALM', 'neurona')}:{user['kc_sub']}"
    else:
        with get_db_connection() as db:
            row = db.execute('SELECT id,created_at FROM users WHERE username=? AND is_active=1', (user.get('username'),)).fetchone()
        if not row:
            raise HTTPException(401, 'Учётная запись недоступна.')
        identity = f"local:{row['id']}:{row['created_at']}"
    return hashlib.sha256(identity.encode()).hexdigest()


def guard(request: Request, response: Response):
    response.headers['Cache-Control'] = 'no-store'
    if request.method not in {'GET', 'HEAD'}:
        if request.headers.get('X-Neurona-Telegram') != '1':
            raise HTTPException(403, 'Обновите страницу перед выполнением действия.')
        origin = request.headers.get('origin')
        if origin and urlsplit(origin).netloc != request.headers.get('host'):
            raise HTTPException(403, 'Запрос с другого сайта запрещён.')
    if request.url.path.startswith('/telegram/api/') and not request.url.path.endswith('/status') and tg.settings():
        raise HTTPException(503, 'Администратору нужно настроить Telegram на сервере.')


router = APIRouter(prefix='/telegram', dependencies=[Depends(guard)])
Owner = Annotated[str, Depends(require_user)]


class Payload(BaseModel):
    model_config = ConfigDict(extra='forbid')


class Phone(Payload):
    phone: str = Field(min_length=7, max_length=25)

    @field_validator('phone')
    @classmethod
    def international_phone(cls, value):
        value = re.sub(r'[\s()-]', '', value)
        if not re.fullmatch(r'\+[1-9]\d{6,14}', value):
            raise ValueError('Введите телефон с кодом страны, например +79991234567')
        return value


class Confirmation(Payload):
    code: str | None = Field(default=None, max_length=12)
    password: str | None = Field(default=None, max_length=256)


class Selection(Payload):
    ids: list[str] = Field(max_length=100)


class Message(Payload):
    text: str = Field(min_length=1, max_length=4096)
    request_id: UUID
    reply_to: int | None = Field(default=None, gt=0)

    @field_validator('text')
    @classmethod
    def meaningful_text(cls, value):
        if not value.strip():
            raise ValueError('Введите сообщение')
        return value


async def call(operation):
    try:
        return await operation
    except tg.TelegramError as error:
        raise HTTPException(error.status, error.message,
                            headers={'Retry-After': str(error.retry_after)} if error.retry_after else None) from None
    except Busy:
        raise HTTPException(409, 'Предыдущий запрос ещё выполняется. Повторите через несколько секунд.') from None


@router.get('')
async def page(request: Request, owner: Owner):
    return templates.TemplateResponse(request, 'telegram.html', {'request': request}, headers={'Cache-Control': 'no-store'})


@router.get('/api/status')
async def status(owner: Owner):
    return tg.status(owner)


@router.post('/api/auth/start')
async def start(payload: Phone, owner: Owner):
    return await call(tg.start_login(owner, payload.phone))


@router.post('/api/auth/confirm')
async def confirm(payload: Confirmation, owner: Owner):
    return await call(tg.finish_login(owner, code=payload.code, password=payload.password))


@router.post('/api/disconnect')
async def disconnect(owner: Owner):
    return await call(tg.disconnect(owner))


@router.get('/api/dialogs')
async def dialogs(owner: Owner, more: bool = False):
    return await call(tg.dialogs(owner, more))


@router.put('/api/chats')
async def selection(payload: Selection, owner: Owner):
    try:
        return tg.select_chats(owner, payload.ids)
    except tg.TelegramError as error:
        raise HTTPException(error.status, error.message) from None
    except Busy:
        raise HTTPException(409, 'Дождитесь завершения текущего запроса.') from None


@router.get('/api/chats/{chat_id}/messages')
async def messages(chat_id: str, owner: Owner, before: int = Query(default=0, ge=0)):
    return await call(tg.messages(owner, chat_id, before))


@router.post('/api/chats/{chat_id}/messages')
async def send(chat_id: str, payload: Message, owner: Owner):
    return await call(tg.send(owner, chat_id, payload.text, str(payload.request_id), payload.reply_to))


@router.get('/api/chats/{chat_id}/attachments/{message_id}')
async def attachment(chat_id: str, message_id: int, owner: Owner):
    content, filename = await call(tg.attachment(owner, chat_id, message_id))
    filename = filename.replace('/', '_').replace('\\', '_').replace('\r', '').replace('\n', '')[:180]
    return Response(content, media_type='application/octet-stream', headers={
        'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff',
        'Content-Disposition': "attachment; filename*=UTF-8''" + quote(filename),
    })
