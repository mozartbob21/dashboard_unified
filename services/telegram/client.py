"""Bounded MTProto operations; every method is scoped to a Neurona user."""
import asyncio
import hashlib
import importlib.util
import io
import os
import re
import time
import uuid
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from functools import lru_cache
from urllib.parse import urlsplit, unquote

from .store import Store, Busy

MAX_MEDIA = 20 * 1024 * 1024


class TelegramError(Exception):
    def __init__(self, message, status=400, retry_after=None):
        self.message, self.status, self.retry_after = message, status, retry_after
        super().__init__(message)


def settings():
    missing = []
    if not os.getenv('TELEGRAM_API_ID', '').isdigit() or int(os.getenv('TELEGRAM_API_ID', '0') or 0) <= 0:
        missing.append('TELEGRAM_API_ID')
    if not re.fullmatch(r'[0-9a-fA-F]{32}', os.getenv('TELEGRAM_API_HASH', '').strip()):
        missing.append('TELEGRAM_API_HASH')
    if importlib.util.find_spec('telethon') is None:
        missing.append('библиотека Telethon')
    return missing


@lru_cache(maxsize=1)
def storage():
    return Store()


def proxy_settings():
    raw = os.getenv('TELEGRAM_PROXY_URL', '').strip()
    if not raw:
        return None
    try:
        parsed = urlsplit(raw)
        valid = parsed.scheme in {'socks5', 'socks4', 'http'} and parsed.hostname and parsed.port
    except ValueError:
        valid = False
    if not valid:
        raise TelegramError('Проверьте TELEGRAM_PROXY_URL на сервере.', 503)
    return {'proxy_type': parsed.scheme, 'addr': parsed.hostname, 'port': parsed.port,
            'username': unquote(parsed.username) if parsed.username else None,
            'password': unquote(parsed.password) if parsed.password else None, 'rdns': True}


def _peer_data(entity):
    from telethon import utils
    peer = utils.get_input_peer(entity)
    return peer.to_dict()


def _peer(data):
    from telethon.tl import types
    constructors = {'InputPeerUser': types.InputPeerUser, 'InputPeerChannel': types.InputPeerChannel,
                    'InputPeerChat': types.InputPeerChat, 'InputPeerSelf': types.InputPeerSelf}
    return constructors[data['_']](**{k: v for k, v in data.items() if k != '_'})


def _public_chat(chat):
    return {key: chat[key] for key in ('id', 'title', 'kind')}


def _selected(account, chat_id):
    if str(chat_id) not in account.get('selected', []):
        raise TelegramError('Этот чат не выбран для работы в «Нейроне».', 403)
    chat = account.get('dialogs', {}).get(str(chat_id))
    if not chat:
        raise TelegramError('Обновите список чатов.', 404)
    return chat


def _client(session):
    from telethon import TelegramClient
    from telethon.sessions import StringSession
    return TelegramClient(StringSession(session or ''), int(os.environ['TELEGRAM_API_ID']),
                          os.environ['TELEGRAM_API_HASH'], proxy=proxy_settings(),
                          connection_retries=1, request_retries=0, timeout=12,
                          flood_sleep_threshold=0, receive_updates=False,
                          device_model='Neurona Web', app_version='3.0', lang_code='ru')


def _map_error(error):
    from telethon import errors
    if isinstance(error, errors.FloodWaitError):
        return TelegramError(f'Telegram просит подождать {error.seconds} сек.', 429, error.seconds)
    if isinstance(error, (errors.PhoneCodeInvalidError, errors.PhoneCodeEmptyError)):
        return TelegramError('Неверный код. Проверьте код из Telegram.')
    if isinstance(error, errors.PhoneCodeExpiredError):
        return TelegramError('Код истёк. Начните подключение заново.')
    if isinstance(error, errors.PasswordHashInvalidError):
        return TelegramError('Неверный пароль двухэтапной защиты.')
    if isinstance(error, errors.PhoneNumberInvalidError):
        return TelegramError('Проверьте номер телефона и код страны.')
    if isinstance(error, errors.UnauthorizedError):
        return TelegramError('Telegram-сессия завершена. Отключите её и подключитесь заново.', 401)
    if isinstance(error, errors.ForbiddenError):
        return TelegramError('Telegram не разрешает это действие в выбранном чате.', 403)
    if isinstance(error, (OSError, TimeoutError, ConnectionError)):
        return TelegramError('Сервер не смог связаться с Telegram. Проверьте соединение или прокси.', 502)
    return TelegramError('Telegram не выполнил запрос. Попробуйте позже или обновите список чатов.', 502)


@asynccontextmanager
async def connection(owner, authorized=True):
    if settings():
        raise TelegramError('Telegram не настроен на сервере.', 503)
    store = storage()
    try:
        with store.claim(owner):
            account = store.get(owner)
            if authorized and account.get('stage') != 'ready':
                raise TelegramError('Подключите аккаунт Telegram.', 409)
            client = _client(account.get('session'))
            try:
                async with asyncio.timeout(45):
                    await client.connect()
                    if authorized and not await client.is_user_authorized():
                        account['stage'] = 'expired'
                        store.put(owner, account)
                        raise TelegramError('Сессия Telegram отозвана. Подключите аккаунт заново.', 401)
                    yield client, account, store
            except TelegramError:
                raise
            except Exception as error:
                raise _map_error(error) from None
            finally:
                try:
                    await asyncio.wait_for(client.disconnect(), 5)
                except (OSError, TimeoutError):
                    pass
    except Busy:
        raise TelegramError('Предыдущий запрос ещё выполняется. Повторите через несколько секунд.', 409) from None


def status(owner):
    missing = settings()
    if missing:
        return {'configured': False, 'missing': missing, 'stage': 'disconnected', 'chats': []}
    account = storage().get(owner)
    stage = account.get('stage', 'disconnected')
    if stage in {'code', 'password'} and account.get('expires', 0) < time.time():
        stage = 'disconnected'
    return {'configured': True, 'stage': stage, 'profile': account.get('profile') if stage == 'ready' else None,
            'chats': [_public_chat(account['dialogs'][key]) for key in account.get('selected', [])
                      if key in account.get('dialogs', {})] if stage == 'ready' else []}


async def start_login(owner, phone):
    async with connection(owner, authorized=False) as (client, account, store):
        if account.get('stage') == 'ready':
            raise TelegramError('Сначала отключите текущий Telegram-аккаунт.', 409)
        if time.time() - account.get('code_sent', 0) < 60:
            raise TelegramError('Повторный код можно запросить через минуту.', 429, 60)
        sent = await client.send_code_request(phone)
        store.put(owner, {'session': client.session.save(), 'phone': phone, 'phone_code_hash': sent.phone_code_hash,
                          'stage': 'code', 'code_sent': time.time(), 'expires': time.time() + 600, 'attempts': 0})
        return {'stage': 'code'}


async def finish_login(owner, code=None, password=None):
    from telethon import errors
    async with connection(owner, authorized=False) as (client, account, store):
        if account.get('stage') not in {'code', 'password'} or account.get('expires', 0) < time.time():
            raise TelegramError('Начните подключение заново: время ввода кода истекло.')
        if account.get('attempts', 0) >= 8:
            raise TelegramError('Слишком много попыток. Запросите новый код.', 429)
        account['attempts'] = account.get('attempts', 0) + 1
        store.put(owner, account)
        try:
            if account['stage'] == 'password':
                if not password:
                    raise TelegramError('Введите пароль двухэтапной защиты.')
                await client.sign_in(password=password)
            else:
                if not code:
                    raise TelegramError('Введите код из Telegram.')
                await client.sign_in(phone=account['phone'], code=code, phone_code_hash=account['phone_code_hash'])
        except errors.SessionPasswordNeededError:
            account.update(stage='password', session=client.session.save())
            store.put(owner, account)
            return {'stage': 'password'}
        me = await client.get_me()
        from telethon.utils import get_display_name
        store.put(owner, {'session': client.session.save(), 'stage': 'ready', 'generation': uuid.uuid4().hex,
                          'profile': {'name': get_display_name(me), 'username': me.username or ''},
                          'dialogs': {}, 'selected': [], 'sent': {}})
        return {'stage': 'ready'}


async def disconnect(owner):
    # An already revoked/unfinished authorization can be removed locally.
    store = storage()
    account = store.get(owner)
    if account.get('stage') in {'ready', 'expired'}:
        async with connection(owner, authorized=False) as (client, account, store):
            if await client.is_user_authorized():
                await client.log_out()
            store.delete(owner)
    else:
        with store.claim(owner):
            store.delete(owner)
    return {'ok': True}


async def dialogs(owner, more=False):
    async with connection(owner) as (client, account, store):
        cursor = account.get('cursor') if more else None
        options = {}
        if cursor:
            options = {'offset_date': datetime.fromtimestamp(cursor['date'], timezone.utc), 'ignore_pinned': True,
                       'offset_id': cursor['message_id'], 'offset_peer': _peer(cursor['peer'])}
        batch = await client.get_dialogs(limit=50, **options)
        known = account.setdefault('dialogs', {})
        page = []
        for dialog in batch:
            key = str(dialog.id)
            chat = {'id': key, 'title': dialog.name or key,
                    'kind': 'channel' if dialog.is_channel and not dialog.is_group else 'group' if dialog.is_group else 'user',
                    'peer': _peer_data(dialog.entity)}
            known[key] = chat
            page.append(_public_chat(chat))
        last = batch[-1] if batch else None
        cursor = {'date': last.date.timestamp(), 'message_id': last.message.id,
                  'peer': _peer_data(last.entity)} if last and last.date and last.message else None
        account['cursor'] = cursor
        account['session'] = client.session.save()
        store.put(owner, account)
        return {'items': page, 'has_more': bool(cursor and len(batch) == 50), 'selected': account.get('selected', [])}


def select_chats(owner, ids):
    store = storage()
    with store.claim(owner):
        account = store.get(owner)
        if account.get('stage') != 'ready':
            raise TelegramError('Подключите Telegram.', 409)
        ids = list(dict.fromkeys(ids))
        if any(key not in account.get('dialogs', {}) for key in ids):
            raise TelegramError('Выбирайте чаты из списка вашего Telegram-аккаунта.', 403)
        account['selected'] = ids
        store.put(owner, account)
    return status(owner)


async def messages(owner, chat_id, before=0):
    from telethon.utils import get_display_name
    async with connection(owner) as (client, account, store):
        chat = _selected(account, chat_id)
        batch = await client.get_messages(_peer(chat['peer']), limit=50, offset_id=before)
        items = []
        for message in reversed(batch):
            sender = message.sender
            file = message.file
            items.append({'id': message.id, 'text': message.raw_text or '',
                          'out': bool(message.out), 'date': message.date.isoformat(),
                          'sender': get_display_name(sender) if sender else 'Участник',
                          'reply_to': message.reply_to_msg_id,
                          'file': {'name': file.name or ('Фото.jpg' if message.photo else 'Вложение'),
                                   'size': file.size or 0, 'downloadable': (file.size or 0) <= MAX_MEDIA}
                                  if file else None,
                          'service': bool(message.action),
                          'unsupported_media': bool(message.media and not file)})
        return {'items': items, 'has_more': len(batch) == 50}


async def send(owner, chat_id, text, request_id, reply_to=None):
    from telethon.errors import RandomIdDuplicateError
    from telethon.tl.functions.messages import SendMessageRequest
    from telethon.tl.types import InputReplyToMessage
    async with connection(owner) as (client, account, store):
        chat = _selected(account, chat_id)
        fingerprint = hashlib.sha256(f'{chat_id}\0{text}\0{reply_to}'.encode()).hexdigest()
        sent = account.setdefault('sent', {})
        receipt = sent.get(request_id)
        if receipt:
            if receipt['fingerprint'] != fingerprint:
                raise TelegramError('Этот идентификатор уже использован для другого сообщения.', 409)
            if receipt.get('done'):
                return {'ok': True, 'already_sent': True}
        random_id = int.from_bytes(hashlib.sha256(f"{owner}:{account['generation']}:{request_id}".encode()).digest()[:8], 'big', signed=True)
        sent[request_id] = {'fingerprint': fingerprint, 'done': False, 'time': time.time()}
        store.put(owner, account)
        duplicate = False
        try:
            await client(SendMessageRequest(peer=_peer(chat['peer']), message=text, random_id=random_id,
                                           no_webpage=True, reply_to=InputReplyToMessage(reply_to) if reply_to else None))
        except RandomIdDuplicateError:
            # Telegram received the earlier attempt even if its response was lost.
            duplicate = True
        sent[request_id]['done'] = True
        account['sent'] = {key: value for key, value in sent.items() if value['time'] > time.time() - 7 * 86400}
        store.put(owner, account)
        return {'ok': True, 'already_sent': duplicate}


class LimitedBuffer(io.BytesIO):
    def write(self, value):
        if self.tell() + len(value) > MAX_MEDIA:
            raise TelegramError('Вложение больше 20 МБ. Загрузка недоступна.', 413)
        return super().write(value)


async def attachment(owner, chat_id, message_id):
    async with connection(owner) as (client, account, store):
        chat = _selected(account, chat_id)
        message = await client.get_messages(_peer(chat['peer']), ids=message_id)
        if not message or not message.file:
            raise TelegramError('Вложение не найдено.', 404)
        if (message.file.size or 0) > MAX_MEDIA:
            raise TelegramError('Вложение больше 20 МБ. Загрузка недоступна.', 413)
        output = LimitedBuffer()
        await client.download_media(message, file=output)
        return output.getvalue(), message.file.name or ('Фото.jpg' if message.photo else 'Вложение')
