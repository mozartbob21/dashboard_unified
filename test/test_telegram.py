import copy
import json
import os
import tempfile
import unittest
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

from fastapi import FastAPI, HTTPException, Request
from fastapi.testclient import TestClient
from starlette.requests import Request as StarletteRequest

from services.telegram import client as tg
from services.telegram.store import Store, Busy
from routers import telegram as routes


def account(name='Alice', chat='100'):
    return {'stage': 'ready', 'session': name + '-private-session', 'generation': name,
            'profile': {'name': name, 'username': name}, 'selected': [chat], 'sent': {},
            'dialogs': {chat: {'id': chat, 'title': name + ' private chat', 'kind': 'user',
                               'peer': {'_': 'InputPeerUser', 'user_id': int(chat), 'access_hash': 123}}}}


class StoreTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.store = Store(self.temp.name)

    def test_encryption_and_reload_isolate_users(self):
        self.store.put('alice', account())
        self.store.put('bob', account('Bob', '200'))
        reopened = Store(self.temp.name)
        self.assertEqual(reopened.get('alice')['session'], 'Alice-private-session')
        self.assertEqual(reopened.get('bob')['session'], 'Bob-private-session')
        self.assertEqual(reopened.get('unknown'), {})
        self.assertNotIn(b'Alice-private-session', self.store.path.read_bytes())
        self.store.delete('alice')
        self.assertEqual(self.store.get('alice'), {})
        self.assertTrue(self.store.get('bob'))

    def test_lease_is_shared_between_workers_and_released_after_error(self):
        another_worker = Store(self.temp.name)
        with self.store.claim('alice'):
            with self.assertRaises(Busy):
                with another_worker.claim('alice'):
                    pass
            with another_worker.claim('bob'):
                pass
        with another_worker.claim('alice'):
            pass

    def test_missing_key_does_not_silently_replace_existing_sessions(self):
        self.store.put('alice', account())
        (Path(self.temp.name) / 'session.key').unlink()
        with self.assertRaises(RuntimeError):
            Store(self.temp.name)

    def test_web_served_storage_is_rejected(self):
        root = Path(tg.__file__).resolve().parents[2] / 'data' / 'telegram-test'
        with self.assertRaises(RuntimeError):
            Store(root)
        self.assertFalse(root.exists())


class ApiTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.store = Store(self.temp.name)
        self.store.put('alice', account())
        self.store.put('bob', account('Bob', '200'))
        for target, value in [('storage', self.store), ('settings', [])]:
            patcher = patch.object(tg, target, return_value=value)
            patcher.start();self.addCleanup(patcher.stop)
        app = FastAPI()
        app.include_router(routes.router)
        async def owner(request: Request):
            value = request.headers.get('test-owner')
            if value not in {'alice', 'bob'}:
                raise HTTPException(401)
            return value
        app.dependency_overrides[routes.require_user] = owner
        self.api = TestClient(app)
        self.headers = {'test-owner': 'alice', 'X-Neurona-Telegram': '1'}

    def test_status_never_discloses_session_and_uses_current_owner(self):
        for owner in ['alice', 'bob']:
            response = self.api.get('/telegram/api/status', headers={'test-owner': owner})
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.headers['cache-control'], 'no-store')
            self.assertEqual(response.json()['profile']['name'].lower(), owner)
            self.assertNotIn('session', response.text)
        self.assertEqual(self.api.get('/telegram/api/status').status_code, 401)

    def test_cannot_select_other_users_chat_or_submit_owner(self):
        response = self.api.put('/telegram/api/chats', headers=self.headers, json={'ids': ['200']})
        self.assertEqual(response.status_code, 403)
        response = self.api.put('/telegram/api/chats', headers=self.headers, json={'ids': ['100'], 'owner': 'bob'})
        self.assertEqual(response.status_code, 422)
        self.assertEqual(self.store.get('bob')['selected'], ['200'])

    def test_deselecting_chat_blocks_read_send_download(self):
        response = self.api.put('/telegram/api/chats', headers=self.headers, json={'ids': []})
        self.assertEqual(response.status_code, 200)
        with self.assertRaises(tg.TelegramError) as error:
            tg._selected(self.store.get('alice'), '100')
        self.assertEqual(error.exception.status, 403)

    def test_mutations_require_same_origin_and_custom_header(self):
        self.assertEqual(self.api.put('/telegram/api/chats', headers={'test-owner': 'alice'}, json={'ids': []}).status_code, 403)
        self.assertEqual(self.api.put('/telegram/api/chats', headers={**self.headers, 'Origin': 'https://untrusted.test'}, json={'ids': []}).status_code, 403)

    def test_sending_validation_and_identity(self):
        route = '/telegram/api/chats/100/messages'
        with patch.object(tg, 'send', new_callable=AsyncMock, return_value={'ok': True}) as send:
            response = self.api.post(route, headers=self.headers, json={'text': 'Тест', 'request_id': 'b8fb59c2-b565-46bc-b194-e0b148a642ef'})
            self.assertEqual(response.status_code, 200)
            self.assertEqual(send.call_args.args[0], 'alice')
            for text in ['', ' ', 'a' * 4097]:
                response = self.api.post(route, headers=self.headers, json={'text': text, 'request_id': 'b8fb59c2-b565-46bc-b194-e0b148a642ef'})
                self.assertEqual(response.status_code, 422)
            self.assertEqual(send.call_count, 1)

    def test_flood_wait_returns_retry_after(self):
        with patch.object(tg, 'dialogs', new_callable=AsyncMock, side_effect=tg.TelegramError('Подождите', 429, 30)):
            response = self.api.get('/telegram/api/dialogs', headers=self.headers)
            self.assertEqual(response.status_code, 429)
            self.assertEqual(response.headers['retry-after'], '30')

    def test_attachment_is_private_download_not_executable_html(self):
        with patch.object(tg, 'attachment', new_callable=AsyncMock, return_value=(b'<script>test</script>', 'test.html')):
            response = self.api.get('/telegram/api/chats/100/attachments/1', headers=self.headers)
            self.assertEqual(response.headers['content-type'], 'application/octet-stream')
            self.assertIn('attachment;', response.headers['content-disposition'])
            self.assertEqual(response.headers['cache-control'], 'no-store')

    def test_real_identity_requires_module_and_stable_db_id(self):
        request = StarletteRequest({'type': 'http', 'method': 'GET', 'path': '/telegram', 'headers': []})
        with self.assertRaises(HTTPException):
            routes.require_user(request)
        request.state.user = {'username': 'alice', 'modules': []}
        with self.assertRaises(HTTPException):
            routes.require_user(request)
        request.state.user['modules'] = ['telegram']
        @contextmanager
        def db():
            yield SimpleNamespace(execute=lambda *args: SimpleNamespace(fetchone=lambda: {'id': 42, 'created_at': '2026-01-01'}))
        with patch.object(routes, 'get_db_connection', db):
            identity = routes.require_user(request)
            request.state.user['username'] = 'renamed'
            self.assertEqual(routes.require_user(request), identity)


class ClientTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.store = Store(self.temp.name)
        self.store.put('alice', account())
        self.store.put('bob', account('Bob', '200'))
        self.client = SimpleNamespace(connect=AsyncMock(), disconnect=AsyncMock(), is_user_authorized=AsyncMock(return_value=True),
                                      session=SimpleNamespace(save=lambda: 'saved-secret'))
        for name, value in [('storage', self.store), ('settings', []), ('_client', self.client)]:
            patcher = patch.object(tg, name, return_value=value)
            patcher.start();self.addCleanup(patcher.stop)

    async def test_operation_connects_only_current_users_session(self):
        async with tg.connection('alice'):
            tg._client.assert_called_with('Alice-private-session')
        self.client.disconnect.assert_awaited_once()

    async def test_revoked_session_requires_reconnect(self):
        self.client.is_user_authorized.return_value = False
        with self.assertRaises(tg.TelegramError) as error:
            async with tg.connection('alice'):
                self.fail('Revoked session must not enter operation')
        self.assertEqual(error.exception.status, 401)
        self.assertEqual(self.store.get('alice')['stage'], 'expired')
        self.assertEqual(self.store.get('bob')['stage'], 'ready')

    async def test_code_2fa_and_password_are_not_stored(self):
        from telethon.errors import SessionPasswordNeededError
        from telethon.tl.types import User
        self.store.delete('alice')
        self.client.send_code_request = AsyncMock(return_value=SimpleNamespace(phone_code_hash='hash'))
        self.client.sign_in = AsyncMock(side_effect=[SessionPasswordNeededError(request=None), None])
        self.client.get_me = AsyncMock(return_value=User(id=7, first_name='Alice', username='alice'))
        await tg.start_login('alice', '+79991112233')
        result = await tg.finish_login('alice', code='12345')
        self.assertEqual(result['stage'], 'password')
        result = await tg.finish_login('alice', password='my-password-secret')
        self.assertEqual(result['stage'], 'ready')
        saved = json.dumps(self.store.get('alice'))
        self.assertNotIn('12345', saved)
        self.assertNotIn('my-password-secret', saved)
        self.assertNotIn('phone_code_hash', saved)

    async def test_send_retries_reuse_telegram_random_id_and_skip_completed_request(self):
        fake = AsyncMock()
        fake.connect = AsyncMock();fake.disconnect = AsyncMock()
        fake.is_user_authorized = AsyncMock(return_value=True)
        fake.side_effect = [TimeoutError(), SimpleNamespace()]
        with patch.object(tg, '_client', return_value=fake):
            with self.assertRaises(tg.TelegramError):
                await tg.send('alice', '100', 'Сообщение', 'retry-id')
            await tg.send('alice', '100', 'Сообщение', 'retry-id')
            self.assertEqual(fake.call_args_list[0].args[0].random_id, fake.call_args_list[1].args[0].random_id)
            self.assertTrue((await tg.send('alice', '100', 'Сообщение', 'retry-id'))['already_sent'])
            self.assertEqual(fake.call_count, 2)
            with self.assertRaises(tg.TelegramError):
                await tg.send('alice', '100', 'Изменённое', 'retry-id')

    async def test_unselected_chat_cannot_be_read_sent_or_downloaded(self):
        for operation in [tg.messages('alice', '200'), tg.send('alice', '200', 'test', 'id'), tg.attachment('alice', '200', 1)]:
            with self.assertRaises(tg.TelegramError) as error:
                await operation
            self.assertEqual(error.exception.status, 403)

    async def test_telegram_duplicate_confirms_previously_delivered_message(self):
        from telethon.errors import RandomIdDuplicateError
        fake = AsyncMock(side_effect=RandomIdDuplicateError(request=None))
        fake.is_user_authorized = AsyncMock(return_value=True)
        with patch.object(tg, '_client', return_value=fake):
            result = await tg.send('alice', '100', 'Текст', 'previous-attempt')
        self.assertTrue(result['already_sent'])
        self.assertTrue(self.store.get('alice')['sent']['previous-attempt']['done'])

    async def test_disconnect_revokes_only_current_account(self):
        self.client.log_out = AsyncMock(return_value=True)
        await tg.disconnect('alice')
        self.client.log_out.assert_awaited_once()
        self.assertEqual(self.store.get('alice'), {})
        self.assertEqual(self.store.get('bob')['stage'], 'ready')

    def test_invalid_proxy_is_configuration_error(self):
        for url in ['socks5://localhost:not-a-port', 'ftp://localhost:21', 'socks5://[invalid']:
            with patch.dict(os.environ, {'TELEGRAM_PROXY_URL': url}):
                with self.assertRaises(tg.TelegramError) as error:
                    tg.proxy_settings()
                self.assertEqual(error.exception.status, 503)

    async def test_message_plain_text_is_returned_without_html_rendering(self):
        from telethon.tl.types import User
        self.client.get_messages = AsyncMock(return_value=[SimpleNamespace(
            id=1, sender=User(id=9, first_name='Друг'), file=None, raw_text='<script>alert(1)</script>',
            out=False, date=datetime.now(timezone.utc), reply_to_msg_id=None, action=None, media=None)])
        result = await tg.messages('alice', '100')
        self.assertEqual(result['items'][0]['text'], '<script>alert(1)</script>')

    def test_media_buffer_stops_oversized_download(self):
        with patch.object(tg, 'MAX_MEDIA', 3):
            stream = tg.LimitedBuffer()
            stream.write(b'ab')
            with self.assertRaises(tg.TelegramError):
                stream.write(b'cd')


if __name__ == '__main__':
    unittest.main()
