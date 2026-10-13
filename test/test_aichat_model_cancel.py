"""Cancelling a chat stream must not release priority before its HTTP worker ends.

The real route is compiled without importing the app or its scheduler. The model
and priority notifications are local fakes; storage and streaming are real.
"""
import ast
import asyncio
from contextlib import ExitStack
import json
import os
from pathlib import Path
import sys
import tempfile
import threading
import types
import unittest
from unittest.mock import AsyncMock, Mock, patch

from fastapi import HTTPException, Request

PROJECT_ROOT = Path(os.environ.get('AICHAT_PROJECT_ROOT', Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(PROJECT_ROOT))


class LocalForm(dict):
    def getlist(self, key):
        return []

    async def close(self):
        pass


class FinalModelCancellationTests(unittest.TestCase):
    def check_disconnect(self, *, provider_fails):
        import services.aichat as package
        from services.aichat import storage
        from services.aichat.teams import normalize_team

        started, released, ended = threading.Event(), threading.Event(), threading.Event()
        changes, active = [], []
        facts = types.ModuleType('services.aichat.waiting_facts')

        def begin_answer():
            self.assertEqual(active, [])
            active.append('model')
            changes.append('begin')

        def end_answer():
            self.assertTrue(released.is_set(), 'Priority cannot end before the blocking model returns')
            self.assertEqual(active, ['model'])
            active.pop()
            changes.append('end')
            ended.set()

        facts.begin_answer, facts.end_answer = begin_answer, end_answer

        def local_model(history, **kwargs):
            self.assertEqual(active, ['model'])
            started.set()
            if not released.wait(5):
                raise AssertionError('Test did not release its local HTTP worker')
            if provider_fails:
                raise RuntimeError('Synthetic provider failure after browser disconnected')
            return 'Ответ после закрытия вкладки'

        model = Mock(side_effect=local_model)
        context = types.ModuleType('services.aichat.report_context')
        context.wants_context = Mock(return_value=False)
        context.to_prompt = lambda value: ''
        roles = types.ModuleType('core.roles')
        roles.effective_modules = Mock(return_value=[])
        source = (PROJECT_ROOT / 'app.py').read_text(encoding='utf-8-sig')
        route = next(node for node in ast.parse(source).body
                     if isinstance(node, ast.AsyncFunctionDef) and node.name == 'aichat_send')
        route.decorator_list = []
        namespace = {'Request': Request, 'HTTPException': HTTPException, 'asyncio': asyncio,
                     'json': json, 'aichat_store': storage, 'validated_chat_team': normalize_team,
                     'aichat_ask': model, 'build_platform_context': Mock()}
        exec(compile(ast.Module(body=[route], type_ignores=[]), str(PROJECT_ROOT / 'app.py'), 'exec'), namespace)
        question = 'Объясни подробно принцип работы водопровода'
        request = Request({'type': 'http', 'method': 'POST', 'path': '/aichat/api/send',
                           'headers': [(b'accept', b'application/x-ndjson')],
                           'state': {'user': {'id': 1, 'username': 'offline-test', 'modules': []}}})
        request.form = AsyncMock(return_value=LocalForm(text=question, skills_mode='off'))

        async def scenario():
            response = await namespace['aichat_send'](request)
            try:
                event = json.loads(await asyncio.wait_for(response.body_iterator.__anext__(), 5))
                self.assertEqual((event['id'], event['status']), ('answer', 'running'))
                self.assertTrue(await asyncio.to_thread(started.wait, 5))
                self.assertEqual(changes, ['begin'])
                self.assertEqual(active, ['model'])

                await asyncio.wait_for(response.body_iterator.aclose(), 5)
                # Coroutine cancellation leaves the synchronous HTTP worker alive.
                self.assertFalse(ended.is_set())
                self.assertEqual(active, ['model'])
                self.assertEqual(changes, ['begin'])
                dialogs = json.loads(storage.DIALOGS_FILE.read_text(encoding='utf-8'))
                self.assertEqual(len(dialogs), 1)
                self.assertEqual([(message['role'], message['content']) for message in dialogs[0]['messages']],
                                 [('user', question)])
                before = storage.DIALOGS_FILE.read_bytes()

                released.set()
                self.assertTrue(await asyncio.to_thread(ended.wait, 5))
                await asyncio.sleep(0)
                self.assertEqual(changes, ['begin', 'end'])
                self.assertEqual(active, [])
                self.assertEqual(storage.DIALOGS_FILE.read_bytes(), before)
                model.assert_called_once()
            finally:
                released.set()
                await response.body_iterator.aclose()
                if started.is_set():
                    await asyncio.to_thread(ended.wait, 5)

        with tempfile.TemporaryDirectory(prefix='aichat-model-cancel-') as directory, ExitStack() as stack:
            root = Path(directory)
            stack.enter_context(patch.object(storage, 'DATA_DIR', root / 'chat'))
            stack.enter_context(patch.object(storage, 'DIALOGS_FILE', root / 'chat/dialogs.json'))
            stack.enter_context(patch.dict(sys.modules, {'services.aichat.waiting_facts': facts,
                                                         'services.aichat.report_context': context,
                                                         'core.roles': roles}))
            # The package may already have a cached real waiting_facts attribute.
            stack.enter_context(patch.object(package, 'waiting_facts', facts, create=True))
            stack.enter_context(patch('socket.create_connection', side_effect=AssertionError('No network in this test')))
            stack.enter_context(patch('socket.socket.connect', side_effect=AssertionError('No network in this test')))
            try:
                asyncio.run(scenario())
            finally:
                released.set()

    def test_disconnect_keeps_priority_until_successful_model_worker_finishes(self):
        self.check_disconnect(provider_fails=False)

    def test_disconnect_keeps_priority_until_failed_model_worker_finishes(self):
        self.check_disconnect(provider_fails=True)


if __name__ == '__main__':
    unittest.main()
