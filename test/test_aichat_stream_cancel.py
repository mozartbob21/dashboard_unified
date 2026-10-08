"""Offline regression for disconnects while a skill is being prepared.

Run with: python -m unittest test.test_aichat_stream_cancel
The production handler is compiled directly from app.py without importing app
or starting its scheduler. Request, StreamingResponse and JSON storage are real;
preparation, platform context and all network connections are mocked or denied.
"""
import ast
import asyncio
import json
import os
import sys
import tempfile
import threading
import types
import unittest
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import AsyncMock, Mock, patch

from fastapi import HTTPException, Request


PROJECT_ROOT = Path(os.environ.get('AICHAT_PROJECT_ROOT', Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(PROJECT_ROOT))


class LocalForm(dict):
    """Already parsed form with no files or external resources."""
    closed = False

    def getlist(self, key):
        return []

    async def close(self):
        self.closed = True


class StreamingCancellationTests(unittest.TestCase):
    def check_disconnect(self, *, explicit_report):
        from services.aichat import storage
        from services.aichat.teams import normalize_team

        question = 'Подготовь краткий итог по предоставленным данным'
        report_scope = {'module': 'edo', 'source': 'synthetic-source',
                        'municipality': 'Тестовый округ', 'modules': ['edo'],
                        'sources': ['synthetic-source'], 'municipalities': ['Тестовый округ']}
        report = {'selection': report_scope, 'sources': [], 'clarification': ''}
        released, started, finished = threading.Event(), threading.Event(), threading.Event()
        prepare_calls = []

        def local_prepare(text, **kwargs):
            prepare_calls.append(text)
            started.set()
            try:
                kwargs['emit']({'type': 'step', 'id': 'skill-routing', 'kind': 'routing',
                                'status': 'running', 'label': 'Подбираю подходящий навык'})
                if not released.wait(5):
                    raise AssertionError('Test did not release its local preparation worker')
                return {'skills': [], 'tools': [], 'warnings': [], 'instructions': '', 'evidence': ''}
            finally:
                finished.set()

        runtime = types.ModuleType('services.aichat.skills.runtime')
        runtime.prepare = local_prepare
        context = types.ModuleType('services.aichat.report_context')
        context.wants_context = Mock(return_value=False)
        context.to_prompt = lambda value: json.dumps(value, ensure_ascii=False)
        roles = types.ModuleType('core.roles')
        roles.effective_modules = Mock(return_value=['edo'])
        model = Mock(side_effect=AssertionError('The model must not run after preparation is cancelled'))
        platform = Mock(return_value=report)

        source = (PROJECT_ROOT / 'app.py').read_text(encoding='utf-8-sig')
        route = next(node for node in ast.parse(source).body
                     if isinstance(node, ast.AsyncFunctionDef) and node.name == 'aichat_send')
        route.decorator_list = []
        namespace = {'Request': Request, 'HTTPException': HTTPException,
                     'asyncio': asyncio, 'json': json, 'aichat_store': storage,
                     'validated_chat_team': normalize_team, 'aichat_ask': model,
                     'build_platform_context': platform}
        exec(compile(ast.Module(body=[route], type_ignores=[]), str(PROJECT_ROOT / 'app.py'), 'exec'), namespace)

        form = LocalForm(text=question, skills_mode='auto')
        if explicit_report:
            form.update(include_context='true', context_module='edo',
                        context_source='synthetic-source', context_municipality='Тестовый округ')
        request = Request({'type': 'http', 'method': 'POST', 'path': '/aichat/api/send',
                           'headers': [(b'accept', b'application/x-ndjson')],
                           'state': {'user': {'id': 1, 'username': 'offline-test', 'modules': ['edo']}}})
        request.form = AsyncMock(return_value=form)

        async def scenario():
            response = await namespace['aichat_send'](request)
            self.assertTrue(form.closed)
            self.assertIn('application/x-ndjson', response.media_type)
            observed = []
            try:
                # Explicit context emits two real events before preparation.
                # Disconnect at the first preparation event, while its worker
                # is still blocked, so no model or assistant can finish first.
                for _ in range(4):
                    event = json.loads(await asyncio.wait_for(response.body_iterator.__anext__(), 5))
                    observed.append(event)
                    if event.get('id') == 'skill-routing':
                        break
                self.assertEqual(observed[-1]['id'], 'skill-routing')
                self.assertEqual(observed[-1]['status'], 'running')
                self.assertTrue(started.is_set())
                self.assertFalse(finished.is_set())
                if not explicit_report:
                    self.assertEqual(len(observed), 1)

                await asyncio.wait_for(response.body_iterator.aclose(), 5)

                # Read the real persisted file, before letting preparation exit.
                dialogs = json.loads(storage.DIALOGS_FILE.read_text(encoding='utf-8'))
                self.assertEqual(len(dialogs), 1)
                self.assertEqual(len(dialogs[0]['messages']), 1)
                message = dialogs[0]['messages'][0]
                self.assertEqual(message['role'], 'user')
                self.assertEqual(message['content'], question)
                self.assertNotIn('skill_run', message)
                self.assertEqual(prepare_calls, [question])
                model.assert_not_called()
                if explicit_report:
                    self.assertEqual(message['report_scope'], report_scope)
                    platform.assert_called_once()
                    self.assertEqual(platform.call_args.args[1], ['edo'])
                else:
                    self.assertNotIn('report_scope', message)
                    platform.assert_not_called()

                # Finishing the already running local thread must not resume
                # the cancelled coroutine, duplicate the user or save an answer.
                before = storage.DIALOGS_FILE.read_bytes()
                released.set()
                self.assertTrue(await asyncio.to_thread(finished.wait, 5))
                await asyncio.sleep(0)
                self.assertEqual(storage.DIALOGS_FILE.read_bytes(), before)
                model.assert_not_called()
            finally:
                released.set()
                await response.body_iterator.aclose()
                if started.is_set():
                    await asyncio.to_thread(finished.wait, 5)

        with tempfile.TemporaryDirectory(prefix='aichat-cancel-offline-') as directory, ExitStack() as stack:
            root = Path(directory)
            stack.enter_context(patch.object(storage, 'DATA_DIR', root / 'chat'))
            stack.enter_context(patch.object(storage, 'DIALOGS_FILE', root / 'chat/dialogs.json'))
            stack.enter_context(patch.dict(sys.modules, {'services.aichat.skills.runtime': runtime,
                                                         'services.aichat.report_context': context,
                                                         'core.roles': roles}))
            stack.enter_context(patch('socket.create_connection', side_effect=AssertionError('Network disabled in this test')))
            stack.enter_context(patch('socket.socket.connect', side_effect=AssertionError('Network disabled in this test')))
            try:
                asyncio.run(scenario())
            finally:
                released.set()

    def test_disconnect_during_prepare_keeps_user_without_assistant(self):
        self.check_disconnect(explicit_report=False)

    def test_disconnect_during_prepare_keeps_completed_explicit_report_scope(self):
        self.check_disconnect(explicit_report=True)


if __name__ == '__main__':
    unittest.main()
