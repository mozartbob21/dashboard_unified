"""Skill HTTP integration; temporary accounts/storage and entirely mocked Qwen."""
import json
from pathlib import Path
from unittest.mock import patch
import unittest

import test.test_access_control as access_fixture


class SkillAPITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        access_fixture.AccessControlTests.setUpClass.__func__(cls)

    @classmethod
    def tearDownClass(cls):
        access_fixture.AccessControlTests.tearDownClass.__func__(cls)

    def setUp(self):
        access_fixture.AccessControlTests.setUp(self)
        from services.aichat import storage, report_context
        self.store = storage
        directory = Path(self.temp.name) / self._testMethodName
        for obj, attr, value in [(storage, 'DATA_DIR', directory / 'chat'),
                                 (storage, 'DIALOGS_FILE', directory / 'chat/dialogs.json'),
                                 (report_context, 'DATA_DIR', directory / 'reports')]:
            guard = patch.object(obj, attr, value)
            guard.start(); self.addCleanup(guard.stop)
        # Any unmocked model network attempt is a test failure, never a live call.
        guard = patch('httpx.post', side_effect=AssertionError('No network in skill tests'))
        guard.start(); self.addCleanup(guard.stop)
        self.login()

    def tearDown(self):
        access_fixture.AccessControlTests.tearDown(self)

    def login(self, username='ordinary', password='Test-password-123'):
        return access_fixture.AccessControlTests.login(self, username, password)

    def send(self, **data):
        return self.client.post('/aichat/api/send', data={'text': 'Вычисли 12.5 * 8', 'skills_mode': 'auto', **data},
                                headers={'Accept': 'application/x-ndjson'})

    def events(self, response):
        self.assertEqual(response.status_code, 200, response.text)
        self.assertIn('application/x-ndjson', response.headers['content-type'])
        return [json.loads(line) for line in response.text.splitlines() if line]

    def test_catalog_toggle_is_account_scoped_and_server_owned(self):
        first = self.client.get('/aichat/api/skills').json()
        self.assertEqual(first['scope'], 'account')
        self.assertGreaterEqual(len(first['items']), 7)
        self.assertFalse(any('body' in item or 'instructions' in item for item in first['items']))
        changed = self.client.patch('/aichat/api/skills/calculations', json={'enabled': False})
        self.assertEqual(changed.status_code, 200)
        self.assertFalse(next(item for item in changed.json()['items'] if item['id'] == 'calculations')['enabled'])
        self.client.cookies.clear(); self.login('legacy')
        other = self.client.get('/aichat/api/skills').json()
        self.assertTrue(next(item for item in other['items'] if item['id'] == 'calculations')['enabled'])

    def test_invalid_or_foreign_settings_never_change_state(self):
        for payload in [{'enabled': 'false'}, {'enabled': 1}, {}, {'enabled': True, 'prompt': 'override'}]:
            self.assertEqual(self.client.patch('/aichat/api/skills/calculations', json=payload).status_code, 400)
        self.assertEqual(self.client.patch('/aichat/api/skills/missing', json={'enabled': False}).status_code, 404)
        self.assertEqual(self.client.patch('/aichat/api/skills/calculations', json={'enabled': False},
                                          headers={'Origin': 'https://invalid.example'}).status_code, 403)
        self.client.cookies.clear()
        self.assertEqual(self.client.get('/aichat/api/skills', follow_redirects=False).status_code, 302)
        self.assertEqual(self.client.patch('/aichat/api/skills/calculations', json={'enabled': True}).status_code, 401)

    def test_stream_executes_real_calculator_and_persists_only_public_trace(self):
        from services.aichat import engine
        from services.aichat.skills import runtime
        replies = ['{"skills":["calculations"]}',
                   '{"actions":[{"skill":"calculations","tool":"calculate","args":{"expression":"12.5 * 8"}}]}']
        with patch.object(runtime, '_model', side_effect=replies) as planner, \
                patch.object(engine, '_qwen_chat', return_value='Результат: 100.') as qwen:
            events = self.events(self.send())
        result = events[-1]
        self.assertEqual(result['type'], 'result')
        self.assertTrue(result['ok'])
        self.assertEqual(result['skill_run']['skills'][0]['id'], 'calculations')
        self.assertEqual(result['skill_run']['tools'][0]['status'], 'done')
        self.assertTrue(any(event.get('kind') == 'tool' and event.get('status') == 'running' for event in events))
        sent = qwen.call_args.args[0]
        self.assertIn('100', sent[-1]['content'])
        self.assertIn('<skill_data>', sent[-1]['content'])
        self.assertNotIn('Инструкции выбранных', str(planner.call_args_list[0].args))
        dialog = self.client.get('/aichat/api/dialogs/' + result['dialog_id']).json()
        trace = dialog['messages'][-1]['skill_run']
        self.assertEqual(trace, result['skill_run'])
        self.assertNotIn('expression', json.dumps(trace))
        self.assertNotIn('instructions', json.dumps(trace))
        self.assertEqual([m['role'] for m in dialog['messages']], ['user', 'assistant'])

    def test_all_disabled_skips_router_and_keeps_ordinary_chat(self):
        from services.aichat.skills import runtime
        items = self.client.get('/aichat/api/skills').json()['items']
        for item in items:
            self.client.patch('/aichat/api/skills/' + item['id'], json={'enabled': False})
        with patch.object(runtime, '_model') as planner, patch.object(self.module, 'aichat_ask', return_value='Привет') as answer:
            result = self.events(self.send(text='Привет'))[-1]
        planner.assert_not_called()
        self.assertEqual(result['skill_run']['skills'], [])
        self.assertNotIn('skill_instructions', answer.call_args.kwargs)

    def test_skill_tools_cannot_expand_platform_grants(self):
        from services.aichat.skills import runtime
        from services.aichat import report_context
        replies = ['{"skills":["platform-report"]}',
                   '{"actions":[{"skill":"platform-report","tool":"platform_report","args":{}}]}']
        with patch.object(runtime, '_model', side_effect=replies), \
                patch.object(report_context, '_render_module') as read, \
                patch.object(self.module, 'aichat_ask') as answer:
            result = self.events(self.send(text='Дай отчёт по МинЖКХ'))[-1]
        self.assertEqual(result['response_kind'], 'clarification')
        self.assertEqual(result['report_sources'], [])
        read.assert_not_called(); answer.assert_not_called()

    def test_explicit_report_setting_is_honored_in_auto_mode(self):
        from services.aichat.skills import runtime
        plan = ['{"skills":["platform-report"]}',
                '{"actions":[{"skill":"platform-report","tool":"platform_report","args":{}}]}']
        with patch.object(runtime, '_model', side_effect=plan), \
                patch.object(self.module, 'build_platform_context') as read, \
                patch.object(self.module, 'aichat_ask', return_value='Без данных') as answer:
            result = self.events(self.send(include_context='false'))[-1]
        read.assert_not_called()
        self.assertEqual(answer.call_args.kwargs['platform_context'], '')
        self.assertEqual(result['skill_run']['tools'][0]['status'], 'error')
        bundle = {'selection': {'modules': ['edo']}, 'sources': [], 'clarification': ''}
        with patch.object(runtime, '_model', return_value='{"skills":[]}'), \
                patch.object(self.module, 'build_platform_context', return_value=bundle) as read, \
                patch.object(self.module, 'aichat_ask', return_value='Сводка') as answer:
            result = self.events(self.send(include_context='true', context_module='edo'))[-1]
        read.assert_called_once()
        self.assertTrue(answer.call_args.kwargs['platform_context'])
        saved = self.store.get_dialog(result['dialog_id'])['messages'][0]
        self.assertEqual(saved['report_scope']['modules'], ['edo'])

    def test_failure_closes_running_step_and_never_exposes_exception(self):
        from services.aichat.skills import runtime
        with patch.object(runtime, '_model', return_value='{"skills":[]}'), \
                patch.object(self.module, 'aichat_ask', side_effect=RuntimeError('PRIVATE_KEY_AND_PROMPT')):
            response = self.send(text='Привет')
        result = self.events(response)[-1]
        self.assertFalse(result['ok'])
        self.assertEqual(result['error_code'], 'AI_INTERNAL_ERROR')
        self.assertNotIn('PRIVATE_KEY_AND_PROMPT', response.text)
        self.assertFalse(any(item['status'] == 'running' for item in result['skill_run']['steps']))
        self.assertEqual(result['skill_run']['steps'][-1]['status'], 'error')

    def test_validation_happens_before_stream_or_dialog_write(self):
        for data, status in [({'text': '', 'skills_mode': 'auto'}, 400),
                             ({'text': 'x', 'skills_mode': 'invalid'}, 400),
                             ({'text': 'x', 'dialog_id': 'missing'}, 404)]:
            with patch.object(self.module, 'aichat_ask') as answer:
                response = self.send(**data)
            self.assertEqual(response.status_code, status)
            answer.assert_not_called()
        self.assertFalse(self.store.DIALOGS_FILE.exists())

    def test_attachment_contents_are_not_sent_to_router_and_are_read_by_tool(self):
        from services.aichat.skills import runtime
        replies = ['{"skills":["document-analysis"]}',
                   '{"actions":[{"skill":"document-analysis","tool":"attachment_text","args":{"attachment":"note.txt"}}]}']
        with patch.object(runtime, '_model', side_effect=replies) as planner, \
                patch.object(self.module, 'aichat_ask', return_value='Краткий итог') as answer:
            response = self.client.post('/aichat/api/send', data={'text': 'Разбери документ', 'skills_mode': 'auto'},
                                        files={'file': ('note.txt', b'SYNTHETIC_DOCUMENT_CONTENT', 'text/plain')},
                                        headers={'Accept': 'application/x-ndjson'})
        self.assertTrue(self.events(response)[-1]['ok'])
        self.assertNotIn('SYNTHETIC_DOCUMENT_CONTENT', str(planner.call_args_list))
        self.assertIn('SYNTHETIC_DOCUMENT_CONTENT', answer.call_args.kwargs['tool_evidence'])
        self.assertNotIn('SYNTHETIC_DOCUMENT_CONTENT', answer.call_args.args[0][-1]['content'])


if __name__ == '__main__':
    unittest.main()
