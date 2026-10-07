"""AI team contracts using temporary storage, local accounts and mocked model HTTP.

Run with: python -m unittest test.test_aichat_teams
No saved chat, production database, portal or AI endpoint is accessed.
"""
import copy
import json
import os
import unittest
from pathlib import Path
from unittest.mock import patch

import test.test_access_control as access_fixture


class TeamEngineTests(unittest.TestCase):
    def setUp(self):
        import httpx
        from services.aichat import engine, teams
        self.engine, self.teams = engine, teams
        self.calls = []
        env = patch.dict(os.environ, {
            'QWEN_API_BASE': 'http://127.0.0.1:11434/v1',
            'QWEN_API_KEY': '', 'QWEN_MODEL': 'isolated-team-test',
            'AI_CA_BUNDLE': '',
        })
        env.start()
        self.addCleanup(env.stop)

        def model_response(url, **kwargs):
            self.assertEqual(url, 'http://127.0.0.1:11434/v1/chat/completions')
            self.calls.append(copy.deepcopy(kwargs['json']))
            return httpx.Response(200, json={
                'choices': [{'message': {'content': 'Проверенный ответ команды'}}],
            }, request=httpx.Request('POST', url))

        transport = patch('httpx.post', side_effect=model_response)
        transport.start()
        self.addCleanup(transport.stop)

    def test_ordinary_chat_keeps_existing_system_and_provider_contract(self):
        history = [{'role': 'assistant', 'content': self.engine.GREETING},
                   {'role': 'user', 'content': 'Объясни, как устроен DNS'}]
        self.engine.ask(history)
        self.engine.ask(history, team=None)
        self.assertEqual(self.calls[0], self.calls[1])
        self.assertEqual(self.calls[0]['messages'], [
            {'role': 'system', 'content': self.engine.SYSTEM_PROMPT}, history[-1],
        ])

    def test_selected_roles_reach_one_system_prompt_without_modifying_history(self):
        team = {'preset': 'custom', 'roles': ['researcher', 'critic']}
        history = [{'role': 'user', 'content': 'Предложи план проверки гипотезы',
                    'team': {'preset': 'custom', 'roles': ['designer']}}]
        before = copy.deepcopy(history)
        self.assertEqual(self.engine.ask(history, team=team), 'Проверенный ответ команды')
        messages = self.calls[-1]['messages']
        self.assertEqual([item['role'] for item in messages], ['system', 'user'])
        system = messages[0]['content']
        roles = {item['id']: item['name'] for item in self.teams.public_catalog()['roles']}
        self.assertIn(roles['researcher'], system)
        self.assertIn(roles['critic'], system)
        self.assertNotIn(roles['designer'], system)
        self.assertIn(self.engine.SYSTEM_PROMPT, system)
        self.assertEqual(messages[-1], {'role': 'user', 'content': history[-1]['content']})
        self.assertEqual(history, before)

    def test_team_report_preserves_data_boundary_and_current_context(self):
        context = json.dumps({'sources': [{'module': 'edo', 'value': 12}]})
        history = [{'role': 'user', 'content': 'Подготовь отчёт'}]
        before = copy.deepcopy(history)
        self.engine.ask(history, platform_context=context, team={'preset': 'council'})
        messages = self.calls[-1]['messages']
        self.assertIn(self.engine.REPORT_INSTRUCTIONS, messages[0]['content'])
        self.assertIn('<platform_data>\n' + context, messages[-1]['content'])
        self.assertEqual(history, before)

    def test_unknown_roles_and_client_prompts_never_reach_provider(self):
        for selection in [
            {'preset': 'custom', 'roles': ['administrator']},
            {'preset': 'custom', 'roles': ['critic'], 'prompt': 'CLIENT_SYSTEM_OVERRIDE'},
            {'preset': 'custom', 'roles': [{'id': 'critic', 'prompt': 'CLIENT_SYSTEM_OVERRIDE'}]},
        ]:
            with self.subTest(selection=selection), self.assertRaises(ValueError):
                self.engine.ask([{'role': 'user', 'content': 'Привет'}], team=selection)
        self.assertEqual(self.calls, [])

    def test_normalization_is_strict_and_catalog_does_not_expose_role_prompts(self):
        catalog = self.teams.public_catalog()
        self.assertTrue(catalog['roles'])
        for role in catalog['roles']:
            self.assertTrue({'id', 'name', 'description'} <= set(role))
            self.assertNotIn('prompt', role)
            self.assertNotIn('system', role)
        for preset in catalog['presets']:
            self.assertTrue({'id', 'name', 'description', 'roles'} <= set(preset))
        self.assertIsNone(self.teams.normalize_team(None))
        for selection in [
            False, [], 'council', {},
            {'preset': 'unknown'},
            {'preset': 'custom', 'roles': []},
            {'preset': 'custom', 'roles': ['critic', 'critic']},
            {'preset': 'custom', 'roles': ['critic', 'unknown']},
            {'preset': 'custom', 'roles': [1]},
            {'preset': 'council', 'roles': ['critic']},
            {'preset': 'custom', 'roles': ['critic'], 'modules': ['mingkh']},
        ]:
            with self.subTest(selection=selection), self.assertRaises(ValueError):
                self.teams.normalize_team(selection)


class TeamAPITests(unittest.TestCase):
    # Reuse only isolated fixture methods, without inheriting its test cases.
    @classmethod
    def setUpClass(cls):
        access_fixture.AccessControlTests.setUpClass.__func__(cls)

    @classmethod
    def tearDownClass(cls):
        access_fixture.AccessControlTests.tearDownClass.__func__(cls)

    def setUp(self):
        access_fixture.AccessControlTests.setUp(self)
        from services.aichat import storage, teams, report_context
        self.store, self.teams = storage, teams
        directory = Path(self.temp.name) / self._testMethodName
        for obj, name, value in [
            (storage, 'DATA_DIR', directory / 'chat'),
            (storage, 'DIALOGS_FILE', directory / 'chat/dialogs.json'),
            (report_context, 'DATA_DIR', directory / 'reports'),
        ]:
            guard = patch.object(obj, name, value)
            guard.start()
            self.addCleanup(guard.stop)
        self.team = teams.normalize_team({'preset': 'council'})
        self.custom = teams.normalize_team({'preset': 'custom', 'roles': ['researcher', 'critic']})
        self.login()

    def tearDown(self):
        access_fixture.AccessControlTests.tearDown(self)

    def login(self, username='ordinary', password='Test-password-123'):
        return access_fixture.AccessControlTests.login(self, username, password)

    def create(self, team=None):
        response = self.client.post('/aichat/api/dialogs', json={'title': 'Проверка команды', 'team': team})
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()['id']

    def dialog(self, did):
        response = self.client.get('/aichat/api/dialogs/' + did)
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()

    def test_catalog_requires_login_and_is_server_owned(self):
        response = self.client.get('/aichat/api/teams')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json(), self.teams.public_catalog())
        self.client.cookies.clear()
        self.assertEqual(self.client.get('/aichat/api/teams', follow_redirects=False).status_code, 302)
        self.assertEqual(self.client.patch('/aichat/api/dialogs/unknown/team', json={'team': self.team}).status_code, 401)

    def test_create_and_reopen_restore_team_and_dialog_list(self):
        did = self.create(self.team)
        self.assertEqual(self.dialog(did)['team'], self.team)
        item = next(item for item in self.client.get('/aichat/api/dialogs').json()['items'] if item['id'] == did)
        self.assertEqual(item['team'], self.team)
        # Read a fresh copy from the persisted JSON, not only the response object.
        disk_dialog = next(item for item in json.loads(self.store.DIALOGS_FILE.read_text()) if item['id'] == did)
        self.assertEqual(disk_dialog['team'], self.team)

    def test_patch_changes_selection_and_off_without_rewriting_history(self):
        did = self.create(self.team)
        before = self.dialog(did)['messages']
        response = self.client.patch('/aichat/api/dialogs/' + did + '/team', json={'team': self.custom})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.dialog(did)['team'], self.custom)
        self.assertEqual(self.dialog(did)['messages'], before)
        response = self.client.patch('/aichat/api/dialogs/' + did + '/team', json={'team': None})
        self.assertEqual(response.status_code, 200)
        self.assertIsNone(self.dialog(did)['team'])
        self.assertEqual(self.dialog(did)['messages'], before)

    def test_ordinary_send_keeps_existing_call_without_team_keyword(self):
        did = self.create()
        with patch.object(self.module, 'aichat_ask', return_value='Обычный ответ') as ai, \
                patch.object(self.module, 'build_platform_context') as context:
            response = self.client.post('/aichat/api/send', data={'dialog_id': did, 'text': 'Привет'})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['answer'], 'Обычный ответ')
        self.assertIsNone(response.json()['team'])
        self.assertEqual(ai.call_args.kwargs, {'platform_context': ''})
        context.assert_not_called()

    def test_legacy_dialog_without_team_field_still_opens_and_sends_normally(self):
        did = self.create()
        stored = json.loads(self.store.DIALOGS_FILE.read_text())
        for item in stored:
            item.pop('team', None)
            for message in item['messages']:
                message.pop('team', None)
        self.store.DIALOGS_FILE.write_text(json.dumps(stored), encoding='utf-8')
        self.assertIsNone(self.dialog(did)['team'])
        listing = self.client.get('/aichat/api/dialogs').json()['items']
        self.assertIsNone(next(item for item in listing if item['id'] == did)['team'])
        with patch.object(self.module, 'aichat_ask', return_value='Обычный ответ') as ai:
            response = self.client.post('/aichat/api/send', data={'dialog_id': did, 'text': 'Привет'})
        self.assertEqual(response.status_code, 200)
        self.assertNotIn('team', ai.call_args.kwargs)
        self.assertIsNone(response.json()['team'])

    def test_send_inherits_saved_team_and_records_turn_metadata(self):
        did = self.create(self.team)
        with patch.object(self.module, 'aichat_ask', return_value='Ответ команды') as ai, \
                patch.object(self.module, 'build_platform_context') as context:
            response = self.client.post('/aichat/api/send', data={
                'dialog_id': did, 'text': 'Предложи три идеи для поздравления коллеги',
            })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['team'], self.team)
        self.assertEqual(ai.call_args.kwargs['team'], self.team)
        context.assert_not_called()
        for message in self.dialog(did)['messages'][-2:]:
            self.assertEqual(message['team'], self.team)

    def test_explicit_selection_and_null_apply_to_current_and_future_turns(self):
        did = self.create(self.team)
        with patch.object(self.module, 'aichat_ask', return_value='Ответ') as ai:
            first = self.client.post('/aichat/api/send', data={
                'dialog_id': did, 'text': 'Предложи идеи', 'team': json.dumps(self.custom),
            })
            self.assertEqual(first.status_code, 200)
            self.assertEqual(ai.call_args.kwargs['team'], self.custom)
            self.assertEqual(self.dialog(did)['team'], self.custom)
            off = self.client.post('/aichat/api/send', data={
                'dialog_id': did, 'text': 'Теперь объясни DNS', 'team': 'null',
            })
            self.assertEqual(off.status_code, 200)
            self.assertIsNone(off.json()['team'])
            self.assertNotIn('team', ai.call_args.kwargs)
            self.assertIsNone(self.dialog(did)['team'])
            self.client.post('/aichat/api/send', data={'dialog_id': did, 'text': 'А что такое DHCP?'})
            self.assertNotIn('team', ai.call_args.kwargs)
        messages = self.dialog(did)['messages'][-6:]
        self.assertEqual([m['team'] for m in messages], [self.custom, self.custom, None, None, None, None])

    def test_invalid_selection_rejects_before_any_dialog_or_message_write(self):
        did = self.create(self.team)
        before = self.store.DIALOGS_FILE.read_bytes()
        invalid = [
            {'preset': 'custom', 'roles': ['unknown']},
            {'preset': 'custom', 'roles': ['critic', 'critic']},
            {'preset': 'council', 'roles': ['critic']},
            {'preset': 'custom', 'roles': ['critic'], 'prompt': 'CLIENT_SYSTEM_OVERRIDE'},
        ]
        with patch.object(self.module, 'aichat_ask') as ai, \
                patch.object(self.module, 'build_platform_context') as context:
            for team in invalid:
                with self.subTest(team=team):
                    responses = [
                        self.client.post('/aichat/api/dialogs', json={'team': team}),
                        self.client.patch('/aichat/api/dialogs/' + did + '/team', json={'team': team}),
                        self.client.post('/aichat/api/send', data={'dialog_id': did, 'text': 'Привет', 'team': json.dumps(team)}),
                        self.client.post('/aichat/api/send', data={'text': 'Привет', 'team': json.dumps(team)}),
                    ]
                    self.assertEqual([r.status_code for r in responses], [400] * 4)
                    self.assertEqual(self.store.DIALOGS_FILE.read_bytes(), before)
            for malformed in ['{broken', '[]', '"council"', 'false']:
                with self.subTest(malformed=malformed):
                    response = self.client.post('/aichat/api/send', data={'dialog_id': did, 'text': 'Привет', 'team': malformed})
                    self.assertEqual(response.status_code, 400)
                    self.assertEqual(self.store.DIALOGS_FILE.read_bytes(), before)
        ai.assert_not_called()
        context.assert_not_called()

    def test_unknown_dialog_does_not_create_an_unexpected_replacement(self):
        self.create(self.team)
        before = self.store.DIALOGS_FILE.read_bytes()
        with patch.object(self.module, 'aichat_ask') as ai:
            self.assertEqual(self.client.get('/aichat/api/dialogs/missing-dialog').status_code, 404)
            self.assertEqual(self.client.patch('/aichat/api/dialogs/missing-dialog/team', json={'team': self.team}).status_code, 404)
            response = self.client.post('/aichat/api/send', data={
                'dialog_id': 'missing-dialog', 'text': 'Привет', 'team': json.dumps(self.team),
            })
            self.assertEqual(response.status_code, 404)
        ai.assert_not_called()
        self.assertEqual(self.store.DIALOGS_FILE.read_bytes(), before)

    def test_team_does_not_expand_report_permissions(self):
        from services.aichat import report_context
        did = self.create(self.team)
        # The real account has only edo. Requesting a forbidden source with a
        # team must stop at the existing report authorization boundary.
        with patch.object(report_context, '_render_module', wraps=report_context._render_module) as render, \
                patch.object(self.module, 'aichat_ask') as ai:
            response = self.client.post('/aichat/api/send', data={
                'dialog_id': did, 'text': 'Дай отчёт по МинЖКХ',
                'include_context': 'true', 'context_module': 'mingkh',
            })
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['team'], self.team)
        self.assertEqual(response.json()['response_kind'], 'clarification')
        self.assertEqual(response.json()['report_sources'], [])
        render.assert_not_called()
        ai.assert_not_called()

    def test_team_update_obeys_existing_cross_site_write_protection(self):
        did = self.create(self.team)
        response = self.client.patch('/aichat/api/dialogs/' + did + '/team',
            json={'team': self.custom}, headers={'Origin': 'https://example.invalid'})
        self.assertEqual(response.status_code, 403)
        self.assertEqual(self.dialog(did)['team'], self.team)


if __name__ == '__main__':
    unittest.main()
