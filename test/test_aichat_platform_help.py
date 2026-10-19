"""Platform help/report orchestration: no live model or portal calls."""
from pathlib import Path
from unittest.mock import patch
import unittest

from services.aichat import engine, platform_guide, storage
from test import test_aichat_skills as skills_fixture


def test_help_is_scoped_and_not_a_metrics_snapshot():
    question = 'Как сохранить доклад в коллективных жалобах?'
    guide = platform_guide.build(question, ['collective'])
    assert 'Сохранить доклад' in guide
    assert '/mingkh/collective' in guide
    assert '/users' not in guide
    assert '/zips' not in guide
    assert '/mingkh/sms' not in guide
    assert 'администратор управляет справочником' not in guide
    admin = platform_guide.build(question, ['collective'], is_manager=True)
    assert '/users' in admin and 'кнопку «Ответственные»' in admin
    denied = platform_guide.build(question, ['edo'])
    assert '/mingkh/collective' not in denied and 'Сохранить доклад' not in denied


def test_ordinary_chat_and_mixed_request_are_distinct():
    for text in ['Напиши поздравление', 'Что такое давление воды?', 'Расскажи про Qwen']:
        assert platform_guide.build(text, ['collective']) == ''
    assert platform_guide.help_only('Как сохранить доклад в коллективных жалобах?')
    assert platform_guide.help_only('Дай инструкцию, как пользоваться блоком МинЖКХ')
    assert platform_guide.help_only('Как посмотреть статистику в Нейроне?')
    assert not platform_guide.help_only('Как изменились показатели ЕДДС за неделю?')
    assert not platform_guide.help_only('Покажи статистику коллективных жалоб и объясни, как сохранить доклад')


def test_old_large_turns_cannot_displace_current_question_or_evidence(monkeypatch):
    captured = []
    monkeypatch.setattr(engine, '_qwen_chat', lambda messages, **kwargs: captured.append(messages) or 'Ответ')
    current = 'Текущий запрос ' + 'я' * 20000
    history = [{'role': 'user', 'content': 'Старый запрос'},
               {'role': 'assistant', 'content': 'Старый огромный ответ' + 'x' * 30000},
               {'role': 'user', 'content': 'Недавний вопрос'},
               {'role': 'assistant', 'content': 'Недавний ответ'},
               {'role': 'user', 'content': current}]
    engine.ask(history, platform_context='{"checked":true}', platform_help='SAFE_UI_GUIDE')
    messages = captured[0]
    assert 'SAFE_UI_GUIDE' in messages[0]['content']
    assert len(messages) == 4
    assert messages[1]['content'] == 'Недавний вопрос'
    assert messages[-1]['content'].startswith(current)
    assert '<platform_data>' in messages[-1]['content']
    assert history[-1]['content'] == current
    assert 'Старый огромный ответ' not in str(messages)


def test_report_period_storage_is_validated_and_survives_reload(tmp_path, monkeypatch):
    monkeypatch.setattr(storage, 'DATA_DIR', tmp_path)
    monkeypatch.setattr(storage, 'DIALOGS_FILE', tmp_path / 'dialogs.json')
    dialog = storage.create_dialog()
    period = {'from': '2026-10-18', 'to': '2026-10-19', 'label': 'Два дня', 'timezone': 'unsafe', 'extra': 'discard'}
    storage.append_message(dialog['id'], 'user', 'Отчёт', report_scope={'requested_period': period})
    saved = storage.get_dialog(dialog['id'])['messages'][-1]['report_scope']['requested_period']
    assert saved == {'from': '2026-10-18', 'to': '2026-10-19', 'label': 'Два дня', 'timezone': 'Europe/Moscow'}
    for invalid in [{'from': '../wrong', 'to': 'today'}, {'from': '2026-10-20', 'to': '2026-10-19'}, {'from': 1}]:
        storage.append_message(dialog['id'], 'user', 'Запрос', report_scope={'requested_period': invalid})
        assert 'requested_period' not in storage.get_dialog(dialog['id'])['messages'][-1]['report_scope']


class PlatformOrchestrationTests(unittest.TestCase):
    setUpClass = classmethod(skills_fixture.SkillAPITests.setUpClass.__func__)
    tearDownClass = classmethod(skills_fixture.SkillAPITests.tearDownClass.__func__)
    setUp = skills_fixture.SkillAPITests.setUp
    tearDown = skills_fixture.SkillAPITests.tearDown
    login = skills_fixture.SkillAPITests.login
    send = skills_fixture.SkillAPITests.send
    events = skills_fixture.SkillAPITests.events

    def test_auto_report_does_not_depend_on_model_selecting_tool(self):
        from services.aichat.skills import runtime
        bundle = {'selection': {'modules': ['edo']}, 'sources': [], 'clarification': ''}
        with patch.object(runtime, '_model', return_value='{"skills":[]}'), \
                patch.object(self.module, 'build_platform_context', return_value=bundle) as read, \
                patch.object(self.module, 'aichat_ask', return_value='Отчёт') as answer:
            result = self.events(self.send(text='Дай отчёт по заполненности данных'))[-1]
        self.assertTrue(result['ok'])
        read.assert_called_once()
        self.assertIn('edo', answer.call_args.kwargs['platform_context'])

    def test_auto_report_respects_disabled_skill(self):
        from services.aichat.skills import runtime
        response = self.client.patch('/aichat/api/skills/platform-report', json={'enabled': False})
        self.assertEqual(response.status_code, 200)
        with patch.object(runtime, '_model', return_value='{"skills":[]}'), \
                patch.object(self.module, 'build_platform_context') as read, \
                patch.object(self.module, 'aichat_ask', return_value='Нет данных') as answer:
            self.events(self.send(text='Дай отчёт по заполненности данных'))
        read.assert_not_called()
        self.assertEqual(answer.call_args.kwargs['platform_context'], '')

    def test_help_uses_server_guide_even_if_model_tries_report_tool(self):
        from services.aichat.skills import runtime
        plan = ['{"skills":["platform-report"]}',
                '{"actions":[{"skill":"platform-report","tool":"platform_report","args":{}}]}']
        with patch.object(runtime, '_model', side_effect=plan), \
                patch.object(self.module, 'build_platform_context') as read, \
                patch.object(self.module, 'aichat_ask', return_value='Откройте «Оформление»') as answer:
            result = self.events(self.send(text='Где в Нейроне найти оформление?'))[-1]
        self.assertEqual(result['response_kind'], 'answer')
        read.assert_not_called()
        self.assertEqual(answer.call_args.kwargs['platform_context'], '')
        self.assertIn('Оформление', answer.call_args.kwargs['platform_help'])
        self.assertNotIn('/users', answer.call_args.kwargs['platform_help'])

    def test_skill_catalog_failure_does_not_break_ordinary_model_answer(self):
        from services.aichat.skills import preferences
        with patch.object(preferences, 'enabled_skills', side_effect=RuntimeError('private-path')), \
                patch.object(self.module, 'build_platform_context') as read, \
                patch.object(self.module, 'aichat_ask', return_value='Инструмент временно недоступен') as answer:
            result = self.events(self.send(text='Дай отчёт по заполненности данных'))[-1]
        self.assertTrue(result['ok'])
        read.assert_not_called()
        self.assertIn('Список навыков недоступен', answer.call_args.kwargs['tool_evidence'])
        self.assertNotIn('private-path', str(result))
