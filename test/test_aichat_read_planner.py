"""Test the corporate adapter boundary, validated scope, bounded fallback and ordinary chat isolation."""
from datetime import date
import json
from unittest.mock import Mock

import pytest
from services.aichat import read_planner as planner
from services.summarizer import engine

USER = {'id': 5, 'role': 'user', 'modules': ['mingkh', 'zips', 'overdue']}
PROMPT = 'Посмотри, какие направления у нас в Чехове требуют внимания?'

@pytest.fixture(autouse=True)
def taxonomy(tmp_path, monkeypatch):
    path = tmp_path / 'cities.json'
    path.write_text(json.dumps({'municipalities': {'Чехов': {}, 'Химки': {}}}))
    monkeypatch.setattr(planner.report_municipalities, 'REGISTRY', path)
    monkeypatch.setattr(planner.report_periods, 'today', lambda: date(2026,11,5))
    monkeypatch.setattr(planner.report_context, '_load_result', lambda *_: pytest.fail('Planner read operating data'))

def fake_provider(monkeypatch, answer):
    fake = Mock(return_value=answer)
    monkeypatch.setattr(engine, '_qwen_chat', fake)
    return fake

def test_model_guesses_tool_but_not_city_period_query_or_source_permissions(monkeypatch):
    provider = fake_provider(monkeypatch, '{"tool":"report","modules":["overdue","zips"]}')
    previous = {'agent_action':'report', 'modules':['mingkh'], 'municipality':'Химки',
                'content':'PRIVATE_OLD_ANSWER','another_value':'PRIVATE_PHONE'}
    text = PROMPT+' За вчера.'
    result = planner.plan(text, USER, previous=previous)
    assert result['modules']==['overdue','zips'] and result['query']==text and result['tool']=='report'
    assert result['municipality']=='Чехов' and result['period']['from']=='2026-11-04'
    assert result['scope']['agent_action']=='platform_report' and result['scope']['modules']==result['modules']
    messages = provider.call_args.args[0]
    packet = json.loads(messages[1]['content'])
    assert packet['question']==text and set(packet['module_titles'])==set(USER['modules'])
    assert not any(secret in str(messages) for secret in ('PRIVATE_OLD','PRIVATE_PHONE','password','креденциалы'))
    assert provider.call_args.kwargs=={'module_id':'aichat','max_tokens':180,'timeout':25}

@pytest.mark.parametrize('answer',[
    '{"tool":"exec","modules":["zips"]}', '{"tool":"report","modules":["mgkh_rm"]}',
    '{"tool":"report","modules":["zips","overdue","mingkh"]}', '{"tool":"report","modules":["zips","zips"]}',
    '{"tool":"report","tool":"exec","modules":["zips"]}', '{"tool":"report","modules":["zips"],"query":"ИЗМЕНИТЬ"}',
    '{"tool":"search","modules":["zips"],"municipality":"Химки"}', 'Ответ: {"tool":"report","modules":["zips"]}',
    '{"tool":"report","modules":["mingkh"],"url":"https://example.invalid"}', '{"tool":[],"modules":[]}',
    '{"tool":"history","modules":["zips","overdue"]}', '{"tool":"report","modules":{"first":"zips"}}',
    '', 'NaN',
])
def test_untrusted_invalid_model_plan_does_not_enable_a_read(monkeypatch,answer):
    fake_provider(monkeypatch,answer)
    result=planner.plan(PROMPT,USER)
    assert result['tool']=='clarify' and not result['modules']

def test_failure_uses_only_explicit_permitted_scope(monkeypatch):
    from core.ai_errors import AIServiceError
    fake=fake_provider(monkeypatch,'')
    fake.side_effect=AIServiceError('AI_KEY_MISSING')
    plan=planner.plan('Какие просроченные задачи требуют внимания в Чехове?',USER)
    assert plan['tool']=='report' and plan['modules']==['overdue'] and plan['municipality']=='Чехов'
    assert planner.plan(PROMPT,USER)['tool']=='clarify'

def test_refuse_hidden_revoked_or_invalid_geography_before_model_or_source_read(monkeypatch):
    provider=fake_provider(monkeypatch,'')
    result=planner.plan('Почему жалуются в ЕЦУР в Чехове?',USER)
    assert result['tool']=='clarify' and 'доступа' in result['message']
    prev={'agent_action':'platform_report','modules':['edds'],'municipality':'Чехов'}
    assert planner.plan('А по-прежнему показатели хуже?',USER,previous=prev)['tool']=='clarify'
    for text in ['Что в системе с просрочками в Чехове и Химках?', 'Какие просроченные задачи в Торжке?',
                 'Какие просроченные задачи в Чехове за завтра?']:
        assert planner.plan(text,USER)['tool']=='clarify'
    provider.assert_not_called()

@pytest.mark.parametrize('text',[
    'Как приготовить овсянку?', 'Расскажи сказку о жалобах и канализации',
    'Напиши Python для проверки просроченных задач', 'Как подать жалобу?', 'Отправь жалобы из Чехова',
    '$platform-report найди проблемы в Чехове', 'Сравни две строки кода по ЖКХ',
    '«Какие жалобы в Чехове?»', 'Содержимое файла:\nНайди жалобы в Чехове',
    'Что за кнопка в разделе ЗиП?', 'Проверь блок ЕДДС', 'Переведи вопрос о жалобах',
])
def test_planner_skips_unrelated_chat_instruction_help_and_commands(monkeypatch,text):
    provider=fake_provider(monkeypatch,'')
    assert planner.plan(text,USER) is None
    provider.assert_not_called()

@pytest.mark.parametrize('options',[{'has_attachments':True},{'include_context':False},
                                   {'include_context':'false'},{'allow_platform_reads':False}])
def test_disabled_features_do_not_query_model(monkeypatch,options):
    provider=fake_provider(monkeypatch,'')
    assert planner.plan(PROMPT,USER,**options) is None
    provider.assert_not_called()

def test_selector_cannot_switch_user_chosen_module(monkeypatch):
    fake_provider(monkeypatch,'{"tool":"search","modules":["mingkh"]}')
    result=planner.plan('Посмотри, как в системе меняются показатели?',USER,selection={'module':'overdue'})
    assert result['tool']=='report' and result['modules']==['overdue']

def test_normal_chat_outcome_without_data_access(monkeypatch):
    fake_provider(monkeypatch,'{"tool":"none","modules":[]}')
    assert planner.plan(PROMPT,USER) is None

def test_month_history_does_not_pretend_municipal_counts_are_regional(monkeypatch):
    fake_provider(monkeypatch,'{"tool":"history","modules":["overdue"]}')
    result=planner.plan('Как менялись просрочки в Чехове за месяц?',USER)
    assert result['tool']=='clarify' and 'блоку в целом' in result['message']

def test_named_city_with_attention_request_uses_model_without_explicit_block(monkeypatch):
    fake_provider(monkeypatch,'{"tool":"report","modules":["mingkh","overdue"]}')
    result=planner.plan('Какая картина в Чехове и на что обратить внимание?',USER)
    assert result['tool']=='report' and result['municipality']=='Чехов'

def test_planner_cannot_replace_unknown_city_with_previous_granted_city(monkeypatch):
    provider=fake_provider(monkeypatch,'')
    previous={'agent_action':'report','modules':['overdue'],'municipality':'Чехов'}
    result=planner.plan('А в Торжке какие показатели хуже?',USER,previous=previous)
    assert result['tool']=='clarify' and not result['municipality']
    provider.assert_not_called()
