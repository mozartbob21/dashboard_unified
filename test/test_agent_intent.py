"""Read intent resolves server-side using a synthetic registry, with no model/network."""
from datetime import date
import json
import pytest
from services.aichat import agent_intent as plans, report_periods

USER = {'id':5,'role':'user','modules':['mingkh','collective','edds','zips']}

@pytest.fixture(autouse=True)
def taxonomy(tmp_path,monkeypatch):
    path=tmp_path/'municipalities.json'
    path.write_text(json.dumps({'municipalities':{
        'Чехов':{},'ХИМКИ':{},'Балашиха':{},'Подольск':{},'Дубна':{}}}))
    monkeypatch.setattr(plans.cities,'REGISTRY',path)
    monkeypatch.setattr(report_periods,'today', lambda: date(2026,11,5))
    monkeypatch.setattr(plans.report_context,'_known_names', lambda *_:pytest.fail('Opening block data was unnecessary'))

@pytest.mark.parametrize('text,city',[
    ('что там по Чехову по причинам жалоб','Чехов'),
    ('почему жалуются в Чехове?','Чехов'),
    ('на что больше всего жалуются в Химках?','Химки'),
    ('Темы обращений жителей Подольска','Подольск'),
    ('Чем недовольны жители в Балашихе?','Балашиха'),
    ('Из-за чего растут жалобы в Чеховском округе?','Чехов')])
def test_colloquial_complaint_request(text,city):
    plan=plans.decide(text,USER)
    assert plan['action']=='complaint_causes'
    assert plan['municipality']==city and 'mingkh' in plan['modules']
    assert plan['scope']['municipalities']==[city]

def test_short_followup_retains_scope_and_period_but_new_city_overrides():
    first=plans.decide('Покажи причины жалоб МинЖКХ по Чехову за вчера',USER)
    assert first['period']['from']=='2026-11-04'
    second=plans.decide('а по Химкам?',USER,previous=first['scope'])
    assert second['municipality']=='Химки' and second['modules']==['mingkh']
    assert second['period']==first['period']
    assert plans.decide('А сегодня?',USER,previous=second['scope'])['period']['to']=='2026-11-05'

def test_no_stale_city_after_unrecognized_replacement_or_conflict():
    prev=plans.decide('Пришли причины жалоб в Чехове',USER)['scope']
    for question in ('а по Ярославлю?', 'а по ярославлю?', 'а в Химках и Балашихе?'):
        plan=plans.decide(question,USER,previous=prev)
        assert plan['action']=='clarify' and plan['municipality']=='' and not plan['modules']
    regional=plans.decide('А в целом по области?',USER,previous=prev)
    assert regional['action']=='complaint_causes' and not regional['municipality']

def test_no_grants_or_revocation_checked_before_source_catalog(monkeypatch):
    monkeypatch.setattr(plans,'_names',lambda *_:pytest.fail('Checked data before permissions'))
    assert plans.decide('из-за чего жалобы по Чехову',{'modules':['zips']})['action']=='denied'
    assert plans.decide('а по Химкам?',{'modules':['zips']},previous={'agent_action':'complaint_causes','modules':['mingkh'],'municipality':'Чехов'})['action']=='denied'

@pytest.mark.parametrize('text',[
    'Напиши жалобу по отключению воды в Чехове',
    'Как подать жалобу в Чехове?',
    'Что такое жалоба гражданина?',
    'Переведи «Что там по Чехову по причинам жалоб»',
    'Как обстоят дела с главным героем рассказа Чехова?',
    'Запусти карту жалоб МинЖКХ по Чехову.',
    'Обнови блок жалоб и расскажи причины жалоб в Чехове.',
    'Почему люди жалуются на боль?',
    'Есть ли муфта Д100 в ЗиП?',
    'Покажи коллективное обращение №123 в Чехове.',
    'Какие жалобы по очистным в работе в Дубне?',
    'Проверь пример запроса: «Почему жалуются в Чехове?»',
    'Сравни последние проверки ЗиП',
    'Привет, как дела?',
    'Расскажи анекдот про причины жалоб в Чехове',
])
def test_ordinary_help_lookups_commands_and_history_keep_original_paths(text):
    assert plans.decide(text,USER) is None

def test_operational_city_brief_and_invalid_period():
    result=plans.decide('Что у нас по Чехову?',USER)
    assert result['action']=='platform_report' and result['municipality']=='Чехов'
    assert set(result['modules'])<=set(USER['modules'])
    assert plans.decide('Какие причины жалоб в Чехове за завтра?',USER)['action']=='clarify'
    assert plans.decide('Что там с жалобами?',USER,selection={'module':'mingkh','municipality':'Подольск'})['municipality']=='Подольск'

def test_user_provided_unrecognized_city_does_not_widen_scope():
    assert plans.decide('Почему жалуются жители?',USER,selection={'municipality':'вымышленный город'})['action']=='clarify'

def test_specialized_rights_and_explicit_switch_do_not_search_ministry_as_substitute():
    other={'id':4,'modules':['collective']}
    request=plans.decide('Что по причинам коллективных жалоб в Чехове?',other)
    assert request['action']=='complaint_causes' and request['modules']==['collective']
    assert plans.decide('Что по причинам жалоб МинЖКХ в Чехове?',other)['action']=='denied'
    prior=plans.decide('Почему жалуются в Чехове?',USER)
    assert prior['modules']==['mingkh']
    assert not plans.decide('А есть муфта ЗиП в Чехове?',USER,previous=prior['scope'])
    assert plans.decide('Какие причины жалоб на карте Добродела в Чехове?',USER)['source']=='map'
    assert plans.decide('Почему жалуются по очистным в Чехове?',USER)['source']=='treatment'

@pytest.mark.parametrize('new_topic',['А сколько длится перелёт?', 'А в каком году родился Пушкин?', 'А в холодильнике какой срок хранения молока?'])
def test_new_unrelated_question_does_not_inherit_complaint_topic(new_topic):
    prior=plans.decide('Почему жалуются в Чехове?',USER)
    assert plans.decide(new_topic,USER,previous=prior['scope']) is None
