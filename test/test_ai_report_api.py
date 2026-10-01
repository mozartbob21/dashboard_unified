"""Test the real chat endpoint with local files and a mocked AI transport."""
import importlib
import json
from unittest.mock import patch

import pytest
from fastapi.testclient import TestClient


@pytest.fixture
def chat_client(tmp_path, monkeypatch):
    import utils.db as db
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'reports-test.sqlite')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr('services.auth.activity.track_response', lambda *args: None)
    with patch('services.scheduler.start'):
        app_module = importlib.import_module('app')
    user = {'username': 'report-tester', 'role': 'user', 'modules': ['water-dashboard']}
    monkeypatch.setattr(app_module, 'get_user_from_token', lambda *args: user)
    from services.auth import security
    monkeypatch.setattr(security, 'get_user_from_token', lambda *args: user)
    from services.aichat import storage, report_context as reports
    from services import water_ai_context as water
    monkeypatch.setattr(storage, 'DATA_DIR', tmp_path / 'chat')
    monkeypatch.setattr(storage, 'DIALOGS_FILE', tmp_path / 'chat/dialogs.json')
    monkeypatch.setattr(reports, 'DATA_DIR', tmp_path / 'results')
    monkeypatch.setattr(water, 'load_snapshot', lambda: {'schema_version': 2, 'sources': {'tasks': {
        'metric_schema': 1, 'ok': True, 'updated_at': '2026-10-01T20:00:00',
        'data_date': '2026-09-30', 'metrics': [{'id': 'tasks', 'label': 'Просроченные задачи', 'value': 12, 'unit': ''}],
    }}, 'table': [{'name': 'Балашиха', 'tasks': 12}, {'name': 'Власиха', 'tasks': 3}, {'name': 'Химки', 'tasks': 9}]})
    client = TestClient(app_module.app)
    client.cookies.set('access_token', 'local-test-token')
    try:
        yield client, app_module, user
    finally:
        client.close()


def test_chat_endpoint_supplies_real_filtered_context_and_preserves_scope(chat_client):
    client, app, user = chat_client
    with patch.object(app, 'aichat_ask', return_value='Проверенный тестовый ответ') as ai:
        reply = client.post('/aichat/api/send', data={'text': 'РМ МИНЖКХ по Балашихе', 'include_context': 'true'})
    assert reply.status_code == 200
    result = reply.json()
    assert result['report_scope']['municipality'] == 'Балашиха'
    ctx = json.loads(ai.call_args.kwargs['platform_context'])
    assert ctx['sources'][0]['rows'][0]['metrics'][0]['value'] == 12
    assert ctx['sources'][0]['metrics'] == []
    assert ctx['sources'][0]['collected_at'] == '2026-10-01T20:00:00'
    dialog = client.get('/aichat/api/dialogs/' + result['dialog_id']).json()
    assert dialog['messages'][0]['report_scope']['module'] == 'water-dashboard'
    assert 'platform_data' not in dialog['messages'][0]['content']
    # History stays shared; revoke the source grant before the follow-up.
    user['modules'] = []
    assert client.get('/aichat/api/dialogs/' + result['dialog_id']).status_code == 200
    with patch.object(app, 'aichat_ask') as ai:
        followup = client.post('/aichat/api/send', data={'dialog_id': result['dialog_id'], 'text': 'Подробнее', 'include_context': 'true'}).json()
    ai.assert_not_called()
    assert not followup['report_sources']
    assert 'недоступ' in followup['answer']


def test_ordinary_chat_does_not_load_context(chat_client):
    client, app, _ = chat_client
    with patch.object(app, 'aichat_ask', return_value='Ответ') as ai, patch.object(app, 'build_platform_context') as prepare:
        assert client.post('/aichat/api/send', data={'text': 'Привет'}).status_code == 200
    prepare.assert_not_called()
    assert ai.call_args.kwargs['platform_context'] == ''


def test_unavailable_ai_returns_safe_saved_numbers(chat_client):
    client, app, _ = chat_client
    with patch.object(app, 'aichat_ask', side_effect=RuntimeError('SECRET_TOKEN=forbidden')):
        data = client.post('/aichat/api/send', data={'text': 'РМ МИНЖКХ', 'include_context': 'true'}).json()
    assert '12' in data['answer'] and '2026-10-01' in data['answer']
    assert 'SECRET_TOKEN' not in data['answer']


def test_options_are_permission_filtered(chat_client):
    client, _, user = chat_client
    options = client.get('/aichat/api/report-options').json()
    assert [m['id'] for m in options['modules']] == ['water-dashboard']
    assert 'Балашиха' in options['municipalities']
    user['modules'] = []
    assert client.get('/aichat/api/report-options').json() == {'modules': [], 'sources': [], 'municipalities': []}


@pytest.mark.parametrize('question,modules', [
    ('Дай краткий отчёт по всем блокам', ['edds', 'mingkh', 'water-dashboard']),
    ('Дай отчёт по Власихе', ['edds', 'water-dashboard']),
    ('Покажи камеры и просрочку по Химкам', ['cameras', 'overdue']),
    ('Что в критичных по технадзору?', ['utnkr']),
    ('Дай отчёт по каждому муниципалитету в блоках ЕДДС и МинЖКХ', ['edds', 'mingkh']),
])
def test_natural_requests_need_no_form_flags(chat_client, question, modules):
    client, app, user = chat_client
    # Keep a permitted saved municipality catalogue while testing other blocks.
    user['modules'] = modules + ['water-dashboard']
    with patch.object(app, 'aichat_ask', return_value='Тестовый отчёт') as ai:
        reply = client.post('/aichat/api/send', data={'text': question})
    assert reply.status_code == 200
    data = reply.json()
    assert 'report_sources' in data
    ctx = json.loads(ai.call_args.kwargs['platform_context'])
    expected = set(user['modules']) if 'всем блокам' in question or 'по Власихе' in question else set(modules)
    assert {s['module'] for s in ctx['sources']} == expected
    if 'каждому муниципалитету' in question:
        assert ctx['selection']['group_by'] == 'municipality'


def test_multiple_scope_persists_and_new_city_clears_old_block(chat_client):
    client, app, user = chat_client
    user['modules'] = ['water-dashboard', 'edds', 'mingkh']
    with patch.object(app, 'aichat_ask', return_value='Тест') as ai:
        first = client.post('/aichat/api/send', data={'text': 'ЕДДС и МинЖКХ по Химкам и Власихе'}).json()
        assert set(first['report_scope']['modules']) == {'edds', 'mingkh'}
        assert set(first['report_scope']['municipalities']) == {'Химки', 'Власиха'}
        follow = client.post('/aichat/api/send', data={'dialog_id': first['dialog_id'], 'text': 'Подробнее'}).json()
        assert follow['report_scope'] == first['report_scope']
        fresh = client.post('/aichat/api/send', data={'dialog_id': first['dialog_id'], 'text': 'Дай отчёт по Балашихе'}).json()
        assert fresh['report_scope']['modules'] == []
        assert fresh['report_scope']['municipalities'] == ['Балашиха']
        assert {s['module'] for s in json.loads(ai.call_args.kwargs['platform_context'])['sources']} == set(user['modules'])
    messages = client.get('/aichat/api/dialogs/' + first['dialog_id']).json()['messages']
    assert messages[0]['report_scope']['modules'] == ['edds', 'mingkh']


def test_explicit_opt_out_legacy_and_file_analysis_still_work(chat_client):
    client, app, _ = chat_client
    with patch.object(app, 'aichat_ask', return_value='Ответ') as ai, patch.object(app, 'build_platform_context') as prepare:
        assert client.post('/aichat/api/send', data={'text': 'Дай отчёт по всем блокам', 'include_context': 'false'}).status_code == 200
        assert ai.call_args.kwargs['platform_context'] == ''
        assert client.post('/aichat/api/send', data={'text': 'Сделай отчёт по файлу'}, files={'file': ('text.txt', b'Local file example', 'text/plain')}).status_code == 200
        assert ai.call_args.kwargs['platform_context'] == ''
    prepare.assert_not_called()


def test_unrelated_turn_stops_scope_inheritance(chat_client):
    client, app, _ = chat_client
    with patch.object(app, 'aichat_ask', return_value='Ответ') as ai:
        first = client.post('/aichat/api/send', data={'text': 'РМ МИНЖКХ по Балашихе'}).json()
        client.post('/aichat/api/send', data={'dialog_id': first['dialog_id'], 'text': 'Спасибо'})
        client.post('/aichat/api/send', data={'dialog_id': first['dialog_id'], 'text': 'Подробнее'})
    assert ai.call_args.kwargs['platform_context'] == ''



def test_remaining_modules_are_automatic_and_permission_checked(chat_client):
    client, app, user = chat_client
    from services.aichat import remaining_sources
    user['modules'] = ['ecur', 'water_rm']
    question = 'Дай отчёт по ЕЦУР и РМ Водоснабжение'
    with patch.object(app, 'aichat_ask', return_value='Тестовый отчёт') as ai:
        response = client.post('/aichat/api/send', data={'text': question})
    assert response.status_code == 200
    context = json.loads(ai.call_args.kwargs['platform_context'])
    assert set(context['selection']['modules']) == {'ecur', 'water_rm'}
    assert {s['module'] for s in context['sources']} == {'ecur', 'water_rm'}
    assert all(s['status'] == 'missing' for s in context['sources'])
    user['modules'] = ['water_rm']
    with patch.object(app, 'aichat_ask', return_value='Тестовый отчёт') as ai, patch.object(remaining_sources, 'reports_for_module', wraps=remaining_sources.reports_for_module) as read:
        client.post('/aichat/api/send', data={'text': question})
    assert {call.args[0] for call in read.call_args_list} == {'water_rm'}
    assert {s['module'] for s in json.loads(ai.call_args.kwargs['platform_context'])['sources']} == {'water_rm'}



def test_remaining_modules_use_real_persisted_safe_city_data(chat_client):
    client, app, user = chat_client
    from services.aichat import remaining_sources as remaining, report_context as reports
    directory = reports.DATA_DIR
    directory.mkdir(parents=True, exist_ok=True)
    (directory / 'municipality_registry.json').write_text(json.dumps({'municipalities': {'Власиха': {}, 'Химки': {}, 'Балашиха': {}}}))
    assert remaining.persist_ecur_snapshot([
        ['Район', 'Категория ЕЦУР', 'Срок', 'Статус', 'Описание'],
        ['Власиха', 'Качество воды', '01.12.2026', 'Назначено', 'SECRET_COMPLAINT'],
        ['Химки', 'Нет воды', '01.12.2026', 'Открыто', 'PRIVATE_TEXT'],
        ['Балашиха', 'UNSELECTED_CATEGORY', '01.12.2026', 'UNSELECTED_STATUS', 'PRIVATE_TEXT'],
    ], data_dir=directory)
    remaining.save_water_check({'rez': {'close': [{'id': 1, 'project': 'Власиха', 'subject': 'SECRET_TASK'}], 'rw': []},
                                'sys': {'close': [], 'ext': [{'id': 2, 'project': 'Химки'}], 'rw': []}, 'total_tasks': 2}, data_dir=directory)
    user['modules'] = ['ecur', 'water_rm']
    with patch.object(app, 'aichat_ask', return_value='Тест') as ai:
        response = client.post('/aichat/api/send', data={'text': 'Дай отчёт по Власихе в ЕЦУР и РМ Водоснабжение'}).json()
        context = json.loads(ai.call_args.kwargs['platform_context'])
        assert response['report_scope']['municipalities'] == ['Власиха']
        assert len(context['sources']) == 2
        assert all(s['metrics'][0]['value'] == 1 for s in context['sources'])
        assert all(s['collected_at'] for s in context['sources'])
        assert 'SECRET_' not in ai.call_args.kwargs['platform_context']
        assert 'PRIVATE_TEXT' not in ai.call_args.kwargs['platform_context']
        client.post('/aichat/api/send', data={'text': 'Дай отчёт по Власихе и Химкам в ЕЦУР и РМ Водоснабжение'})
        grouped = json.loads(ai.call_args.kwargs['platform_context'])
    assert all('category_counts' not in s and 'status_counts' not in s for s in grouped['sources'])
    assert 'UNSELECTED_' not in json.dumps(grouped)
    assert 'Категория: Качество воды' in grouped['sources'][0]['municipality_table']['columns']
