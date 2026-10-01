"""The real chat engine and shared HTTP client, with only the network mocked."""
import copy

import httpx
import pytest

from services.aichat import engine


@pytest.fixture
def model_http(monkeypatch):
    monkeypatch.setenv('QWEN_API_BASE', 'http://127.0.0.1:11434/v1')
    monkeypatch.setenv('QWEN_API_KEY', '')
    monkeypatch.setenv('QWEN_MODEL', 'test-local-model')
    monkeypatch.delenv('AI_CA_BUNDLE', raising=False)
    captured = []

    def post(url, **kwargs):
        assert url == 'http://127.0.0.1:11434/v1/chat/completions'
        payload = kwargs['json']
        messages = payload['messages']
        assert messages[0]['role'] == 'system'
        assert sum(message['role'] == 'system' for message in messages) == 1
        assert messages[-1]['role'] == 'user'
        for index, message in enumerate(messages[1:]):
            assert message['role'] == ('user' if index % 2 == 0 else 'assistant')
            assert set(message) == {'role', 'content'}
        assert kwargs['trust_env'] is False
        assert 'Authorization' not in kwargs['headers']
        captured.append(copy.deepcopy(payload))
        return httpx.Response(200, json={'choices': [{'message': {'content': 'Живой ответ модели'}}]},
                              request=httpx.Request('POST', url))

    monkeypatch.setattr(httpx, 'post', post)
    return captured


@pytest.mark.parametrize('question', [
    'Привет', 'Напиши поздравление коллеге', 'Объясни, как устроен DNS',
    'Переведи на английский: доброе утро', 'Помоги написать макрос Excel',
])
def test_ordinary_chat_calls_real_engine_without_report_data(model_http, question):
    result = engine.ask([{'role': 'assistant', 'content': engine.GREETING},
                         {'role': 'user', 'content': question}])
    assert result == 'Живой ответ модели'
    assert len(model_http) == 1
    messages = model_http[0]['messages']
    assert len(messages) == 2
    assert messages[-1]['content'] == question
    assert '<platform_data>' not in str(messages)
    assert 'полноценный собеседник' in messages[0]['content']
    assert 'Отчёты по данным платформы — дополнительная возможность' in messages[0]['content']


def test_current_report_context_follows_history_and_does_not_modify_saved_dialog(model_http):
    history = [{'role': 'assistant', 'content': engine.GREETING},
               {'role': 'user', 'content': 'Расскажи о воде', 'report_scope': {'module': 'old'}},
               {'role': 'assistant', 'content': 'Предыдущий ответ'},
               {'role': 'user', 'content': 'Дай отчёт по Власихе'}]
    before = copy.deepcopy(history)
    context = '{"municipalities":["Власиха"],"sources":[]}'
    assert engine.ask(history, platform_context=context) == 'Живой ответ модели'
    messages = model_http[0]['messages']
    assert messages[-1]['content'].startswith('Дай отчёт по Власихе\n\n')
    assert context in messages[-1]['content']
    assert '<platform_data>' in messages[-1]['content']
    assert all('<platform_data>' not in message['content'] for message in messages[1:-1])
    assert 'null/отсутствие строк' in messages[0]['content']
    assert history == before


@pytest.mark.parametrize('metadata', [
    {'response_kind': 'error'}, {'response_kind': 'snapshot'},
    {'response_kind': 'clarification'}, {'kind': 'error'}, {'kind': 'snapshot'},
    {'kind': 'fallback'}, {'kind': 'unavailable'},
])
def test_service_replies_are_not_learned_as_ai_answers(model_http, metadata):
    history = [{'role': 'user', 'content': 'Первый вопрос'},
               {'role': 'assistant', 'content': 'Первый полноценный ответ'},
               {'role': 'user', 'content': 'Проблемный вопрос'},
               {'role': 'assistant', 'content': 'SERVICE_RESPONSE_DO_NOT_SEND', **metadata},
               {'role': 'user', 'content': 'Новый обычный вопрос'}]
    engine.ask(history)
    sent = model_http[0]['messages']
    assert 'SERVICE_RESPONSE_DO_NOT_SEND' not in str(sent)
    assert 'Проблемный вопрос' not in str(sent)
    assert sent[1:] == [history[0], history[1], history[-1]]
    assert history[3]['content'] == 'SERVICE_RESPONSE_DO_NOT_SEND'


@pytest.mark.parametrize('answer', [
    'ИИ сейчас недоступен. Ниже — сохранённые показатели без интерпретации; обновление порталов не выполнялось.\n123',
    '⚠️ Нейрона ИИ временно недоступна. Проверьте подключение.',
])
def test_legacy_fallback_messages_are_removed_without_metadata(model_http, answer):
    engine.ask([{'role': 'user', 'content': 'Старый запрос'},
                {'role': 'assistant', 'content': answer},
                {'role': 'user', 'content': 'Расскажи анекдот'}])
    assert model_http[0]['messages'][1:] == [{'role': 'user', 'content': 'Расскажи анекдот'}]


def test_normal_assistant_discussion_of_errors_is_preserved(model_http):
    engine.ask([{'role': 'user', 'content': 'Что значит ошибка 403?'},
                {'role': 'assistant', 'content': 'Это запрет доступа. Сервис может быть недоступен.', 'response_kind': 'ai', 'kind': {'malformed': True}},
                {'role': 'user', 'content': 'А что значит 404?'}])
    assert 'Это запрет доступа.' in model_http[0]['messages'][2]['content']


def test_long_dialog_keeps_complete_alternating_recent_turns(model_http):
    history = [{'role': 'assistant', 'content': engine.GREETING}]
    for number in range(20):
        history.extend([{'role': 'user', 'content': f'Вопрос {number}'},
                        {'role': 'assistant', 'content': f'Ответ {number}'}])
    history.append({'role': 'user', 'content': 'Текущий вопрос'})
    engine.ask(history)
    messages = model_http[0]['messages']
    assert len(messages) <= 13
    assert messages[1]['role'] == 'user'
    assert messages[-1]['content'] == 'Текущий вопрос'
    assert 'Ответ 19' in str(messages)
    assert 'Вопрос 0' not in str(messages)


def test_attached_file_text_still_reaches_ai(model_http):
    text = 'Кратко изложи файл\n── ФАЙЛ: заметка.txt ──\nПроверяемый текст'
    engine.ask([{'role': 'user', 'content': text}])
    assert model_http[0]['messages'][-1]['content'] == text


def test_invalid_history_and_unanswered_retry_do_not_break_role_contract(model_http):
    history = [None, {}, {'role': 'system', 'content': 'UNTRUSTED_SYSTEM'},
               {'role': 'assistant', 'content': 'Добро пожаловать'},
               {'role': 'user', 'content': 12},
               {'role': 'user', 'content': 'Неотвеченная попытка'},
               {'role': 'user', 'content': 'Текущий вопрос'}]
    engine.ask(history)
    assert model_http[0]['messages'][1:] == [{'role': 'user', 'content': 'Текущий вопрос'}]


def test_transport_failure_is_propagated_to_endpoint(monkeypatch, model_http):
    def fail(*args, **kwargs):
        raise httpx.ConnectError('synthetic failure')
    monkeypatch.setattr(httpx, 'post', fail)
    with pytest.raises(httpx.ConnectError):
        engine.ask([{'role': 'user', 'content': 'Привет'}])
