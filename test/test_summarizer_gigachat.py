"""A legacy GigaChat selection must use Qwen and never contact GigaChat."""
import json
from unittest.mock import patch

import httpx
import pytest

from services.summarizer import engine

TEXT = 'Синтетическое оперативное сообщение об аварии на водопроводе. ' * 2
QWEN_URL = 'https://aiplatform.mosreg.ru/api/user-models/v1/chat/completions'


@pytest.fixture(autouse=True)
def configured_qwen_without_real_network(monkeypatch):
    for name in ('SUMMARIZER_BACKEND', 'SUMMARIZER_QWEN_API_KEY', 'AI_CA_BUNDLE'):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv('QWEN_API_BASE', 'https://aiplatform.mosreg.ru/api/user-models/v1')
    monkeypatch.setenv('QWEN_API_KEY', 'synthetic-qwen-key')
    monkeypatch.setenv('QWEN_MODEL', 'synthetic-qwen-model')
    # Stale settings must not reactivate the removed provider.
    monkeypatch.setenv('SUMMARIZER_GIGACHAT_CREDENTIALS', 'unused-summary-giga-key')
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', 'unused-general-giga-key')
    monkeypatch.setenv('GIGACHAT_SCOPE', 'GIGACHAT_API_PERS')
    monkeypatch.setenv('GIGACHAT_MODE', 'GigaChat-2-Pro')
    def forbidden(*args, **kwargs):
        raise AssertionError('Real network is forbidden in this test')
    monkeypatch.setattr(httpx, 'post', forbidden)
    monkeypatch.setattr(httpx.Client, 'send', forbidden)


def successful_response(url=QWEN_URL):
    return httpx.Response(200, json={'choices': [{'message': {'content': 'Результат Госчата'}}]},
                          request=httpx.Request('POST', url))


@pytest.mark.parametrize('explicit,configured', [
    (None, None), (None, ''), (None, 'qwen'), (None, 'gigachat'),
    ('gigachat', None), (' GigaChat ', None), ('gigachat', 'algo'),
    ('qwen', 'gigachat'),
])
def test_default_and_retired_provider_use_the_same_qwen_transport(monkeypatch, explicit, configured):
    if configured is not None:
        monkeypatch.setenv('SUMMARIZER_BACKEND', configured)
    calls = []
    def post(url, **kwargs):
        assert url == QWEN_URL
        assert kwargs['headers']['Authorization'] == 'Bearer synthetic-qwen-key'
        assert kwargs['json']['model'] == 'synthetic-qwen-model'
        assert 'giga-key' not in json.dumps({'headers': kwargs['headers'], 'body': kwargs['json']})
        calls.append(url)
        return successful_response(url)
    monkeypatch.setattr(httpx, 'post', post)
    result = engine.summarize(TEXT, backend=explicit, creds='unused-legacy-explicit-giga-key')
    assert result['ok'] and result['backend'] == 'qwen'
    assert result['report_text'] == 'Результат Госчата'
    assert calls == [QWEN_URL]
    assert 'warning' not in result


@pytest.mark.parametrize('explicit,configured', [('algo', 'gigachat'), (None, 'algo')])
def test_local_algorithm_remains_local(monkeypatch, explicit, configured):
    monkeypatch.setenv('SUMMARIZER_BACKEND', configured)
    with patch.object(engine, '_qwen_chat') as qwen, patch.object(httpx, 'post') as network:
        result = engine.summarize(TEXT, backend=explicit)
    assert result['backend'] == 'algo'
    qwen.assert_not_called()
    network.assert_not_called()


def test_qwen_override_key_is_preserved_and_never_replaced_by_legacy_credentials(monkeypatch):
    monkeypatch.setenv('SUMMARIZER_QWEN_API_KEY', 'synthetic-summary-qwen-key')
    def post(url, **kwargs):
        assert url == QWEN_URL
        assert kwargs['headers']['Authorization'] == 'Bearer synthetic-summary-qwen-key'
        return successful_response(url)
    monkeypatch.setattr(httpx, 'post', post)
    assert engine.summarize(TEXT, backend='gigachat')['backend'] == 'qwen'


def test_failed_qwen_is_labeled_as_algorithm_without_trying_removed_provider(monkeypatch, caplog):
    calls = []
    def post(url, **kwargs):
        calls.append(url)
        assert url == QWEN_URL
        return httpx.Response(401, json={'error': 'PRIVATE_RESPONSE_BODY'}, request=httpx.Request('POST', url))
    monkeypatch.setattr(httpx, 'post', post)
    result = engine.summarize(TEXT, backend='gigachat')
    assert result['backend'] == 'algo'
    assert result['requested_backend'] == 'qwen'
    assert result['ai_error_code'] == 'AI_AUTH_FAILED'
    assert 'алгоритмом' in result['warning']
    assert calls == [QWEN_URL]
    assert 'PRIVATE_RESPONSE_BODY' not in caplog.text + json.dumps(result)
    assert 'synthetic-qwen-key' not in caplog.text + json.dumps(result)


def test_unknown_provider_does_not_trigger_network():
    with patch.object(engine, '_qwen_chat') as qwen, patch.object(httpx, 'post') as network:
        result = engine.summarize(TEXT, backend='unapproved-provider')
    assert not result['ok']
    qwen.assert_not_called()
    network.assert_not_called()


def test_regular_chat_and_tools_still_use_qwen(monkeypatch, tmp_path):
    from services.aichat import engine as chat_engine
    from services.tools import macros
    from services.tools.pptx_converter import parse_html_to_slides_ai
    with patch.object(chat_engine, '_qwen_chat', return_value='Обычный ответ') as qwen:
        assert chat_engine.ask([{'role': 'user', 'content': 'Привет'}]) == 'Обычный ответ'
    qwen.assert_called_once()
    macro = {'code': 'Option Explicit\nPublic Sub Example()\nMsgBox "Test"\nEnd Sub',
             'instructions': 'Import module', 'assumptions': ''}
    slide = [{'title': 'Synthetic slide', 'texts': ['Content']}]
    with patch.object(engine, '_qwen_chat', side_effect=[json.dumps(macro), json.dumps(slide)]) as qwen:
        assert macros.generate('Создай простой тестовый макрос.', 'Excel VBA', tmp_path)['file'] == 'macro.bas'
        assert parse_html_to_slides_ai('<h1>Synthetic slide</h1>')[0]['title'] == 'Synthetic slide'
    assert qwen.call_count == 2
