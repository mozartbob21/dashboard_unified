"""Long final answers keep the approved provider and short connection limits."""
import httpx
import pytest

from core.ai_errors import describe_ai_error
from services.aichat import engine
from services.aichat.config import model_timeout
from services.summarizer.engine import _qwen_chat


@pytest.fixture
def provider(monkeypatch):
    monkeypatch.delenv('AICHAT_MODEL_READ_TIMEOUT_SECONDS', raising=False)
    monkeypatch.delenv('AI_CA_BUNDLE', raising=False)
    monkeypatch.setenv('QWEN_API_BASE', 'https://aiplatform.mosreg.ru/api/user-models/v1')
    monkeypatch.setenv('QWEN_API_KEY', 'synthetic-test-key')
    monkeypatch.setenv('QWEN_MODEL', 'configured-model')
    calls = []

    def post(url, **kwargs):
        calls.append({'url': url, **kwargs})
        timeout = httpx.Timeout(kwargs['timeout'])
        # Simulate a provider that finishes after four minutes of silence.
        # The former 180-second chat timeout could never receive this answer.
        if timeout.read <= 240:
            raise httpx.ReadTimeout('Simulated four-minute response')
        return httpx.Response(200, json={'choices': [{'message': {'content': 'Подробный ответ'}}]},
                              request=httpx.Request('POST', url))

    monkeypatch.setattr(httpx, 'post', post)
    return calls


def test_long_answer_keeps_provider_privacy_and_short_connection_limits(provider):
    assert engine.ask([{'role': 'user', 'content': 'Объясни подробно'}]) == 'Подробный ответ'
    call, = provider
    assert call['url'] == 'https://aiplatform.mosreg.ru/api/user-models/v1/chat/completions'
    assert call['headers']['Authorization'] == 'Bearer synthetic-test-key'
    assert call['json']['model'] == 'configured-model'
    assert call['json']['messages'][-1]['content'] == 'Объясни подробно'
    assert call['json']['stream'] is False
    assert call['verify'] is not False
    assert call['trust_env'] is False and call['follow_redirects'] is False
    assert call['timeout'].read == 900
    assert call['timeout'].connect == 15
    assert call['timeout'].write == 30
    assert call['timeout'].pool == 15


def test_chat_timeout_setting_does_not_extend_other_modules(provider, monkeypatch):
    monkeypatch.setenv('AICHAT_MODEL_READ_TIMEOUT_SECONDS', '1200')
    assert engine.ask([{'role': 'user', 'content': 'Длинный ответ'}]) == 'Подробный ответ'
    assert provider[-1]['timeout'].read == 1200
    with pytest.raises(httpx.ReadTimeout):
        _qwen_chat([{'role': 'user', 'content': 'Суммируй'}])
    assert provider[-1]['timeout'] == 180


@pytest.mark.parametrize('setting, expected', [
    ('600', 600), ('1200.5', 1200.5), ('1', 60), ('3600', 1800),
    ('invalid', 900), ('', 900), ('0', 900), ('-10', 900),
    ('nan', 900), ('inf', 900), ('-inf', 900),
])
def test_timeout_setting_is_finite_bounded_and_cannot_disable_deadlines(monkeypatch, setting, expected):
    monkeypatch.setenv('AICHAT_MODEL_READ_TIMEOUT_SECONDS', setting)
    timeout = model_timeout()
    assert timeout.read == expected
    assert timeout.connect == 15 and timeout.write == 30 and timeout.pool == 15


def test_eventual_timeout_is_still_an_error_without_automatic_resubmission(provider, monkeypatch):
    monkeypatch.setenv('AICHAT_MODEL_READ_TIMEOUT_SECONDS', '200')
    with pytest.raises(httpx.ReadTimeout) as caught:
        engine.ask([{'role': 'user', 'content': 'Вопрос'}])
    assert len(provider) == 1
    assert describe_ai_error(caught.value)['code'] == 'AI_TIMEOUT'


def test_qwen_compatibility_retry_keeps_long_timeout(provider, monkeypatch):
    calls = []

    def post(url, **kwargs):
        calls.append(kwargs)
        body = ({'error': 'unsupported chat_template_kwargs'} if len(calls) == 1
                else {'choices': [{'message': {'content': 'Совместимый ответ'}}]})
        return httpx.Response(400 if len(calls) == 1 else 200, json=body,
                              request=httpx.Request('POST', url))

    monkeypatch.setattr(httpx, 'post', post)
    assert engine.ask([{'role': 'user', 'content': 'Вопрос'}]) == 'Совместимый ответ'
    assert len(calls) == 2
    assert all(call['timeout'].read == 900 and call['timeout'].connect == 15 for call in calls)
    assert 'chat_template_kwargs' not in calls[-1]['json']
