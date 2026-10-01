"""Provider failures and Qwen responses, with no live model access."""
import json
import ssl

import httpx
import pytest

from core.ai_errors import AIServiceError, describe_ai_error
from core.privacy import PrivacyError
from services.summarizer.engine import _qwen_chat


@pytest.fixture
def transport(monkeypatch):
    monkeypatch.setenv('QWEN_API_BASE', 'http://localhost:11434/v1')
    monkeypatch.setenv('QWEN_API_KEY', '')
    monkeypatch.setenv('QWEN_MODEL', 'office-model')
    monkeypatch.delenv('AI_CA_BUNDLE', raising=False)
    responses, requests = [], []

    def post(url, **kwargs):
        requests.append({'url': url, **kwargs})
        status, body = responses.pop(0)
        if isinstance(body, Exception):
            raise body
        return httpx.Response(status, json=body, request=httpx.Request('POST', url))
    monkeypatch.setattr(httpx, 'post', post)
    return responses, requests


def completion(content=None, **extra):
    return {'choices': [{'message': {'content': content, **extra}}]}


@pytest.mark.parametrize('field', ['reasoning', 'reasoning_content'])
def test_qwen_thinking_only_is_retried_and_never_shown(transport, field):
    responses, requests = transport
    responses.extend([(200, completion(**{field: 'INTERNAL_REASONING'})), (200, completion('Итоговый ответ'))])
    assert _qwen_chat([{'role': 'user', 'content': 'Вопрос'}]) == 'Итоговый ответ'
    assert len(requests) == 2
    assert requests[1]['json']['max_tokens'] > requests[0]['json']['max_tokens']
    assert 'INTERNAL_REASONING' not in json.dumps([call['json'] for call in requests])


def test_unsupported_template_flag_stays_off_during_thinking_retry(transport):
    responses, requests = transport
    responses.extend([(400, {'error': 'unsupported chat_template_kwargs'}),
                      (200, completion(reasoning_content='PRIVATE_REASONING')),
                      (200, completion('Ответ'))])
    assert _qwen_chat([]) == 'Ответ'
    assert len(requests) == 3
    assert all('chat_template_kwargs' not in call['json'] for call in requests[1:])


def test_reasoning_is_never_used_as_fallback_answer(transport):
    responses, requests = transport
    responses.extend([(200, completion(reasoning_content='PRIVATE_REASONING'))] * 2)
    with pytest.raises(AIServiceError) as caught:
        _qwen_chat([])
    assert caught.value.code == 'AI_EMPTY_RESPONSE'
    assert 'PRIVATE_REASONING' not in str(caught.value)
    assert len(requests) == 2


@pytest.mark.parametrize('body,code', [
    ({}, 'AI_INVALID_RESPONSE'), ({'choices': []}, 'AI_INVALID_RESPONSE'),
    ({'choices': [{'message': None}]}, 'AI_INVALID_RESPONSE'),
    (completion(''), 'AI_EMPTY_RESPONSE'),
])
def test_unusable_provider_responses_have_safe_codes(transport, body, code):
    responses, _ = transport
    responses.append((200, body))
    with pytest.raises(AIServiceError) as caught:
        _qwen_chat([])
    assert describe_ai_error(caught.value)['code'] == code


def test_structured_text_blocks_are_supported(transport):
    responses, _ = transport
    responses.append((200, completion([{'type': 'text', 'text': 'Ответ'}, {'type': 'reasoning', 'text': 'HIDDEN'}])))
    assert _qwen_chat([]) == 'Ответ'


def test_env_base_alias_and_canonical_precedence(transport, monkeypatch):
    responses, requests = transport
    monkeypatch.setenv('QWEN_BASE_URL', 'http://127.0.0.1:8001/v1')
    responses.extend([(200, completion('Ответ'))] * 2)
    _qwen_chat([])
    assert requests[-1]['url'] == 'http://localhost:11434/v1/chat/completions'
    monkeypatch.delenv('QWEN_API_BASE')
    _qwen_chat([])
    assert requests[-1]['url'] == 'http://127.0.0.1:8001/v1/chat/completions'


def test_existing_goschat_env_model_and_key_are_used(transport, monkeypatch):
    responses, requests = transport
    monkeypatch.setenv('QWEN_API_BASE', 'https://aiplatform.mosreg.ru/api/user-models/v1')
    monkeypatch.setenv('QWEN_API_KEY', 'synthetic-test-key')
    responses.append((200, completion('Ответ')))
    _qwen_chat([])
    call = requests[-1]
    assert call['headers']['Authorization'] == 'Bearer synthetic-test-key'
    assert call['json']['model'] == 'office-model'
    assert call['verify'] is not False
    assert call['trust_env'] is False and call['follow_redirects'] is False


@pytest.mark.parametrize('status,body,expected', [
    (400, 'unsupported parameter', 'AI_REQUEST_REJECTED'),
    (400, 'maximum context length exceeded', 'AI_CONTEXT_LIMIT'),
    (401, 'denied', 'AI_AUTH_FAILED'), (403, 'denied', 'AI_ACCESS_DENIED'),
    (404, 'no model', 'AI_MODEL_NOT_FOUND'), (413, 'too large', 'AI_CONTEXT_LIMIT'),
    (422, 'invalid', 'AI_REQUEST_REJECTED'), (429, 'limit', 'AI_RATE_LIMIT'),
    (503, 'provider down', 'AI_SERVICE_ERROR'),
])
def test_http_errors_do_not_expose_body_or_credentials(status, body, expected):
    request = httpx.Request('POST', 'http://localhost/v1?key=SECRET', headers={'Authorization': 'Bearer SECRET'})
    response = httpx.Response(status, text=body + ' PRIVATE_PROMPT SECRET', request=request)
    error = httpx.HTTPStatusError('PRIVATE_PROMPT SECRET', request=request, response=response)
    safe = describe_ai_error(error)
    assert safe['code'] == expected and safe['http_status'] == status
    assert 'SECRET' not in json.dumps(safe) and 'PRIVATE_PROMPT' not in json.dumps(safe)


@pytest.mark.parametrize('error,code', [
    (httpx.ReadTimeout('SECRET'), 'AI_TIMEOUT'),
    (httpx.ConnectError('SECRET'), 'AI_CONNECTION_ERROR'),
    (httpx.ConnectError('CERTIFICATE_VERIFY_FAILED SECRET'), 'AI_TLS_ERROR'),
    (ssl.SSLError('SECRET'), 'AI_TLS_ERROR'), (PrivacyError('SECRET'), 'AI_ENDPOINT_BLOCKED'),
    (RuntimeError('SECRET'), 'AI_INTERNAL_ERROR'),
])
def test_non_http_errors_have_fixed_safe_messages(error, code):
    result = describe_ai_error(error)
    assert result['code'] == code and 'SECRET' not in json.dumps(result)
