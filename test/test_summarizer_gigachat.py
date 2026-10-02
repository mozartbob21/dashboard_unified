"""GigaChat REST contract; all requests use synthetic credentials and HTTP mocks."""
import importlib
import ssl
import time
import uuid
from unittest.mock import patch

import httpx
import pytest

OAUTH_URL = 'https://ngw.devices.sberbank.ru:9443/api/v2/oauth'
CHAT_URL = 'https://api.giga.chat/v1/chat/completions'
LEGACY_CHAT_URL = 'https://gigachat.devices.sberbank.ru/api/v1/chat/completions'
GENERAL_CREDS = 'Z2VuZXJhbDpzeW50aGV0aWM='
SUMMARY_CREDS = 'c3VtbWFyeTpzeW50aGV0aWM='
EXPLICIT_CREDS = 'ZXhwbGljaXQ6c3ludGhldGlj'
PROMPT = [{'role': 'user', 'content': 'Синтетический тест'}]


@pytest.fixture(autouse=True)
def no_real_network(monkeypatch):
    def blocked(*args, **kwargs):
        raise AssertionError('Live network is forbidden in GigaChat tests')
    monkeypatch.setattr(httpx, 'post', blocked)
    monkeypatch.setattr(httpx.Client, 'send', blocked)


@pytest.fixture
def giga(monkeypatch):
    for name in ('SUMMARIZER_GIGACHAT_CREDENTIALS', 'GIGACHAT_CREDENTIALS',
                 'GIGACHAT_MODEL', 'GIGACHAT_MODE', 'GIGACHAT_SCOPE',
                 'GIGACHAT_BASE_URL', 'GIGACHAT_CA_BUNDLE_FILE', 'AI_CA_BUNDLE'):
        monkeypatch.delenv(name, raising=False)
    module = importlib.import_module('services.summarizer.gigachat_client')
    module._TOKENS.clear()
    yield module
    module._TOKENS.clear()


def response(url, status=200, data=None):
    return httpx.Response(status, json=data or {}, request=httpx.Request('POST', url))


def token(url=OAUTH_URL, value='synthetic-access-token', expired=False):
    expires = time.time() - 10 if expired else time.time() + 3600
    return response(url, data={'access_token': value, 'expires_at': int(expires * 1000)})


def answer(url=CHAT_URL, text='Синтетический ответ'):
    return response(url, data={'choices': [{'message': {'content': text}}]})


def assert_verified_transport(kwargs):
    assert kwargs['trust_env'] is False
    assert kwargs['follow_redirects'] is False
    assert isinstance(kwargs['verify'], ssl.SSLContext)
    assert kwargs['verify'].verify_mode == ssl.CERT_REQUIRED
    assert kwargs['verify'].check_hostname is True


def test_oauth_basic_scope_rquid_and_chat_bearer_are_separate(giga, monkeypatch):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    monkeypatch.setenv('GIGACHAT_SCOPE', 'GIGACHAT_API_PERS')
    calls = []
    def post(url, **kwargs):
        calls.append((url, kwargs))
        assert_verified_transport(kwargs)
        if url == OAUTH_URL:
            assert kwargs['headers']['Authorization'] == 'Basic ' + GENERAL_CREDS
            assert str(uuid.UUID(kwargs['headers']['RqUID'])) == kwargs['headers']['RqUID']
            assert kwargs['data'] == {'scope': 'GIGACHAT_API_PERS'}
            assert 'json' not in kwargs
            return token()
        assert url == CHAT_URL
        assert kwargs['headers']['Authorization'] == 'Bearer synthetic-access-token'
        assert kwargs['json']['messages'] == PROMPT
        assert kwargs['json']['model'] == 'GigaChat-2-Pro'
        assert kwargs['json']['max_tokens'] == 432
        assert GENERAL_CREDS not in str(kwargs['json'])
        return answer()
    monkeypatch.setattr(httpx, 'post', post)
    assert giga.chat(PROMPT, max_tokens=432) == 'Синтетический ответ'
    assert [url for url, _ in calls] == [OAUTH_URL, CHAT_URL]


@pytest.mark.parametrize('explicit,has_summary,expected', [
    (None, False, GENERAL_CREDS),
    (None, True, SUMMARY_CREDS),
    (EXPLICIT_CREDS, True, EXPLICIT_CREDS),
])
def test_credentials_precedence_is_explicit_then_summarizer_then_general(giga, monkeypatch, explicit, has_summary, expected):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    if has_summary:
        monkeypatch.setenv('SUMMARIZER_GIGACHAT_CREDENTIALS', SUMMARY_CREDS)
    def post(url, **kwargs):
        if url == OAUTH_URL:
            assert kwargs['headers']['Authorization'] == 'Basic ' + expected
            return token()
        return answer(url)
    monkeypatch.setattr(httpx, 'post', post)
    assert giga.chat(PROMPT, creds=explicit) == 'Синтетический ответ'


@pytest.mark.parametrize('canonical,alias,explicit,expected', [
    (None, 'GigaChat-2-Max', None, 'GigaChat-2-Max'),
    ('GigaChat-2-Pro', 'GigaChat-2-Max', None, 'GigaChat-2-Pro'),
    ('GigaChat-2-Pro', 'GigaChat-2-Max', 'GigaChat-2', 'GigaChat-2'),
])
def test_model_alias_does_not_override_canonical_or_explicit_model(giga, monkeypatch, canonical, alias, explicit, expected):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    if canonical: monkeypatch.setenv('GIGACHAT_MODEL', canonical)
    if alias: monkeypatch.setenv('GIGACHAT_MODE', alias)
    def post(url, **kwargs):
        if url == OAUTH_URL: return token()
        assert kwargs['json']['model'] == expected
        return answer(url)
    monkeypatch.setattr(httpx, 'post', post)
    assert giga.chat(PROMPT, model=explicit) == 'Синтетический ответ'


def test_access_token_is_reused_until_expiry(giga, monkeypatch):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    calls = []
    def post(url, **kwargs):
        calls.append(url)
        return token() if url == OAUTH_URL else answer(url)
    monkeypatch.setattr(httpx, 'post', post)
    assert giga.chat(PROMPT)
    assert giga.chat(PROMPT)
    assert calls == [OAUTH_URL, CHAT_URL, CHAT_URL]


@pytest.mark.parametrize('expires', ['NaN', 'Infinity', '-Infinity', 0])
def test_invalid_token_expiry_is_rejected_before_chat(giga, monkeypatch, expires):
    from core.ai_errors import describe_ai_error
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    calls = []
    def post(url, **kwargs):
        calls.append(url)
        return response(url, data={'access_token': 'synthetic-token', 'expires_at': expires})
    monkeypatch.setattr(httpx, 'post', post)
    with pytest.raises(Exception) as failed:
        giga.chat(PROMPT)
    assert describe_ai_error(failed.value)['code'] == 'AI_INVALID_RESPONSE'
    assert calls == [OAUTH_URL]
    assert not giga._TOKENS


def test_chat_401_refreshes_token_once_and_retries(giga, monkeypatch):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    sequence = [token(value='old-token'), response(CHAT_URL, 401), token(value='new-token'), answer()]
    calls = []
    def post(url, **kwargs):
        calls.append((url, kwargs))
        return sequence.pop(0)
    monkeypatch.setattr(httpx, 'post', post)
    assert giga.chat(PROMPT) == 'Синтетический ответ'
    assert [url for url, _ in calls] == [OAUTH_URL, CHAT_URL, OAUTH_URL, CHAT_URL]
    assert calls[1][1]['headers']['Authorization'] == 'Bearer old-token'
    assert calls[3][1]['headers']['Authorization'] == 'Bearer new-token'
    assert not sequence


def test_repeated_401_does_not_loop_forever(giga, monkeypatch):
    from core.ai_errors import describe_ai_error
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    responses = [token(), response(CHAT_URL, 401), token(), response(CHAT_URL, 401)]
    def post(url, **kwargs):
        assert responses, 'Unbounded authentication retries'
        return responses.pop(0)
    monkeypatch.setattr(httpx, 'post', post)
    with pytest.raises(Exception) as failed:
        giga.chat(PROMPT)
    assert describe_ai_error(failed.value)['code'] == 'AI_AUTH_FAILED'
    assert not responses


def test_different_credentials_and_scopes_do_not_share_tokens(giga, monkeypatch):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    authorizations = []
    def post(url, **kwargs):
        if url == OAUTH_URL:
            authorizations.append((kwargs['headers']['Authorization'], kwargs['data']['scope']))
            return token(value='token-' + str(len(authorizations)))
        return answer(url)
    monkeypatch.setattr(httpx, 'post', post)
    giga.chat(PROMPT)
    monkeypatch.setenv('SUMMARIZER_GIGACHAT_CREDENTIALS', SUMMARY_CREDS)
    giga.chat(PROMPT)
    monkeypatch.setenv('GIGACHAT_SCOPE', 'GIGACHAT_API_CORP')
    giga.chat(PROMPT)
    assert len(authorizations) == 3
    assert authorizations[0][0] != authorizations[1][0]
    assert authorizations[1][1] != authorizations[2][1]


def test_official_legacy_endpoint_is_supported(giga, monkeypatch):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    monkeypatch.setenv('GIGACHAT_BASE_URL', 'https://gigachat.devices.sberbank.ru/api/v1')
    calls = []
    def post(url, **kwargs):
        calls.append(url)
        return token() if url == OAUTH_URL else answer(url)
    monkeypatch.setattr(httpx, 'post', post)
    giga.chat(PROMPT)
    assert calls == [OAUTH_URL, LEGACY_CHAT_URL]


def test_unapproved_endpoint_is_rejected_before_credentials_are_sent(giga, monkeypatch):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    monkeypatch.setenv('GIGACHAT_BASE_URL', 'https://unapproved.invalid/v1')
    with patch.object(httpx, 'post') as network:
        with pytest.raises(Exception):
            giga.chat(PROMPT)
    network.assert_not_called()


def test_oauth_failures_have_safe_error_codes_without_credentials(giga, monkeypatch, capsys):
    from core.ai_errors import describe_ai_error
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    def post(url, **kwargs):
        return response(url, 401, {'error': 'SECRET_RESPONSE_BODY'})
    monkeypatch.setattr(httpx, 'post', post)
    with pytest.raises(Exception) as failed:
        giga.chat(PROMPT)
    description = describe_ai_error(failed.value)
    assert description['code'] == 'AI_AUTH_FAILED'
    assert GENERAL_CREDS not in str(description)
    assert 'SECRET_RESPONSE_BODY' not in str(description)
    assert 'SECRET_RESPONSE_BODY' not in capsys.readouterr().out


def test_normal_ai_chat_stays_on_qwen_when_giga_credentials_exist(giga, monkeypatch):
    from services.aichat import engine as chat_engine
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    monkeypatch.setenv('SUMMARIZER_GIGACHAT_CREDENTIALS', SUMMARY_CREDS)
    with patch.object(giga, 'chat') as giga_call, patch.object(chat_engine, '_qwen_chat', return_value='Qwen response') as qwen:
        assert chat_engine.ask([{'role': 'user', 'content': 'Привет'}]) == 'Qwen response'
    qwen.assert_called_once()
    giga_call.assert_not_called()


def test_summarizer_giga_selection_calls_giga_and_preserves_ai_backend(giga, monkeypatch):
    from services.summarizer import engine
    monkeypatch.setenv('SUMMARIZER_GIGACHAT_CREDENTIALS', SUMMARY_CREDS)
    with patch.object(giga, 'chat', return_value='Проверенный результат GigaChat') as giga_call, patch.object(engine, '_qwen_chat') as qwen:
        result = engine.summarize('Достаточно длинный синтетический текст для проверки работы сумматора.', backend='gigachat')
    assert result['ok'] is True and result['backend'] == 'gigachat'
    assert result['report_text'] == 'Проверенный результат GigaChat'
    giga_call.assert_called_once()
    qwen.assert_not_called()


def test_token_expiration_in_milliseconds_triggers_new_oauth(giga, monkeypatch):
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    calls = []
    def post(url, **kwargs):
        calls.append(url)
        return token() if url == OAUTH_URL else answer(url)
    monkeypatch.setattr(httpx, 'post', post)
    clock = time.time()
    giga.chat(PROMPT)
    monkeypatch.setattr(giga.time, 'time', lambda: clock + 7200)
    giga.chat(PROMPT)
    assert calls == [OAUTH_URL, CHAT_URL, OAUTH_URL, CHAT_URL]


@pytest.mark.parametrize('setting,expected', [
    ('missing_credentials', 'GIGACHAT_CREDENTIALS_MISSING'),
    ('invalid_scope', 'GIGACHAT_SCOPE_INVALID'),
])
def test_invalid_configuration_stops_before_network(giga, monkeypatch, setting, expected):
    from core.ai_errors import describe_ai_error
    if setting == 'invalid_scope':
        monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
        monkeypatch.setenv('GIGACHAT_SCOPE', 'unrecognized-scope')
    with patch.object(httpx, 'post') as network:
        with pytest.raises(Exception) as failed:
            giga.chat(PROMPT)
    assert describe_ai_error(failed.value)['code'] == expected
    network.assert_not_called()


def test_macros_and_html_converter_stay_on_qwen(giga, monkeypatch, tmp_path):
    import json
    from services.summarizer import engine
    from services.tools import macros
    from services.tools.pptx_converter import parse_html_to_slides_ai
    monkeypatch.setenv('GIGACHAT_CREDENTIALS', GENERAL_CREDS)
    monkeypatch.setenv('SUMMARIZER_GIGACHAT_CREDENTIALS', SUMMARY_CREDS)
    macro = {'code': 'Option Explicit\nPublic Sub Example()\nMsgBox "Test"\nEnd Sub',
             'instructions': 'Import module', 'assumptions': ''}
    slide = [{'title': 'Synthetic slide', 'texts': ['Content']}]
    with patch.object(giga, 'chat') as giga_call, patch.object(engine, '_qwen_chat', side_effect=[json.dumps(macro), json.dumps(slide)]) as qwen:
        result = macros.generate('Создай простой тестовый макрос.', 'Excel VBA', tmp_path)
        slides = parse_html_to_slides_ai('<div class="slide"><h1>Synthetic slide</h1><p>Content</p></div>')
    assert result['file'] == 'macro.bas'
    assert slides[0]['title'] == 'Synthetic slide'
    assert qwen.call_count == 2
    giga_call.assert_not_called()


def test_summarizer_auth_failure_is_labeled_as_algorithm_not_ai(giga, monkeypatch, caplog):
    from services.summarizer import engine
    monkeypatch.setenv('SUMMARIZER_GIGACHAT_CREDENTIALS', SUMMARY_CREDS)
    failed_response = response(CHAT_URL, 401, {'error': 'PRIVATE_RESPONSE'})
    error = httpx.HTTPStatusError('PRIVATE_EXCEPTION_WITH_KEY', request=failed_response.request, response=failed_response)
    with patch.object(giga, 'chat', side_effect=error), patch.object(engine, '_qwen_chat') as qwen:
        result = engine.summarize('Достаточно длинный синтетический текст для проверки резервного алгоритма.', backend='gigachat')
    assert result['backend'] == 'algo'
    assert result['requested_backend'] == 'gigachat'
    assert result['ai_error_code'] == 'AI_AUTH_FAILED'
    assert 'алгоритмом' in result['warning']
    assert 'PRIVATE_' not in result['warning']
    assert 'PRIVATE_' not in caplog.text
    qwen.assert_not_called()
