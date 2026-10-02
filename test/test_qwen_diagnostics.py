"""Diagnostics use synthetic .env files and a mocked network only."""
import importlib.util
import json
import os
from unittest.mock import patch
from pathlib import Path

import httpx
import pytest

spec = importlib.util.spec_from_file_location(
    'qwen_diagnostics_test', Path(__file__).resolve().parents[1] / 'scripts' / 'check_qwen.py')
diag = importlib.util.module_from_spec(spec)
spec.loader.exec_module(diag)


@pytest.fixture
def settings(tmp_path, monkeypatch):
    with patch.dict(os.environ):
        for name in diag.VARIABLES:
            os.environ.pop(name, None)
        monkeypatch.setattr(diag, 'ENV_FILE', tmp_path / '.env')
        yield tmp_path / '.env'


def write(env, content):
    env.write_text(content, encoding='utf-8')


def assert_redacted(report, *values):
    text = json.dumps(report, ensure_ascii=False)
    for value in values:
        assert value not in text
    assert 'fingerprint' not in text and 'length' not in text


def test_readonly_diagnosis_uses_project_file_without_printing_values(settings, monkeypatch):
    write(settings, 'QWEN_API_KEY=synthetic-private-key\nQWEN_API_BASE=http://127.0.0.1:11434/v1\nQWEN_MODEL=synthetic-model\n')
    before = dict(os.environ)
    result = diag.diagnose()
    assert result['env_file_found'] is True
    assert result['variables']['QWEN_API_KEY']['source'] == 'dotenv'
    assert result['variables']['QWEN_API_KEY']['dotenv_has_value'] is True
    assert result['variables']['QWEN_API_KEY']['conflict'] is False
    assert result['endpoint_source'] == 'QWEN_API_BASE:dotenv'
    assert result['model_source'] == 'dotenv'
    assert dict(os.environ) == before
    assert_redacted(result, 'synthetic-private-key', 'synthetic-model', '127.0.0.1')


def test_inherited_key_overrides_dotenv_and_is_reported_as_conflict(settings, monkeypatch):
    write(settings, 'QWEN_API_KEY=fresh-file-key\n')
    monkeypatch.setenv('QWEN_API_KEY', 'stale-inherited-key')
    result = diag.diagnose()
    assert result['variables']['QWEN_API_KEY']['source'] == 'process'
    assert result['variables']['QWEN_API_KEY']['conflict'] is True
    assert 'PROCESS_ENV_OVERRIDES_DOTENV' in result['warnings']
    assert_redacted(result, 'fresh-file-key', 'stale-inherited-key')


def test_empty_inherited_key_also_blocks_dotenv_without_silent_override(settings, monkeypatch):
    write(settings, 'QWEN_API_KEY=file-key\n')
    monkeypatch.setenv('QWEN_API_KEY', '')
    result = diag.diagnose()
    assert result['variables']['QWEN_API_KEY']['source'] == 'process'
    assert result['variables']['QWEN_API_KEY']['process_has_value'] is False
    assert result['variables']['QWEN_API_KEY']['conflict'] is True
    assert 'KEY_NOT_CONFIGURED' in result['warnings']
    assert os.environ['QWEN_API_KEY'] == ''


def test_same_effective_value_is_not_a_conflict(settings, monkeypatch):
    write(settings, 'QWEN_API_KEY=matching-key\n')
    monkeypatch.setenv('QWEN_API_KEY', ' matching-key ')
    assert diag.diagnose()['variables']['QWEN_API_KEY']['conflict'] is False


def test_duplicate_qwen_names_and_bearer_prefix_are_safe(settings):
    write(settings, 'QWEN_API_KEY=first-secret\nQWEN_API_KEY="Bearer second-secret"\nSECRET_AS_NAME=hidden\nSECRET_AS_NAME=hidden-again\n')
    result = diag.diagnose()
    assert result['duplicate_variables'] == ['QWEN_API_KEY']
    assert 'KEY_CONTAINS_BEARER_PREFIX' in result['warnings']
    assert 'DUPLICATE_VARIABLES_IN_DOTENV' in result['warnings']
    assert_redacted(result, 'first-secret', 'second-secret', 'SECRET_AS_NAME', 'hidden')


def test_missing_file_and_bare_variable_do_not_claim_configured_key(settings):
    result = diag.diagnose()
    assert result['env_file_found'] is False and result['env_file_error'] == 'not_found'
    assert result['endpoint_source'] == 'default'
    write(settings, 'QWEN_API_KEY\n')
    result = diag.diagnose()
    assert result['variables']['QWEN_API_KEY']['source'] == 'unset'


def test_malformed_lines_do_not_leak_contents(settings, capsys):
    write(settings, 'QWEN_API_KEY="unterminated-supersecret\n')
    result = diag.diagnose()
    assert result['dotenv_parse_errors'] == 1
    assert 'DOTENV_PARSE_ERRORS' in result['warnings']
    assert_redacted(result, 'unterminated-supersecret')
    assert 'unterminated-supersecret' not in capsys.readouterr().err


def test_default_command_does_not_access_network_and_ignores_cwd_dotenv(settings, tmp_path, monkeypatch, capsys):
    write(settings, 'QWEN_API_KEY=project-file-key\n')
    other = tmp_path / 'other'; other.mkdir()
    write(other / '.env', 'QWEN_API_KEY=wrong-cwd-key\n')
    monkeypatch.chdir(other)
    def forbidden(*args, **kwargs):
        raise AssertionError('network access in read-only diagnostics')
    monkeypatch.setattr(httpx, 'post', forbidden)
    assert diag.main([]) == 0
    result = json.loads(capsys.readouterr().out)
    assert 'probe' not in result
    assert result['variables']['QWEN_API_KEY']['source'] == 'dotenv'
    assert_redacted(result, 'project-file-key', 'wrong-cwd-key')
    assert 'QWEN_API_KEY' not in os.environ


def test_probe_uses_actual_transport_preserves_conflict_and_redacts_401(settings, monkeypatch, capsys):
    original = ('QWEN_API_KEY=fresh-file-key\nQWEN_API_BASE=http://127.0.0.1:11434/v1\n'
                'QWEN_MODEL=synthetic-model\n')
    write(settings, original)
    monkeypatch.setenv('QWEN_API_KEY', 'stale-process-key')
    calls = []
    def post(url, **kwargs):
        calls.append(kwargs)
        assert kwargs['headers']['Authorization'] == 'Bearer stale-process-key'
        assert kwargs['json']['messages'] == [{'role': 'user', 'content': 'Ответь: OK'}]
        return httpx.Response(401, text='raw-body-secret', request=httpx.Request('POST', url))
    monkeypatch.setattr(httpx, 'post', post)
    assert diag.main(['--probe']) == 1
    captured = capsys.readouterr()
    result = json.loads(captured.out)
    assert result['probe'] == {'ok': False, 'code': 'AI_AUTH_FAILED', 'http_status': 401}
    assert result['variables']['QWEN_API_KEY']['conflict'] is True
    assert len(calls) == 1 and settings.read_text() == original
    assert_redacted(result, 'raw-body-secret', 'stale-process-key', 'fresh-file-key', 'synthetic-model', '127.0.0.1', 'Ответь: OK')
    assert 'secret' not in captured.err


def test_successful_probe_never_prints_model_answer(settings, monkeypatch, capsys):
    write(settings, 'QWEN_API_BASE=http://127.0.0.1:11434/v1\nQWEN_MODEL=synthetic-model\n')
    def post(url, **kwargs):
        return httpx.Response(200, json={'choices': [{'message': {'content': 'DO_NOT_PRINT_MODEL_ANSWER'}}]},
                              request=httpx.Request('POST', url))
    monkeypatch.setattr(httpx, 'post', post)
    assert diag.main(['--probe']) == 0
    result = json.loads(capsys.readouterr().out)
    assert result['probe']['ok'] is True and result['probe']['code'] == 'AI_OK'
    assert_redacted(result, 'DO_NOT_PRINT_MODEL_ANSWER', 'synthetic-model', '127.0.0.1')


def test_summarizer_override_is_compared_without_exposing_keys(settings):
    write(settings, 'QWEN_API_KEY=general-secret\nSUMMARIZER_QWEN_API_KEY=different-summary-secret\n')
    result = diag.diagnose()
    assert result['summarizer_override'] == {'configured': True, 'differs_from_chat_key': True, 'source': 'dotenv'}
    assert_redacted(result, 'general-secret', 'different-summary-secret')
    write(settings, 'QWEN_API_KEY=matching-secret\nSUMMARIZER_QWEN_API_KEY=matching-secret\n')
    assert diag.diagnose()['summarizer_override']['differs_from_chat_key'] is False
