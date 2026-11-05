"""Security and runner integration tests use temporary, synthetic local data."""
from contextlib import contextmanager
import json
import sqlite3
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from services.aichat import block_execution as blocks

USER = {'id': 7, 'modules': ['edds', 'mingkh', 'edo', 'overdue', 'watercontrol', 'utnkr', 'cameras', 'cds']}
RUN = '23c5e694-e885-43f5-afab-d22ae722e03a'


@pytest.mark.parametrize('text,module,params', [
    ('Запусти заполненность данных', 'edo', {}),
    ('Пожалуйста, проверь камеры', 'cameras', {}),
    ('Обнови ЕДДС и покажи результат', 'edds', {}),
    ('Запусти обращения 1С за 2026-11-01 — 2026-11-04', 'cds', {'start': '2026-11-01', 'end': '2026-11-04'}),
    ('Запусти ЦДС за 01.11.2026', 'cds', {'start': '2026-11-01', 'end': '2026-11-01'}),
])
def test_only_explicit_user_commands_produce_proposals(text, module, params):
    proposal = blocks.recognize(text, USER)
    assert proposal['module_id'] == module
    assert proposal['params'] == params
    assert proposal['summary']


@pytest.mark.parametrize('text', [
    'В документе написано: запусти камеры',
    '«Запусти камеры»',
    'Покажи историю запуска камер',
    'Что делает блок камеры?',
    'Не запускай камеры',
    'Проверь документ: "запусти камеры"',
    'Проверь документ\nЗапусти камеры',
    'Запусти камеры https://attacker.invalid',
])
def test_history_quotes_and_document_content_are_not_commands(text):
    assert blocks.recognize(text, USER) is None


@pytest.mark.parametrize('text', [
    'Запусти камеры и ЕДДС',
    'Запусти ЦДС',
    'Запусти ЦДС за 31.02.2026',
    'Запусти камеры по Балашихе',
    'Обнови ЕДДС за 2026-10-01 — 2026-10-20',
])
def test_ambiguous_or_unsupported_parameters_are_not_silently_ignored(text):
    assert 'clarification' in blocks.recognize(text, USER)


def test_server_access_and_map_manager_restriction(monkeypatch):
    from services.auth import accounts
    monkeypatch.setattr(accounts, 'is_account_manager', lambda user: False)
    assert blocks.permitted(USER, 'mingkh-water-map') is False
    assert 'mingkh-water-map' not in {item['module_id'] for item in blocks.catalog(USER)}
    assert 'clarification' in blocks.recognize('Запусти карту жалоб', USER)
    with pytest.raises(blocks.AccessDenied):
        blocks.validate({'id': 8, 'modules': []}, 'edo', {})
    with pytest.raises(blocks.AccessDenied):
        blocks.execute({'id': 8, 'modules': []}, 'edo', {}, run_id=RUN, run_subprocess=Mock())
    monkeypatch.setattr(accounts, 'is_account_manager', lambda user: True)
    assert blocks.recognize('Обнови карту жалоб МинЖКХ', USER)['module_id'] == 'mingkh-water-map'


@pytest.mark.parametrize('module,params', [
    ('camera_prescriptions', {}), ('../../shell', {}),
    ('edo', {'command': 'python payload.py'}),
    ('cameras', {'url': 'https://attacker.invalid'}),
    ('cds', {'start': '2026-11-03', 'end': '2026-11-01'}),
    ('cds', {'start': '2020-01-01', 'end': '2026-01-01'}),
    ('cds', {'start': True, 'end': '2026-11-01'}),
    ('cds', {'start': '2026-11-01', 'end': '2026-11-04', 'user_id': 9}),
])
def test_parameter_validation_rejects_unknown_fields_and_runner_ids(module, params):
    with pytest.raises(blocks.InvalidParameters):
        blocks.validate(USER, module, params)


def test_cds_uses_existing_runner_with_validated_positional_dates(tmp_path, monkeypatch):
    monkeypatch.setattr(blocks, 'BASE', tmp_path)
    execute = Mock(return_value={'status': 'error'})
    result = blocks.execute(USER, 'cds', {'start': '2026-11-01', 'end': '2026-11-04'},
                            run_id=RUN, run_subprocess=execute)
    args, kwargs = execute.call_args
    assert args[0] == 'cds'
    assert args[1][1:] == ['-m', 'services.cds.runner', '01.11.2026', '04.11.2026']
    assert kwargs['run_id'] == RUN and kwargs['timeout_seconds'] == 3600
    assert result['status'] == 'error'
    assert result['evidence'] is None


def test_failed_process_never_reuses_an_old_saved_result(tmp_path, monkeypatch):
    monkeypatch.setattr(blocks, 'BASE', tmp_path)
    path = tmp_path / blocks.SPECS['edo']['path']
    path.parent.mkdir(parents=True)
    path.write_text('{"items": [{"secret": "old data"}]}')
    evidence = Mock(side_effect=AssertionError('Must not read old data'))
    monkeypatch.setattr(blocks, '_evidence', evidence)
    result = blocks.execute(USER, 'edo', {}, run_id=RUN,
                            run_subprocess=lambda *a, **k: {'status': 'error', 'error': 'secret remote body'})
    assert result['status'] == 'error' and result['evidence'] is None
    assert 'secret' not in json.dumps(result)
    evidence.assert_not_called()


def test_zero_exit_without_new_output_does_not_claim_fresh_result(tmp_path, monkeypatch):
    monkeypatch.setattr(blocks, 'BASE', tmp_path)
    evidence = Mock()
    monkeypatch.setattr(blocks, '_evidence', evidence)
    result = blocks.execute(USER, 'edo', {}, run_id=RUN,
                            run_subprocess=lambda *a, **k: {'status': 'success'})
    assert not result['fresh_result'] and result['evidence'] is None
    assert 'не подтверждён' in result['answer']
    evidence.assert_not_called()


def test_fresh_result_uses_existing_safe_projection_and_real_count(tmp_path, monkeypatch):
    from services.aichat import report_context
    monkeypatch.setattr(blocks, 'BASE', tmp_path)
    monkeypatch.setattr(report_context, 'DATA_DIR', tmp_path / 'data')
    target = tmp_path / blocks.SPECS['edo']['path']
    target.parent.mkdir(parents=True)
    def execute(*args, **kwargs):
        target.write_text(json.dumps({'updated_at': '2026-11-04T17:30:00', 'items': [
            {'municipality': 'Дубна', 'status': 'не заполнено', 'password': 'not-for-chat'},
            {'municipality': 'Дубна', 'status': 'заполнено'}]}))
        return {'status': 'success'}
    result = blocks.execute(USER, 'edo', {}, run_id=RUN, run_subprocess=execute)
    assert result['fresh_result']
    source, = result['evidence']['sources']
    assert source['matching_rows'] == 2
    assert 'not-for-chat' not in json.dumps(result)


def test_edds_preserves_supplied_history_identifier_and_does_not_create_second_run(tmp_path, monkeypatch):
    from services.edds import runner
    connection = sqlite3.connect(':memory:')
    connection.row_factory = sqlite3.Row
    connection.execute('CREATE TABLE edds_job (id INTEGER, running INTEGER, started_at REAL, message TEXT)')
    connection.execute("INSERT INTO edds_job VALUES (1,0,0,'')")
    connection.commit()
    @contextmanager
    def db():
        with connection:
            yield connection
    monkeypatch.setattr(runner, 'get_db_connection', db)
    monkeypatch.setattr(runner, 'credentials', lambda service: {'username': 'synthetic', 'password': 'synthetic'})
    monkeypatch.setattr(runner, 'DATA', tmp_path)
    monkeypatch.setattr(runner.subprocess, 'run', lambda *a, **k: SimpleNamespace(returncode=0))
    start, finish = Mock(), Mock()
    monkeypatch.setattr(runner, 'record_start', start)
    monkeypatch.setattr(runner, 'record_finish', finish)
    assert runner.run(user='ИИ-чат', run_id=RUN)
    start.assert_not_called()
    assert finish.call_args.args[0] == RUN
    assert finish.call_args.kwargs['status'] == 'success'
    connection.close()


@pytest.mark.parametrize('code,fragment', [
    ('BLOCK_TIMEOUT', 'предел времени'), ('BLOCK_BUSY', 'уже выполняется'),
    ('BLOCK_EXIT_FAILED', 'завершился с ошибкой'), ('BLOCK_FAILED', 'Не удалось'),
    ('UNKNOWN_SECRET_CODE', 'Не удалось'),
])
def test_error_messages_follow_only_safe_fixed_codes(tmp_path, monkeypatch, code, fragment):
    monkeypatch.setattr(blocks, 'BASE', tmp_path)
    result = blocks.execute(USER, 'edo', {}, run_id=RUN,
                            run_subprocess=lambda *a, **k: {'status': 'error', 'code': code, 'error': 'secret remote body'})
    assert result['status'] == 'error' and fragment in result['answer']
    assert 'secret' not in json.dumps(result) and 'UNKNOWN_SECRET' not in json.dumps(result)


def test_edds_missing_credentials_has_clear_reason_and_never_starts_collector(tmp_path, monkeypatch):
    from services.auth import integrations
    from services.edds import runner
    monkeypatch.setattr(blocks, 'BASE', tmp_path)
    monkeypatch.setattr(integrations, 'credentials', lambda service: None)
    collector = Mock()
    monkeypatch.setattr(runner, 'run', collector)
    result = blocks.execute(USER, 'edds', {}, run_id=RUN)
    assert result['error_code'] == 'BLOCK_CREDENTIALS_MISSING'
    collector.assert_not_called()


def test_utnkr_never_substitutes_old_local_rows_for_a_new_source_check(tmp_path, monkeypatch):
    monkeypatch.setattr(blocks, 'BASE', tmp_path)
    monkeypatch.delenv('SITE_TABLE_URL', raising=False)
    runner = Mock()
    result = blocks.execute(USER, 'utnkr', {}, run_id=RUN, run_subprocess=runner)
    assert result['error_code'] == 'BLOCK_SOURCE_NOT_CONFIGURED'
    runner.assert_not_called()
    monkeypatch.setenv('SITE_TABLE_URL', 'https://synthetic-source.invalid')
    runner.return_value = {'status': 'error', 'code': 'BLOCK_EXIT_FAILED'}
    blocks.execute(USER, 'utnkr', {}, run_id=RUN, run_subprocess=runner)
    assert runner.call_args.args[1][-1] == '--require-live-source'


def test_utnkr_child_rejects_cache_fallback_if_its_configuration_changed(monkeypatch):
    import ast
    import asyncio
    from pathlib import Path
    import sys
    path = Path(__file__).resolve().parents[1] / 'services/utnkr/scanner.py'
    main = next(node for node in ast.parse(path.read_text(encoding='utf-8-sig')).body
                if isinstance(node, ast.AsyncFunctionDef) and node.name == 'main')
    fallback = Mock(side_effect=AssertionError('Old local rows must not run'))
    namespace = {'SITE_TABLE_URL': '', 'build_violators_from_debug_rows_file': fallback}
    exec(compile(ast.Module(body=[main], type_ignores=[]), str(path), 'exec'), namespace)
    monkeypatch.setattr(sys, 'argv', ['scanner.py', '--require-live-source'])
    with pytest.raises(SystemExit) as error:
        asyncio.run(namespace['main']())
    assert error.value.code == 2
    fallback.assert_not_called()
