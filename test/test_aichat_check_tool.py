"""Actual output files and original projections; collectors are synthetic stubs."""
from datetime import date
import gzip
import json
from unittest.mock import Mock

import pytest

from services.aichat import block_execution as blocks, check_tool, tool_catalog, report_context

USER = {'id': 7, 'username': 'synthetic', 'is_active': True,
        'modules': ['edds', 'mingkh', 'edo', 'overdue', 'watercontrol', 'utnkr', 'cameras', 'cds']}
RUN = 'synthetic-server-run-identifier'


@pytest.fixture(autouse=True)
def isolated_environment(tmp_path, monkeypatch):
    from services.auth import accounts, integrations
    monkeypatch.setattr(accounts, 'is_account_manager', lambda user: True)
    monkeypatch.setattr(integrations, 'credentials', lambda source: {'username': 'synthetic', 'password': 'synthetic'})
    monkeypatch.setattr(blocks, 'BASE', tmp_path)
    monkeypatch.setattr(report_context, 'DATA_DIR', tmp_path/'data')
    monkeypatch.setattr(check_tool.report_periods, 'today', lambda: date(2026, 11, 6))
    monkeypatch.setattr(check_tool.complaint_causes, '_known_cities', lambda: {'чехов': 'Чехов', 'дубна': 'Дубна'})
    monkeypatch.setenv('SITE_TABLE_URL', 'https://synthetic.invalid')


def payload(module):
    if module == 'edds':
        return {'updated': '2026-11-06T12:00:00', 'days': {'2026-11-05': {'Чехов': [1, 2, 3], 'Дубна': [90, 90, 90]}}}
    if module == 'mingkh-water-map':
        return {'meta': {'live_updated': '2026-11-06T12:00:00', 'from': '2026-11-01', 'to': '2026-11-06'},
                'cols': ['omsu', 'kind'], 'rows': [[0, 0], [0, 1], [1, 2]], 'omsu': ['Чехов', 'Дубна']}
    result = {'updated_at': '2026-11-06T12:00:00', 'items': [
        {'municipality': 'Чехов', 'status': 'Нарушение', 'overdue_count': 3, 'password': 'secret'},
        {'municipality': 'Дубна', 'status': 'Норма', 'overdue_count': 90}]}
    if module == 'cds':
        result.update(date_from='01.10.2026', date_to='31.10.2026')
    return result


def write_output(module, data):
    path = blocks.BASE/blocks.SPECS[module]['path']
    path.parent.mkdir(parents=True, exist_ok=True)
    content = json.dumps(data, ensure_ascii=False).encode()
    path.write_bytes(gzip.compress(content) if str(path).endswith('.gz') else content)
    return path


@pytest.mark.parametrize('module', list(blocks.SPECS))
def test_each_existing_runner_returns_only_new_city_filtered_result(module, monkeypatch):
    def run(*args, **kwargs):
        write_output(module, payload(module))
        return {'status': 'success'}
    if module == 'edds':
        from services.edds import runner
        monkeypatch.setattr(runner, 'run', lambda **kwargs: bool(run()))
    elif module == 'mingkh-water-map':
        from services.mingkh import water_map_runner
        monkeypatch.setattr(water_map_runner, 'run', run)
    params = {'municipality': 'Чехов'}
    if module == 'cds':
        params.update(start='2026-10-01', end='2026-10-31')
    worker = Mock(side_effect=run)
    result = tool_catalog.execute(USER, 'check:'+module, params, run_id=RUN,
                                  run_subprocess=worker, confirmed=True)
    assert result['status'] == 'success' and result['fresh_result'] is True
    assert result['filters']['municipality'] == 'Чехов'
    assert '_output_version' not in result
    source, = result['evidence']['sources']
    assert source['scope'] == 'municipality'
    assert source['module'] == blocks.SPECS[module]['module']
    assert 'secret' not in json.dumps(result)
    if module == 'edds':
        assert [metric['value'] for metric in source['metrics']] == [1, 2, 3]
    elif module == 'mingkh-water-map':
        assert [metric['value'] for metric in source['metrics']] == [1, 1, 0, 0, 0]
    else:
        assert source['matching_rows'] == 1
        assert source['rows'][0]['Муниципалитет'] == 'Чехов'
    if module not in ('edds', 'mingkh-water-map'):
        command = worker.call_args.args[1]
        assert 'Чехов' not in command
        if module == 'cds':
            assert command[-2:] == ['01.10.2026', '31.10.2026']


def test_old_output_and_failed_collector_never_become_a_city_answer():
    write_output('edo', payload('edo'))
    result = check_tool.execute(USER, 'check:edo', {'municipality': 'Чехов'}, run_id=RUN,
                                run_subprocess=lambda *args, **kwargs: {'status': 'error'}, confirmed=True)
    assert result['status'] == 'error' and result['evidence'] is None
    no_write = check_tool.execute(USER, 'check:edo', {'municipality': 'Чехов'}, run_id=RUN,
                                  run_subprocess=lambda *args, **kwargs: {'status': 'success'}, confirmed=True)
    assert no_write['fresh_result'] is False and no_write['evidence'] is None


def test_replaced_output_version_cannot_be_used_for_another_run(monkeypatch):
    path = write_output('edo', payload('edo'))
    version = list(blocks._fingerprint(path))
    write_output('edo', {'items': [{'municipality': 'Чехов', 'overdue_count': 999999}]})
    monkeypatch.setattr(blocks, 'execute', lambda *args, **kwargs: {'status': 'success', 'fresh_result': True,
                          '_output_version': version, 'evidence': {'sources': ['old']}, 'answer': 'ok'})
    result = check_tool.execute(USER, 'check:edo', {}, run_id=RUN, confirmed=True)
    assert result['error_code'] == 'CHECK_RESULT_NOT_CONFIRMED'
    assert result['fresh_result'] is False and result['evidence'] is None
    assert '999999' not in json.dumps(result)


def test_output_change_during_read_is_detected(monkeypatch):
    path = write_output('edo', payload('edo'))
    version = blocks._fingerprint(path)
    monkeypatch.setattr(blocks, '_fingerprint', Mock(side_effect=[version, (version[0]+1, version[1], version[2])]))
    with pytest.raises(ValueError):
        check_tool._read(path, list(version))


def test_overdue_fallback_with_source_error_is_not_accepted_as_a_new_check():
    def run(*args, **kwargs):
        write_output('overdue', {**payload('overdue'), 'extraction_note': 'source failed; stale parser used'})
        return {'status': 'success'}
    result = check_tool.execute(USER, 'check:overdue', {}, run_id=RUN, confirmed=True, run_subprocess=run)
    assert result['status'] == 'error' and result['evidence'] is None
    assert 'stale parser' not in result['answer']


def test_missing_city_and_wrong_cds_period_are_not_reported_as_success():
    def run_edo(*args, **kwargs):
        write_output('edo', {'updated_at': '2026-11-06T12:00:00', 'items': [{'municipality': 'Дубна'}]})
        return {'status': 'success'}
    result = check_tool.execute(USER, 'check:edo', {'municipality': 'Чехов'}, run_id=RUN, confirmed=True, run_subprocess=run_edo)
    assert result['status'] == 'error' and result['evidence'] is None
    def run_cds(*args, **kwargs):
        write_output('cds', {**payload('cds'), 'date_from': '01.09.2026'})
        return {'status': 'success'}
    result = check_tool.execute(USER, 'check:cds', {'start': '2026-10-01', 'end': '2026-10-31'},
                                run_id=RUN, confirmed=True, run_subprocess=run_cds)
    assert result['status'] == 'error' and result['evidence'] is None


def test_check_requires_rights_confirmation_and_rejects_unsupported_filters(monkeypatch):
    execute = Mock()
    monkeypatch.setattr(blocks, 'execute', execute)
    with pytest.raises(blocks.AccessDenied):
        check_tool.execute({**USER, 'modules': []}, 'check:edo', {}, run_id=RUN, confirmed=True)
    with pytest.raises(blocks.InvalidParameters):
        check_tool.execute(USER, 'check:edo', {}, run_id=RUN)
    for params in ({'municipality': 'Unknown'}, {'start': '2026-01-01'}, {'url': 'https://bad.invalid'}, {'command': 'shell'}):
        with pytest.raises(blocks.InvalidParameters):
            check_tool.validate(USER, 'check:edo', params)
    execute.assert_not_called()
