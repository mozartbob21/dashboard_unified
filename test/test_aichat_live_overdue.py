"""Real collector interfaces with synthetic network captures; no live portal call."""
import json
from contextlib import contextmanager
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from services.aichat import live_overdue_tool as live
from services.overdue import service

USER = {'id': 7, 'is_active': True, 'modules': ['overdue']}


@pytest.fixture(autouse=True)
def taxonomy(monkeypatch):
    monkeypatch.setattr(live.complaint_causes, '_known_cities', lambda: {'чехов': 'Чехов', 'дубна': 'Дубна'})


def write_capture(paths, rows=None, *, wrong_id=False, status=200):
    paths.responses.mkdir(parents=True, exist_ok=True)
    meta = {'url': 'https://synthetic.invalid/getPublicEntry', 'status': 200,
            'data': {'data': {'tabs': [{'items': [{'type': 'widget', 'data': {'tabs': [{'title': 'ОМСУ', 'chartId': 'reviewed-chart'}]}}]}]}}}
    values = rows if rows is not None else [['Чехов', 0], ['Дубна', 17]]
    chart = {'url': 'https://synthetic.invalid/charts/api/run', 'status': status,
             'data': {'id': 'different-chart' if wrong_id else 'reviewed-chart',
                      'columns': [{'name': 'ОМСУ'}, {'name': 'Количество задач'}], 'data': values}}
    (paths.responses/'001.json').write_text(json.dumps(meta, ensure_ascii=False))
    (paths.responses/'002.json').write_text(json.dumps(chart, ensure_ascii=False))
    return {'public_entry_seen': True, 'chart_run_count': 1, 'capture_limit_hit': False}


def test_live_current_city_zero_is_confirmed_and_shared_files_are_untouched(tmp_path, monkeypatch):
    saved = tmp_path/'shared.json'
    saved.write_text('{"old":"must survive"}')
    monkeypatch.setattr(service, 'DATA_FILE', saved)
    originals = service.PLAYWRIGHT_PROFILE_DIR, service.DEBUG_ENV_DIR, service.RESPONSES_DIR
    seen = []
    def fetch(paths, **options):
        seen.append(paths)
        assert options == {'headless': True, 'quiet': True, 'max_capture_bytes': live.MAX_BYTES, 'max_wait_seconds': 180}
        assert paths.profile != service.PLAYWRIGHT_PROFILE_DIR and paths.responses != service.RESPONSES_DIR
        return write_capture(paths)
    monkeypatch.setattr(service, 'fetch_dashboard_data', fetch)
    progress = []
    result = live.execute(USER, live.ID, {'municipality': 'Чехов'}, progress.append)
    assert result['status'] == 'ok' and result['metrics'][0]['value'] == 0
    assert result['effective_period'] is None and result['coverage'] == 'snapshot'
    assert result['as_of'] and result['snapshot_id'] and len(progress) == 2
    assert saved.read_text() == '{"old":"must survive"}'
    assert originals == (service.PLAYWRIGHT_PROFILE_DIR, service.DEBUG_ENV_DIR, service.RESPONSES_DIR)
    assert not seen[0].responses.parent.exists()
    assert 'synthetic.invalid' not in json.dumps(result)


@pytest.mark.parametrize('rows', [[['Дубна', 17]], [['Чехов', None]], [['Чехов', 'неизвестно']], [['Чехов', -1]], [['Чехов', True]], [['Чехов', 1], ['Чехов', 2]]])
def test_missing_or_invalid_source_counts_never_become_zero(rows, monkeypatch):
    monkeypatch.setattr(service, 'fetch_dashboard_data', lambda paths, **options: write_capture(paths, rows))
    result = live.execute(USER, live.ID, {'municipality': 'Чехов'})
    assert result['status'] == 'missing' and result['metrics'] == []
    assert result['warnings']


def test_source_failure_never_uses_previous_debug_results(monkeypatch):
    monkeypatch.setattr(service, 'fetch_dashboard_data', Mock(side_effect=RuntimeError('secret server credentials')))
    extract = Mock(side_effect=AssertionError('Legacy shared snapshot must not be read'))
    monkeypatch.setattr(service, 'extract_dashboard_data', extract)
    result = live.execute(USER, live.ID, {})
    assert result['status'] == 'missing' and not result['metrics']
    assert 'secret' not in json.dumps(result)
    extract.assert_not_called()


@pytest.mark.parametrize('options', [{'wrong_id': True}, {'status': 502}])
def test_only_confirmed_omsu_widget_success_is_accepted(options, monkeypatch):
    monkeypatch.setattr(service, 'fetch_dashboard_data', lambda paths, **kwargs: write_capture(paths, **options))
    result = live.execute(USER, live.ID, {})
    assert result['status'] == 'missing' and not result['metrics']


def test_source_capture_limits_invalidate_partial_collection(monkeypatch):
    def fetch(paths, **options):
        return {**write_capture(paths), 'capture_limit_hit': True}
    monkeypatch.setattr(service, 'fetch_dashboard_data', fetch)
    assert live.execute(USER, live.ID, {})['status'] == 'missing'


def test_access_and_historical_period_rejected_before_browser(monkeypatch):
    fetch = Mock()
    monkeypatch.setattr(service, 'fetch_dashboard_data', fetch)
    with pytest.raises(live.block_execution.AccessDenied):
        live.execute({**USER, 'modules': []}, live.ID, {})
    with pytest.raises(live.block_execution.AccessDenied):
        live.execute({**USER, 'is_active': False}, live.ID, {})
    for params in ({'start': '2026-02-01', 'end': '2026-02-28'}, {'url': 'https://evil.invalid'}, {'municipality': 'Несуществующий'}):
        with pytest.raises(live.block_execution.InvalidParameters):
            live.execute(USER, live.ID, params)
    fetch.assert_not_called()


def test_default_runpaths_preserve_existing_collector_configuration(tmp_path):
    paths = service._paths()
    assert paths.profile == service.PLAYWRIGHT_PROFILE_DIR
    assert paths.responses == service.RESPONSES_DIR
    assert paths.debug == service.DEBUG_ENV_DIR
    isolated = service.RunPaths.isolated(tmp_path)
    service.prepare_dirs(isolated)
    assert isolated.profile.is_dir() and isolated.responses.is_dir() and isolated.debug.is_dir()
    assert service.load_wrappers(isolated) == []


@pytest.mark.parametrize('limit,expected_limit_hit', [(live.MAX_BYTES, False), (2, True)])
def test_existing_fetch_function_writes_only_explicit_work_paths_and_honors_capture_limit(tmp_path, monkeypatch, capsys, limit, expected_limit_hit):
    paths = service.RunPaths.isolated(tmp_path/'private-run')
    shared = tmp_path/'shared-sentinel'
    shared.mkdir()
    marker = shared/'keep.json'
    marker.write_text('unchanged')
    monkeypatch.setattr(service, 'RESPONSES_DIR', shared)
    monkeypatch.setattr(service, 'DEBUG_ENV_DIR', shared)
    monkeypatch.setattr(service, 'PLAYWRIGHT_PROFILE_DIR', shared)
    monkeypatch.setattr(service, 'stabilize_page', lambda page: None)
    payloads = [('getPublicEntry', {'data': {'tabs': []}}), ('getPublicDashState', {'state': True}),
                ('charts/api/run', {'id': 'chart', 'columns': ['ОМСУ', 'Количество задач'], 'data': [['Чехов', 3]]})]
    callbacks = {}
    class Page:
        url = 'https://synthetic.invalid'
        def on(self, name, callback): callbacks[name] = callback
        def goto(self, *args, **kwargs):
            for suffix, payload in payloads:
                callbacks['response'](SimpleNamespace(url=self.url+'/'+suffix, status=200,
                    headers={'content-type': 'application/json'}, json=lambda payload=payload: payload))
        def wait_for_load_state(self, *args, **kwargs): pass
        def wait_for_timeout(self, *args, **kwargs): pass
        def title(self): return 'synthetic dashboard'
        def content(self): return '<html>synthetic</html>'
        def screenshot(self, path, **kwargs):
            from pathlib import Path
            Path(path).write_bytes(b'not-a-real-image')
        def evaluate(self, script): return {'source': 'synthetic', 'items': []}
    context = SimpleNamespace(pages=[Page()], close=Mock())
    launch = Mock(return_value=context)
    @contextmanager
    def playwright():
        yield SimpleNamespace(chromium=SimpleNamespace(launch_persistent_context=launch))
    monkeypatch.setattr(service, 'sync_playwright', playwright)
    fetched = service.fetch_dashboard_data(paths, headless=True, quiet=True, max_capture_bytes=limit, max_wait_seconds=0)
    assert fetched['capture_limit_hit'] is expected_limit_hit
    assert launch.call_args.kwargs['user_data_dir'] == str(paths.profile)
    assert launch.call_args.kwargs['headless'] is True
    assert marker.read_text() == 'unchanged'
    assert list(shared.iterdir()) == [marker]
    assert paths.debug.joinpath('dashboard_page.html').is_file()
    context.close.assert_called_once()
    assert capsys.readouterr().out == ''
