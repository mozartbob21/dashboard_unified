import io
from unittest.mock import patch

import pytest

from services.water_dashboard import builder, diagnostics, runner, scraper
from services.water_dashboard.browser import WaterDashboardBrowserError


@pytest.mark.parametrize('error,code', [
    (ModuleNotFoundError('private detail', name='playwright'), 'playwright_missing'),
    (ModuleNotFoundError('private detail', name='dotenv'), 'dependency'),
    (PermissionError('C:/Users/private'), 'storage'),
    (RuntimeError('net::ERR_NAME_NOT_RESOLVED at private url'), 'network'),
    (RuntimeError('net::ERR_CERT_AUTHORITY_INVALID at private url'), 'certificate'),
    (RuntimeError('Page.goto: Timeout 90000ms exceeded'), 'timeout'),
    (WaterDashboardBrowserError('private detail', code='browser_missing'), 'browser_missing'),
    (ValueError('private detail'), 'unknown'),
])
def test_cli_returns_failure_with_safe_diagnostic(error, code):
    stdout, stderr = io.StringIO(), io.StringIO()
    with patch.object(runner, 'run_water_dashboard_pipeline', side_effect=error), \
         patch('sys.stdout', stdout), patch('sys.stderr', stderr):
        assert runner.main() == 1
    assert stdout.getvalue().strip() == diagnostics.ERROR_PREFIX + code
    message = diagnostics.failure_message(stdout.getvalue())
    assert message == diagnostics.MESSAGES[code]
    assert 'private' not in message
    assert stderr.getvalue()  # Original diagnostic remains in server log.


def test_unrecognized_worker_output_never_becomes_public_error():
    assert diagnostics.failure_message('WATER_DASHBOARD_ERROR:private/password\n') == diagnostics.MESSAGES['unknown']
    assert diagnostics.failure_message('Exception with password=private') == diagnostics.MESSAGES['unknown']


def test_windows_console_can_log_browser_unicode_without_crashing_worker():
    buffer = io.BytesIO()
    console = io.TextIOWrapper(buffer, encoding='cp1251')
    with patch('sys.stdout', console):
        diagnostics.log_server_line('[WATER_DASHBOARD] Браузер ✓ 🧪')
    console.flush()
    assert 'Браузер' in buffer.getvalue().decode('cp1251')


def test_successful_pipeline_saves_actual_source_result(tmp_path):
    data = {'tasks': {'tables': [{'headers': ['ОМСУ', 'Количество задач'], 'rows': [['Тестовый округ', '18']]}]}}
    path = tmp_path / 'snapshot.json'
    with patch.object(scraper, 'scrape_all', return_value=data), patch.object(builder, 'SNAPSHOT_FILE', path):
        result = runner.run_water_dashboard_pipeline()
    assert result['kpis']['tasks_total'] == 18
    assert result['sources_updated']['tasks'] is True
    assert path.exists()


def test_all_sources_failed_persists_status_without_losing_previous_data(tmp_path):
    with patch.object(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json'):
        before = builder.build_snapshot({'tasks': {'tables': [{'headers': ['ОМСУ', 'Количество задач'], 'rows': [['Тестовый округ', '18']]}]}})
        with patch.object(scraper, 'scrape_all', return_value={'tasks': {'error': 'Нет соединения'}}):
            with pytest.raises(runner.SourcesUnavailableError):
                runner.run_water_dashboard_pipeline()
        import json
        after = json.loads(builder.SNAPSHOT_FILE.read_text())
    assert after['kpis']['tasks_total'] == 18
    assert after['sources']['tasks']['updated_at'] == before['sources']['tasks']['updated_at']
    assert after['sources']['tasks']['error'] == 'Нет соединения'
    assert not any(after['sources_updated'].values())


def test_scraper_keeps_other_sources_and_closes_browser_after_portal_error(tmp_path):
    from unittest.mock import MagicMock
    page, context, playwright = MagicMock(), MagicMock(), MagicMock()
    context.pages = [page]
    page.frames = []
    table = {'headers': ['ОМСУ', 'Количество задач'], 'rows': [['Тестовый округ', '18']]}
    page.evaluate.side_effect = [[], 'Внутренняя ошибка\nSomething went wrong', [table], 'Свод задач']
    sources = [{'id': 'flush', 'name': 'Промывки', 'url': 'https://example.invalid/flush'},
               {'id': 'tasks', 'name': 'Задачи', 'url': 'https://example.invalid/tasks'}]
    with patch.object(scraper, 'DEBUG_DIR', tmp_path / 'debug'), \
         patch.object(scraper, 'PLAYWRIGHT_PROFILE_DIR', tmp_path / 'profile'), \
         patch.object(scraper, 'SOURCES', sources), \
         patch.object(scraper, 'sync_playwright', return_value=playwright), \
         patch.object(scraper, 'launch_context', return_value=(context, 'Microsoft Edge')), \
         patch.object(scraper, '_wait_for_content'):
        result = scraper.scrape_all()
    assert result['flush']['error'] == diagnostics.MESSAGES['source_error']
    assert result['tasks']['tables'] == [table]
    context.close.assert_called_once()
    assert page.goto.call_count == 2
