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
    page.evaluate.side_effect = [[], [], 'Внутренняя ошибка\nSomething went wrong', '<body></body>',
                                 [table], [], 'Свод задач', '<body>Свод задач</body>']
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


def test_nvos_still_loading_after_date_does_not_reuse_old_text(tmp_path):
    from unittest.mock import MagicMock
    page, context, playwright = MagicMock(), MagicMock(), MagicMock()
    context.pages = [page]
    page.frames = []
    with patch.object(scraper, 'DEBUG_DIR', tmp_path / 'debug'), \
         patch.object(scraper, 'PLAYWRIGHT_PROFILE_DIR', tmp_path / 'profile'), \
         patch.object(scraper, 'SOURCES', [{'id': 'nvos', 'name': 'НВОС', 'url': 'https://example.invalid/nvos'}]), \
         patch.object(scraper, 'sync_playwright', return_value=playwright), \
         patch.object(scraper, 'launch_context', return_value=(context, 'Microsoft Edge')), \
         patch.object(scraper, '_select_latest_date', return_value='2026-09-10') as select_date, \
         patch.object(scraper, 'stab_wait_after_date'), \
         patch.object(scraper, '_wait_for_content', side_effect=[None, TimeoutError('Timeout')]) as ready:
        result = scraper.scrape_all()
    assert ready.call_count == 2 and select_date.call_count == 1
    assert result['nvos']['error'] == diagnostics.MESSAGES['timeout']
    assert result['nvos']['text'] == '' and result['nvos']['widgets'] == []
    context.close.assert_called_once()


def test_single_source_pipeline_passes_selection_and_reports_its_own_failure(tmp_path):
    import json
    with patch.object(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json'):
        before = builder.build_snapshot({'meetings': {'tables': [{'headers': ['ОМСУ', 'Явка, %'], 'rows': [['Округ', '60']]}]}})
        with patch.object(scraper, 'scrape_all', return_value={'tasks': {'error': 'Нет сети'}}) as collect:
            with pytest.raises(runner.SourcesUnavailableError):
                runner.run_water_dashboard_pipeline(source='tasks')
        collect.assert_called_once_with(source_ids=['tasks'])
        after = json.loads(builder.SNAPSHOT_FILE.read_text())
    assert after['sources']['meetings'] == before['sources']['meetings']
    assert after['sources_updated']['meetings'] is True
    assert after['last_updated_sources'] == []


def test_cli_source_argument_is_strict():
    with patch.object(runner, 'run_water_dashboard_pipeline') as pipeline:
        assert runner.main(['--source', 'tasks']) == 0
        pipeline.assert_called_once_with(source='tasks')
    with pytest.raises(SystemExit) as exc:
        runner.main(['--source', 'flush'])
    assert exc.value.code == 2


def test_scraper_selected_source_visits_only_its_url(tmp_path):
    from unittest.mock import MagicMock
    page, context, playwright = MagicMock(), MagicMock(), MagicMock()
    context.pages = [page]
    page.frames = []
    page.evaluate.side_effect = [[], [], 'Свод задач', '<body>Свод задач</body>']
    sources = [{'id': 'valves', 'name': 'Задвижки', 'url': 'https://example.invalid/valves'},
               {'id': 'tasks', 'name': 'Задачи', 'url': 'https://example.invalid/tasks'}]
    with patch.object(scraper, 'DEBUG_DIR', tmp_path / 'debug'), \
         patch.object(scraper, 'PLAYWRIGHT_PROFILE_DIR', tmp_path / 'profile'), \
         patch.object(scraper, 'SOURCES', sources), \
         patch.object(scraper, 'sync_playwright', return_value=playwright), \
         patch.object(scraper, 'launch_context', return_value=(context, 'Microsoft Edge')), \
         patch.object(scraper, '_wait_for_content'):
        result = scraper.scrape_all(source_ids=['tasks'])
    assert set(result) == {'tasks'}
    page.goto.assert_called_once_with('https://example.invalid/tasks', wait_until='domcontentloaded', timeout=90000)
    context.close.assert_called_once()


def test_full_edo_table_uses_bounded_navigation_and_preserves_kpis(tmp_path):
    from unittest.mock import MagicMock
    from services.water_dashboard.details import EDO_CURRENT_HEADER
    page, context, playwright = MagicMock(), MagicMock(), MagicMock()
    context.pages = [page]
    page.frames = []
    overview = {'headers': ['РСО', 'ОМСУ'], 'rows': [['Заполнено', 'Округ']]}
    widgets = [{'label': 'Доля (%) должностных лиц, имеющих право подписи и ЭЦП', 'value': '61 %'}]
    ranking = {'headers': ['РСО', 'ОМСУ', EDO_CURRENT_HEADER], 'rows': [['Водоканал', 'Округ', '83']]}
    page.evaluate.side_effect = [[overview], widgets, 'Обзор ЭДО', [ranking], '<body>таблица</body>']
    with patch.object(scraper, 'DEBUG_DIR', tmp_path / 'debug'), \
         patch.object(scraper, 'PLAYWRIGHT_PROFILE_DIR', tmp_path / 'profile'), \
         patch.object(scraper, 'SOURCES', [{'id': 'edo_rso', 'name': 'ЭДО', 'url': 'https://datalens.yandex/f5wqqij889haz'}]), \
         patch.object(scraper, 'sync_playwright', return_value=playwright), \
         patch.object(scraper, 'launch_context', return_value=(context, 'Microsoft Edge')), \
         patch.object(scraper, '_wait_for_content') as wait:
        result = scraper.scrape_all(source_ids=['edo_rso'])['edo_rso']
    assert result['tables'] == [overview] and result['widgets'] == widgets
    assert result['ranking_tables'] == [ranking]
    assert result['ranking_url'] == 'https://datalens.yandex/f5wqqij889haz?tab=EL'
    assert page.goto.call_count == 2 and wait.call_count == 2
    assert page.goto.call_args.kwargs['timeout'] == 90000
    context.close.assert_called_once()


def test_edo_ranking_timeout_is_separate_from_overview_metrics(tmp_path):
    from unittest.mock import MagicMock
    page, context, playwright = MagicMock(), MagicMock(), MagicMock()
    context.pages = [page]
    page.frames = []
    widgets = [{'label': 'Доля (%) должностных лиц, имеющих право подписи и ЭЦП', 'value': '61 %'}]
    page.evaluate.side_effect = [[], widgets, 'Обзор ЭДО', '<body>таблица</body>']
    with patch.object(scraper, 'DEBUG_DIR', tmp_path / 'debug'), \
         patch.object(scraper, 'PLAYWRIGHT_PROFILE_DIR', tmp_path / 'profile'), \
         patch.object(scraper, 'SOURCES', [{'id': 'edo_rso', 'name': 'ЭДО', 'url': 'https://datalens.yandex/f5wqqij889haz'}]), \
         patch.object(scraper, 'sync_playwright', return_value=playwright), \
         patch.object(scraper, 'launch_context', return_value=(context, 'Microsoft Edge')), \
         patch.object(scraper, '_wait_for_content', side_effect=[None, TimeoutError('Timeout')]):
        result = scraper.scrape_all(source_ids=['edo_rso'])['edo_rso']
    assert 'error' not in result and result['widgets'] == widgets
    assert result['ranking_tables'] == []
    assert result['ranking_error'].startswith('Полная таблица РСО не получена.')
    context.close.assert_called_once()
