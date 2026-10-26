"""Verified-source meeting totals and per-stage counts, using synthetic data."""
from unittest.mock import Mock, patch

from services.water_dashboard import builder, scraper
from services.water_dashboard.details import (
    KONTUR_STAGE_HEADERS, KONTUR_COMPLETION_HEADER, KONTUR_TASK_HEADER,
    kontur_stages, source_details,
)
from services.water_dashboard.metrics import source_metrics


def rso_table(rows):
    return {'headers': ['РСО', 'ОМСУ', 'Присутствовали на перекличках, %', 'Всего перекличек'], 'rows': rows}


def stage_table(rows):
    return {'headers': ['РСО', 'ОМСУ', KONTUR_TASK_HEADER, KONTUR_COMPLETION_HEADER, *KONTUR_STAGE_HEADERS], 'rows': rows}


def metrics(sid, data):
    return {m['id']: m['value'] for m in source_metrics(sid, data, builder.build_table({sid: data}))}


def test_omsu_and_rso_meeting_totals_are_separate_and_not_roll_call_counts():
    data = {'widgets': [{'label': 'Проведено совещаний', 'value': '2'}],
            'tables': [{'headers': ['ОМСУ', 'Присутствовали на перекличках, %'], 'rows': [['Город', '70']]}],
            'rso': {'widgets': [{'label': 'Проведено совещаний', 'value': '3'}],
                    'tables': [rso_table([['АО А', 'Город', '0', '800'], ['АО Б', 'Город', '60', '900']])]}}
    data['rso']['tables'] *= 2  # DOM clones must not double the population.
    assert metrics('meetings', data) == {'attendance_pct': 70, 'meetings': 2, 'rso_attendance_pct': 30, 'rso_meetings': 3}
    detail = source_details('meetings', data)
    assert [g['title'] for g in detail['groups']] == ['Топ-5 отстающих ОМСУ', 'Топ-5 отстающих РСО']
    assert len(detail['rso_entities']) == 2
    assert detail['groups'][1]['items'][0]['value'] == 0


def test_rso_conflicting_branches_exclude_average_not_guess():
    data = {'rso': {'tables': [rso_table([['АО А', 'Город', '20', '100'], ['АО А', 'Город', '40', '100'], ['АО Б', 'Город', '70', '100']])]}}
    assert metrics('meetings', data)['rso_attendance_pct'] is None
    assert source_details('meetings', data)['rso_coverage'] == {'entities': 1, 'excluded': 1}


def test_rso_collection_failure_never_reuses_omsu_total():
    data = {'widgets': [{'label': 'Проведено совещаний', 'value': '6'}],
            'rso': {'error': 'Сводка недоступна', 'widgets': [{'label': 'Проведено совещаний', 'value': '9'}]}}
    assert metrics('meetings', data)['meetings'] == 6
    assert metrics('meetings', data)['rso_meetings'] is None


def test_rso_summary_survives_snapshot_persistence(tmp_path, monkeypatch):
    monkeypatch.setattr(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json')
    data = {'widgets': [{'label': 'Проведено совещаний', 'value': '2'}],
            'tables': [{'headers': ['ОМСУ', 'Присутствовали на перекличках, %'], 'rows': [['Город', '70']]}],
            'rso': {'widgets': [{'label': 'Проведено совещаний', 'value': '3'}], 'tables': [rso_table([['АО А', 'Город', '60', '7']])]}}
    snap = builder.build_snapshot({'meetings': data}, source_ids=['meetings'])
    assert snap['sources']['meetings']['rso'] == data['rso']
    assert snap['sources']['meetings']['details']['rso_entities'][0]['name'] == 'АО А'


def test_stages_deduplicate_full_tables_and_preserve_branch_identity():
    row1 = ['АО А', 'Город', '101', '10', 'Закрыта', 'Новая', 'В работе', 'Просрочена', 'На проверке', 'Получен отказ']
    row2 = ['АО А', 'Город', '102', '20', 'Закрыта', 'Закрыта', 'Новая', 'На проверке', 'Возвращена на доработку', 'Новая']
    data = {'ranking_tables': [stage_table([row1]), stage_table([row2]), stage_table([row1])]}
    stages = kontur_stages(data)
    assert len(stages) == 6
    assert all(stage['total'] == 2 for stage in stages)
    assert stages[0]['pending'] == 0
    assert stages[1]['pending'] == 1
    assert stages[3]['statuses']['Просрочена'] == 1
    assert metrics('gkh_kontur', data)['stage_6_pending'] == 2
    assert source_details('gkh_kontur', data)['stages'] == stages


def test_unknown_or_conflicting_stage_is_not_zero_and_incomplete_pagination_is_not_total():
    row = ['АО А', 'Город', '101', '10', 'Закрыта', '—', 'Новая', 'Новая', 'Новая', 'Новая']
    data = {'ranking_tables': [stage_table([row])]}
    stages = kontur_stages(data)
    assert stages[0]['pending'] == 0
    assert stages[1]['pending'] is None and stages[1]['unknown'] == 1
    conflicting = row.copy(); conflicting[4] = 'Новая'
    data['ranking_tables'].append(stage_table([conflicting]))
    assert all(s['pending'] is None for s in kontur_stages(data))
    assert kontur_stages({**data, 'ranking_error': 'Вторая страница не получена'}) == []
    assert kontur_stages({'tables': [stage_table([row])]}) == []


def test_rso_reader_reads_every_page_and_fails_closed_on_stuck_pagination():
    first = rso_table([['АО А', 'Город', '50', '10']])
    second = rso_table([['АО Б', 'Город', '0', '10']])
    page = Mock()
    next_button = page.locator.return_value
    next_button.count.return_value = 1
    next_button.is_enabled.side_effect = [True, False]
    page.evaluate.side_effect = [[{'label': 'Проведено совещаний', 'value': '5'}], [first], [first], [second], [second]]
    with patch.object(scraper, '_wait_for_content'):
        result = scraper._read_meeting_rso(page, 'https://datalens.yandex/example')['rso']
    assert len(result['tables']) == 2 and not result.get('error')
    assert '?tab=Wk' in result['url']
    next_button.click.assert_called_once()
    page = Mock()
    page.locator.return_value.count.return_value = 1
    page.locator.return_value.is_enabled.return_value = True
    page.evaluate.side_effect = lambda script: [] if script == scraper.WIDGETS_JS else [first]
    with patch.object(scraper, '_wait_for_content'):
        result = scraper._read_meeting_rso(page, 'https://datalens.yandex/example')['rso']
    assert result['error'] and not result['tables']
