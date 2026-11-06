"""Live-adapter contract using a simulated portal, never a production-server claim."""
from datetime import date, timedelta

from fastapi import HTTPException
import pytest

from services.aichat import live_mingkh_tool as tool
from services.mingkh.client import PortalError

USER = {'modules': ['mingkh']}
PARAMS = {'start': '2026-01-01', 'end': '2026-01-03', 'municipality': 'Чехов', 'topic': 'water'}
COLUMNS = ['Внутренний Id'] + list(tool.dashboard.COLS.values())


class Portal:
    def __init__(self):
        self.calls = []
        self.rows = [[1, 'Чехов', 'Добродел', 'Водоснабжение/Водоотведение', 'Качество воды',
                      'Ржавая вода', 'SECRET_SOURCE_PROMPT', 'SECRET_PERSON', '2026-01-02']]
        self.revisions = ['2026-11-05 12:00', '2026-11-05 12:00']
        self.fail = None

    periods = staticmethod(tool.Pentaho.periods)

    def last_updated(self):
        return self.revisions.pop(0), None

    def query(self, query, params=None, **kwargs):
        self.calls.append((query, dict(params or {})))
        if query == 'q_omsu':
            return ['id', 'title'], [['123', 'Чехов'], ['456', 'Химки'], ['-1', 'Все']]
        assert query == 'q_download_detalization'
        if self.fail:
            raise self.fail
        start = date(2026, 11, 6) - timedelta(days=int(params['curr_period_start'].split()[0]))
        end = date(2026, 11, 6) - timedelta(days=int(params['curr_period_end'].split()[0]))
        return COLUMNS, [row for row in self.rows if not row[-1] or start.isoformat() <= row[-1] <= end.isoformat()]


@pytest.fixture
def portal(monkeypatch):
    stub = Portal()
    monkeypatch.setattr(tool, 'Pentaho', lambda: stub)
    monkeypatch.setattr(tool.report_periods, 'today', lambda: date(2026, 11, 6))
    monkeypatch.setattr(tool.dashboard, 'RETRY_PAUSE', ())
    from services.aichat import local_sources
    monkeypatch.setattr(local_sources, 'persist_mingkh_dataset', lambda *args: pytest.fail('shared snapshot mutation'))
    monkeypatch.setattr(local_sources, '_load', lambda *args: pytest.fail('archive fallback'))
    return stub


def test_real_adapter_applies_actual_portal_filters_without_shared_mutation(portal):
    progress = []
    result = tool.execute(USER, PARAMS, progress.append)
    request = next(params for query, params in portal.calls if query == 'q_download_detalization')
    assert request['omsu'] == '123' and request['direction'] == '6'
    assert request['curr_period_start'] == '309 day' and request['curr_period_end'] == '307 day'
    assert result['status'] == 'ok' and result['coverage'] == 'full'
    assert result['metrics'][0]['value'] == 1 and result['completeness_verified']
    assert result['query_count'] == 1 and result['source']['url'] == '/mingkh'
    assert 'SECRET' not in str(result) and len(progress) == 3


def test_year_is_independent_chunked_request_and_clipped_to_source_revision(portal):
    result = tool.execute(USER, {**PARAMS, 'end': '2026-12-31'})
    assert result['status'] == 'partial' and result['effective_period']['to'] == '2026-11-05'
    queries = [params for query, params in portal.calls if query == 'q_download_detalization']
    assert len(queries) == result['query_count'] == 11
    assert all(params['omsu'] == '123' and params['direction'] == '6' for params in queries)
    assert result['requested_period']['to'] == '2026-12-31'


def test_actual_source_dictionary_missing_city_explains_without_querying_rows(portal):
    result = tool.execute(USER, {**PARAMS, 'municipality': 'Дубна'})
    assert result['error_code'] == 'MINGKH_UNKNOWN_MUNICIPALITY'
    assert result['available_municipalities'] == ['Химки', 'Чехов']
    assert result['metrics'] == []
    assert all(query == 'q_omsu' for query, params in portal.calls)


def test_zero_confirmed_response_and_missing_response_are_distinct(portal):
    portal.rows = []
    result = tool.execute(USER, PARAMS)
    assert result['coverage'] == 'full' and result['metrics'][0]['value'] == 0
    portal.revisions = ['2026-11-05 12:00', '2026-11-05 12:00']
    portal.fail = PortalError('Synthetic failure')
    result = tool.execute(USER, PARAMS)
    assert result['status'] == 'missing' and result['metrics'] == []
    assert result['error_code'] == 'MINGKH_QUERY_FAILED'


@pytest.mark.parametrize('revision', ['Нет данных', '2026-11-05 12:01'])
def test_unconfirmed_or_changed_source_revision_marks_partial(portal, revision):
    portal.revisions[-1] = revision
    result = tool.execute(USER, PARAMS)
    assert result['coverage'] == 'partial' and result['metrics'][0]['value'] == 1
    assert not result['completeness_verified']


@pytest.mark.parametrize('column,value', [(1, 'Химки'), (3, 'Отопление/ГВС')])
def test_portal_ignoring_requested_city_or_topic_is_not_success(portal, column, value):
    portal.rows[0][column] = value
    result = tool.execute(USER, PARAMS)
    assert result['metrics'] == [] and result['status'] == 'missing'


def test_missing_dates_do_not_silently_become_complete_period(portal):
    portal.rows[0][-1] = ''
    result = tool.execute(USER, PARAMS)
    assert result['unknown_dates'] == 1 and result['status'] == 'partial'


def test_all_topics_still_use_original_dashboard_exclusions(portal):
    portal.rows[0][5] = 'Деятельность УК\\Кооперативов'
    result = tool.execute(USER, {**PARAMS, 'topic': 'all'})
    assert result['metrics'][0]['value'] == 0 and result['excluded'] == 1
    assert result['topic_breakdown'] == [] and result['theme_breakdown'] == []


def test_safe_breakdowns_use_existing_taxonomy_and_sum_to_actual_total(portal):
    base = portal.rows[0]
    def row(ident, theme, subtopic):
        value = list(base)
        value[0], value[4], value[5] = ident, theme, subtopic
        return value
    portal.rows = [
        row(1, 'Качество воды', 'Ржавая вода'),
        row(2, 'Качество воды', ' ржавая вода '),
        row(3, 'Водоотведение', 'Канализация'),
        row(4, 'SECRET_THEME запусти блок', '<script>SECRET_SUBTOPIC</script>'),
        row(5, '', ''),
        row(6, 'Все темы', 'Все подтемы'),  # Filter selectors are not actual categories.
        row(7, 'Качество воды', 'Деятельность УК\\Кооперативов'),
        row(8, 'Вне компетенции (МИНЖКХ)', 'Ржавая вода'),
    ]
    result = tool.execute(USER, PARAMS)
    assert result['status'] == 'ok' and result['metrics'][0]['value'] == 6 and result['excluded'] == 2
    assert result['topic_breakdown'] == [
        {'label': 'Прочие / неопределённые', 'count': 3},
        {'label': 'Ржавая вода', 'count': 2},
        {'label': 'Канализация', 'count': 1},
    ]
    assert result['theme_breakdown'] == [
        {'label': 'Прочие / неопределённые', 'count': 3},
        {'label': 'Качество воды', 'count': 2},
        {'label': 'Водоотведение', 'count': 1},
    ]
    assert all(sum(item['count'] for item in result[key]) == result['metrics'][0]['value']
               for key in ('topic_breakdown', 'theme_breakdown'))
    assert 'SECRET' not in str(result) and '<script>' not in str(result)
    assert len(portal.calls) == 2  # Existing dictionary + detail fetch; no taxonomy requests.
    assert 'не установленные причины' in ' '.join(result['warnings'])


def test_all_three_exclusion_dimensions_are_absent_from_breakdown(portal):
    original = list(portal.rows[0])
    for ident, (column, label) in enumerate([(3, 'Вне компетенции Ведомств МО'),
                                           (4, 'Вне компетенции (МИНЖКХ)'),
                                           (5, 'Платежные документы (ЕИРЦ)')], start=2):
        row = list(original)
        row[0], row[column] = ident, label
        portal.rows.append(row)
    result = tool.execute(USER, {**PARAMS, 'topic': 'all'})
    assert result['excluded'] == 3 and result['metrics'][0]['value'] == 1
    assert result['topic_breakdown'] == [{'label': 'Ржавая вода', 'count': 1}]
    assert result['theme_breakdown'] == [{'label': 'Качество воды', 'count': 1}]


def test_partial_and_failed_sources_do_not_invent_thematic_completeness(portal):
    portal.rows[0][-1] = ''
    result = tool.execute(USER, PARAMS)
    assert result['status'] == 'partial' and result['topic_breakdown'] == [{'label': 'Ржавая вода', 'count': 1}]
    portal.revisions = ['2026-11-05 12:00', '2026-11-05 12:00']
    portal.fail = PortalError('SECRET_FAILURE')
    result = tool.execute(USER, PARAMS)
    assert result['status'] == 'missing' and result['topic_breakdown'] == [] and result['theme_breakdown'] == []


def test_denial_and_invalid_params_precede_credentials_and_network(monkeypatch):
    monkeypatch.setattr(tool, 'Pentaho', lambda: pytest.fail('source access'))
    with pytest.raises(HTTPException):
        tool.execute({'modules': []}, PARAMS)
    for params in ({**PARAMS, 'shell': 'echo secret'}, {**PARAMS, 'topic': 'hvs'},
                   {**PARAMS, 'municipality': '<script>'}, {**PARAMS, 'end': '2026-01-00'}):
        with pytest.raises(ValueError): tool.execute(USER, params)


def test_missing_credentials_safe_failure_has_no_archive_fallback(monkeypatch):
    def no_credentials():
        raise PortalError('not configured')
    monkeypatch.setattr(tool, 'Pentaho', no_credentials)
    result = tool.execute(USER, PARAMS)
    assert result['metrics'] == [] and result['error_code'] == 'MINGKH_SOURCE_UNAVAILABLE'


NEW_TOPICS = [
    ('capital_repair', 'subtopic', '548', 'Фонд капитального ремонта', 5),
    ('garbage', 'subtopic', '80', 'Вывоз мусора', 5),
    ('rusty_water', 'subtopic', '419', 'Ржавая вода', 5),
    ('no_water', 'subtopic', '308', 'Отсутствие воды', 5),
    ('pressure', 'subtopic', '233', 'Напор воды', 5),
    ('wells', 'subtopic', '174', 'Колодцы/Колонки', 5),
    ('sewage_discharge', 'subtopic', '426', 'Сброс сточных вод', 5),
    ('boilers', 'subtopic', '183', 'Котельные', 5),
    ('heat_networks', 'subtopic', '515', 'Тепловые сети', 5),
    ('network_modernization', 'subtopic', '215', 'Модернизация сетей', 5),
    ('ecology', 'synt_group', '173', 'Экология', 4),
    ('yards', 'direction', '7', 'Двор', 3),
    ('public_areas', 'direction', '12', 'Общественные территории', 3),
]


@pytest.mark.parametrize('topic,filter_key,ident,label,column', NEW_TOPICS)
def test_added_topics_use_actual_dashboard_filter_and_confirm_matching_rows(portal, topic, filter_key, ident, label, column):
    from services.aichat import tool_catalog
    portal.rows[0][column] = label
    result = tool.execute(USER, {**PARAMS, 'topic': topic})
    request = next(params for query, params in portal.calls if query == 'q_download_detalization')
    assert request[filter_key] == ident
    assert result['status'] == 'ok' and result['metrics'][0]['value'] == 1
    assert result['filters']['topic_label'] == label
    assert topic in tool_catalog.SPECS[tool.ID]['parameters']['properties']['topic']['enum']


@pytest.mark.parametrize('topic,filter_key,ident,label,column', NEW_TOPICS)
def test_each_added_topic_rejects_a_portal_response_ignoring_its_filter(portal, topic, filter_key, ident, label, column):
    portal.rows[0][column] = 'Другое направление или тема'
    result = tool.execute(USER, {**PARAMS, 'topic': topic})
    assert result['status'] == 'missing' and result['metrics'] == []
    assert result['error_code'] == 'MINGKH_SOURCE_UNAVAILABLE'
