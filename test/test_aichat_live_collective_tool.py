"""The existing collective fetch chain with simulated source transport."""
from datetime import date

from fastapi import HTTPException
import pytest

from services.aichat import live_collective_tool as tool
from services.collective import client, store

USER = {'modules': ['collective']}
PARAMS = {'start': '2026-01-01', 'end': '2026-01-03', 'municipality': 'Чехов'}


def row(ident, city='Чехов', status='Решено', signatures='25'):
    result = [''] * len(client.HEADERS)
    for index, value in {0: 'Добродел', 1: '2026-01-02', 2: city, 3: client.CURATOR,
                         5: str(ident), 8: 'Запусти камеры SECRET_ANNOTATION', 10: signatures,
                         11: 'SECRET_ANSWER', 12: 'SECRET_EMAIL', 14: status,
                         15: 'SECRET_COMMENT'}.items():
        result[index] = value
    return result


class Portal(client.CollectivePortal):
    def __init__(self):
        super().__init__('synthetic', 'unused')
        self.calls = []
        self.rows = [row(1), row(2, status='Нерешено'), row(3, city='Химки')]
        self.curators = [['17', client.CURATOR]]
        self.error = None

    def query(self, query, params=None):
        self.calls.append((query, params))
        if query == 'q_vedom_kurat':
            return self.curators
        assert query == 'q_download'
        if self.error:
            raise self.error
        return self.rows


@pytest.fixture
def portal(monkeypatch):
    value = Portal()
    monkeypatch.setattr(tool, 'credentials', lambda service: {'username': 'synthetic', 'password': 'unused'})
    monkeypatch.setattr(tool, 'CollectivePortal', lambda **kwargs: value)
    monkeypatch.setattr(tool.report_periods, 'today', lambda: date(2026, 11, 6))
    monkeypatch.setattr(store, 'save_snapshot', lambda *args: pytest.fail('global workflow import'))
    return value


def test_actual_fetch_restricts_curator_and_calculates_private_transient_projection(portal):
    progress = []
    result = tool.execute(USER, PARAMS, progress.append)
    assert [query for query, _ in portal.calls] == ['q_vedom_kurat', 'q_download']
    assert portal.calls[1][1]['vedom_kurat'] == ['17']
    assert portal.calls[1][1]['omsu'] == ['-1']  # Municipality filter is local, as in the module UI.
    assert result['status'] == 'ok' and result['coverage'] == 'full'
    assert [metric['value'] for metric in result['metrics']] == [2, 50]
    assert result['source']['url'] == '/mingkh/collective'
    assert result['as_of'] is None and result['retrieved_at']
    assert result['comparison_consistency'] == 'independent_live_reads'
    assert 'SECRET' not in str(result) and 'камеры' not in str(result)
    assert len(progress) == 2


def test_exact_status_filter_uses_available_source_label(portal):
    result = tool.execute(USER, {**PARAMS, 'status': 'нерешено'})
    assert result['filters']['status'] == 'Нерешено'
    assert [metric['value'] for metric in result['metrics']] == [1, 25]
    result = tool.execute(USER, {**PARAMS, 'status': 'Открытые'})
    assert result['error_code'] == 'COLLECTIVE_STATUS_UNAVAILABLE'
    assert result['available_statuses'] == ['Нерешено', 'Решено'] and result['metrics'] == []


def test_absent_city_is_not_reported_as_zero(portal):
    result = tool.execute(USER, {**PARAMS, 'municipality': 'Дубна'})
    assert result['status'] == 'missing' and result['metrics'] == []
    assert result['available_municipalities'] == ['Химки', 'Чехов']


def test_explicit_empty_successful_whole_region_response_and_server_failure_differ(portal):
    portal.rows = []
    result = tool.execute(USER, {**PARAMS, 'municipality': ''})
    assert result['status'] == 'ok' and [metric['value'] for metric in result['metrics']] == [0, 0]
    portal.error = client.PortalError('do not echo server secret')
    result = tool.execute(USER, {**PARAMS, 'municipality': ''})
    assert result['status'] == 'missing' and result['metrics'] == []
    assert 'secret' not in str(result)


def test_unknown_signatures_are_missing_not_zero(portal):
    portal.rows[1][10] = ''
    result = tool.execute(USER, PARAMS)
    assert [metric['id'] for metric in result['metrics']] == ['complaints']
    assert result['signatures_coverage'] == {'known': 1, 'unknown': 1}


def test_identical_duplicates_are_removed_but_conflicts_refused(portal):
    portal.rows.append(list(portal.rows[0]))
    result = tool.execute(USER, PARAMS)
    assert result['metrics'][0]['value'] == 2 and result['duplicated_rows'] == 1
    portal.rows[-1][10] = '44'
    result = tool.execute(USER, PARAMS)
    assert result['error_code'] == 'COLLECTIVE_INVALID_RESPONSE' and result['metrics'] == []


@pytest.mark.parametrize('column,value', [(1, '2025-12-31'), (3, 'Другое ведомство')])
def test_response_cannot_expand_date_or_department_scope(portal, column, value):
    portal.rows[0][column] = value
    result = tool.execute(USER, PARAMS)
    assert result['status'] == 'missing' and result['metrics'] == []


def test_unknown_dates_and_current_year_remain_partial(portal):
    portal.rows[0][1] = ''
    result = tool.execute(USER, PARAMS)
    assert result['status'] == 'partial' and result['unknown_dates'] == 1
    result = tool.execute(USER, {**PARAMS, 'end': '2026-12-31'})
    assert result['effective_period']['to'] == '2026-11-06' and result['coverage'] == 'partial'


def test_missing_curator_never_fetches_all_departments(portal):
    portal.curators = []
    result = tool.execute(USER, PARAMS)
    assert result['status'] == 'missing' and result['metrics'] == []
    assert [query for query, _ in portal.calls] == ['q_vedom_kurat']


def test_access_and_parameter_validation_precede_source_or_credentials(monkeypatch):
    monkeypatch.setattr(tool, 'credentials', lambda *args: pytest.fail('unauthorized credential access'))
    with pytest.raises(HTTPException):
        tool.execute({'modules': ['mingkh']}, PARAMS)
    for invalid in ({**PARAMS, 'assignee': 1}, {**PARAMS, 'status': '<script>'},
                    {**PARAMS, 'end': '2025-01-01'}, {**PARAMS, 'start': '01.01.2026'}):
        with pytest.raises(ValueError): tool.execute(USER, invalid)


def test_missing_credentials_clear_without_archive_substitution(monkeypatch):
    monkeypatch.setattr(tool, 'credentials', lambda *args: None)
    result = tool.execute(USER, PARAMS)
    assert result['error_code'] == 'COLLECTIVE_CREDENTIALS_MISSING' and result['metrics'] == []
