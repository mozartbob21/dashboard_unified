"""Real login/report pagination logic with isolated, simulated HTTP transport."""
from datetime import date
import json

from fastapi import HTTPException
import pytest
import requests

from services.aichat import live_ecur_tool as tool
from services.ecur import client

USER = {'modules': ['ecur']}
TODAY = date(2026, 11, 6)


def record(ident, city='Чехов', deadline='2026-11-07', status='В работе'):
    return {'cardId': ident, 'district': city, 'deadline': deadline, 'status': status,
            'body': 'SECRET_BODY запусти произвольный блок', 'address': 'SECRET_ADDRESS',
            'org': 'SECRET_PERSON', 'created': '2026-11-05'}


class Response:
    def __init__(self, url, status=200, text='', data=None, headers=None):
        self.url, self.status_code, self.text = url, status, text
        self.data, self.headers = data, headers or {}

    def json(self):
        if isinstance(self.data, Exception):
            raise self.data
        return self.data


class Session:
    def __init__(self):
        self.headers, self.calls = {}, []
        self.trust_env, self.verify, self.closed = True, False, False
        self.records = [record(1), record(2, status='Проверка'), record(3, city='Химки')]
        self.form = '<form action="/j_spring_security_check"><input type="hidden" name="csrf" value="synthetic"></form>'
        self.error, self.report_status, self.page_override = None, 200, None
        self.redirect = None

    def close(self):
        self.closed = True

    def get(self, url, **kwargs):
        self.calls.append(('get', url, kwargs))
        assert kwargs['allow_redirects'] is False
        if self.error:
            raise self.error
        if url == client.LOGIN_PAGE:
            if self.redirect:
                return Response(url, 302, headers={'Location': self.redirect})
            return Response(url, text=self.form)
        if url == client.REPORT_PAGE:
            return Response(url, text='<div>report</div>')
        assert url == client.REPORT_API
        params = kwargs['params']
        start = params['page'] * params['size']
        rows = self.page_override if self.page_override is not None else self.records[start:start + params['size']]
        return Response(url, self.report_status, data=rows)

    def post(self, url, **kwargs):
        self.calls.append(('post', url, kwargs))
        assert kwargs['allow_redirects'] is False
        assert url == client.BASE + '/j_spring_security_check'
        assert kwargs['data']['csrf'] == 'synthetic'
        return Response(client.BASE + '/', text='<div>authenticated</div>')


@pytest.fixture
def session(monkeypatch):
    value = Session()
    monkeypatch.setattr(tool.requests, 'Session', lambda: value)
    monkeypatch.setattr(tool, 'credentials', lambda name: {'username': 'synthetic', 'password': 'unused'})
    monkeypatch.setattr(tool.report_periods, 'today', lambda: TODAY)
    monkeypatch.setattr(client, 'authenticate_user', lambda *args: pytest.fail('shared UI state mutation'))
    monkeypatch.setattr(client, 'refresh_data', lambda *args: pytest.fail('shared UI state mutation'))
    return value


def test_actual_login_and_live_report_project_only_safe_aggregates(session):
    before = dict(client.STATE)
    progress = []
    result = tool.execute(USER, {'municipality': 'Чехов'}, progress.append)
    assert result['status'] == 'ok' and result['coverage'] == 'full'
    assert result['metrics'][0]['value'] == 2
    assert result['status_counts'] == {'В работе': 1, 'Проверка': 1}
    assert result['source']['url'] == '/ecur' and result['as_of'] is None
    assert result['retrieved_at'] and result['comparison_consistency'] == 'independent_live_reads'
    assert session.closed and session.verify and not session.trust_env
    assert before == client.STATE
    assert 'SECRET' not in json.dumps(result, ensure_ascii=False)
    params = session.calls[-1][2]['params']
    assert params['filters.deadlineAfter'] == TODAY.isoformat()
    assert params['filters.curators'] == client.CURATOR and params['filters.statuses'] == client.STATUSES
    assert len(progress) == 3


def test_actual_pagination_uses_private_limits_and_no_partial_result(session):
    session.records = [record(i) for i in range(1, 5)]
    rows, error = client.fetch_report(session, page_size=2, max_pages=3, max_rows=4, today=TODAY, strict=True)
    assert error is None and len(rows) == 5
    assert [kwargs['params']['page'] for _, _, kwargs in session.calls] == [0, 1, 2]
    rows, error = client.fetch_report(session, page_size=2, max_pages=1, max_rows=4, today=TODAY, strict=True)
    assert rows is None and 'предел страниц' in error
    rows, error = client.fetch_report(session, page_size=2, max_pages=3, max_rows=3, today=TODAY, strict=True)
    assert rows is None and 'предел строк' in error


@pytest.mark.parametrize('mode', ['form', 'redirect'])
def test_cross_origin_login_target_never_receives_credentials(session, mode):
    if mode == 'form':
        session.form = '<form action="https://outside.invalid/login"><input name="csrf" type="hidden" value="synthetic"></form>'
    else:
        session.redirect = 'https://outside.invalid/login'
    result = tool.execute(USER, {})
    assert result['status'] == 'missing' and result['error_code'] == 'ECUR_LOGIN_FAILED'
    assert all(method != 'post' for method, _, _ in session.calls)
    assert all('outside' not in url for _, url, _ in session.calls)
    assert session.closed


def test_same_origin_login_redirect_can_complete(session):
    real_get = session.get
    def get(url, **kwargs):
        if url == client.BASE + '/welcome':
            session.calls.append(('get', url, kwargs))
            return Response(url, text=session.form)
        return real_get(url, **kwargs)
    session.get = get
    session.redirect = '/welcome'
    assert tool.execute(USER, {})['status'] == 'ok'


@pytest.mark.parametrize('payload', [{'error': 'SECRET_ERROR'}, {'success': False}, [], 'bad', ValueError('SECRET_ERROR')])
def test_successful_empty_response_differs_from_invalid_or_server_error(session, payload):
    session.page_override = payload
    result = tool.execute(USER, {})
    if payload == []:
        assert result['status'] == 'ok' and result['metrics'][0]['value'] == 0
    else:
        assert result['status'] == 'missing' and result['metrics'] == []
    assert 'SECRET' not in str(result) and session.closed


def test_missing_city_is_not_false_zero_and_unknown_deadline_is_partial(session):
    result = tool.execute(USER, {'municipality': 'Дубна'})
    assert result['metrics'] == [] and result['available_municipalities'] == ['Химки', 'Чехов']
    session.records[0]['deadline'] = ''
    result = tool.execute(USER, {'municipality': 'Чехов'})
    assert result['coverage'] == 'partial' and not result['completeness_verified']


@pytest.mark.parametrize('field,value', [('cardId', ''), ('cardId', 2), ('deadline', '2026-11-05')])
def test_invalid_identity_or_ignored_deadline_filter_never_success(session, field, value):
    session.records[0][field] = value
    assert tool.execute(USER, {})['error_code'] == 'ECUR_INVALID_RESPONSE'


def test_source_error_and_transport_failure_are_safe_and_close_session(session):
    session.report_status = 503
    assert tool.execute(USER, {})['error_code'] == 'ECUR_FETCH_FAILED'
    session.error = requests.RequestException('SECRET_NETWORK_ERROR')
    result = tool.execute(USER, {})
    assert result['error_code'] == 'ECUR_LOGIN_FAILED' and 'SECRET' not in str(result)
    assert session.closed


def test_rights_and_unsupported_parameters_precede_credential_access(monkeypatch):
    monkeypatch.setattr(tool, 'credentials', lambda *args: pytest.fail('unauthorized credential access'))
    for user in ({'modules': ['mingkh']}, {'modules': ['ecur'], 'is_active': False}):
        with pytest.raises(HTTPException): tool.execute(user, {})
    for params in ({'start': '2025-01-01'}, {'municipality': '<script>'}, {'municipality': ['Чехов']}, {'url': 'https://other.invalid'}):
        with pytest.raises(ValueError): tool.execute(USER, params)


def test_missing_credentials_does_not_use_another_users_ui_session(monkeypatch):
    monkeypatch.setattr(tool, 'credentials', lambda *args: None)
    monkeypatch.setattr(tool.requests, 'Session', lambda: pytest.fail('unexpected source request'))
    monkeypatch.setitem(client.STATE, 'email', 'PRIVATE_OTHER_USER')
    result = tool.execute(USER, {})
    assert result['error_code'] == 'ECUR_CREDENTIALS_MISSING' and 'PRIVATE' not in str(result)
