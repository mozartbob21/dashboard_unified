"""Synthetic, offline contracts for the integrated SMS collectors and shared storage."""
from concurrent.futures import ThreadPoolExecutor
from datetime import date
import json
from pathlib import Path
import ssl
from unittest.mock import patch

import httpx
import pytest

from services.mingkh import client as base_client
from services.mingkh.sms import client, context, dataset, rez_data, rm_data, svod_obrashcheniya as svod


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setattr(context, 'today', lambda: date(2026, 10, 19))
    monkeypatch.setattr(svod, 'local_today', lambda: date(2026, 10, 19))
    monkeypatch.setattr(rez_data, 'local_today', lambda: date(2026, 10, 19))
    monkeypatch.setattr(svod, 'DATA', tmp_path)
    monkeypatch.setattr(rez_data, 'STORE', str(tmp_path / 'rez_rows'))
    monkeypatch.setattr(rez_data, 'FOUND', str(tmp_path / 'rez_found.json'))
    monkeypatch.setattr(dataset, '_working', {'query': None})
    monkeypatch.setattr(client, 'credentials', lambda service: {'username': 'synthetic', 'password': 'synthetic-secret'})
    # A missed mock must fail rather than contact any real portal.
    monkeypatch.setattr(httpx.HTTPTransport, 'handle_request', lambda *a, **kw: pytest.fail('Live network is forbidden'))


def test_month_windows_leap_and_last_completed_wednesday():
    assert svod.month_windows(date(2024, 2, 27), date(2024, 3, 2)) == [
        (2024, 2, date(2024, 2, 27), date(2024, 2, 29)), (2024, 3, date(2024, 3, 1), date(2024, 3, 2))]
    assert svod._year_back(date(2024, 2, 29)) == date(2023, 2, 28)
    assert svod.last_wednesday(date(2026, 10, 14)) == date(2026, 10, 7)
    assert svod.last_wednesday(date(2026, 10, 15)) == date(2026, 10, 14)


def test_concurrent_dictionary_fill_queries_once(monkeypatch):
    import time
    calls=[]
    class Portal:
        def periods(self, *args): return {}
        def query(self, *args, **kwargs):
            calls.append(1)
            time.sleep(.05)
            return [], [['1','Власиха']]
    monkeypatch.setattr(svod, 'Pentaho', Portal)
    with ThreadPoolExecutor(max_workers=3) as pool:
        results=list(pool.map(lambda _: svod.omsu_list(), range(3)))
    assert calls == [1]
    assert all(rows == [('1','Власиха')] for rows in results)


def test_build_deduplicates_sources_and_atomic_archive(monkeypatch, tmp_path):
    calls = []
    def collect(start, end, flt, log):
        calls.append(flt)
        count = 3 if flt['source'] == '4' else 5
        rows = {'419': {'title': 'Ржавая вода', 'n': count}}
        return {'date_from': start.isoformat(), 'date_to': end.isoformat(), 'filters': flt,
                'portal_updated': '2026-10-19', 'months': [dict(year=2026, month=10, **{'from':start.isoformat(), 'to':end.isoformat(), 'curr':rows, 'prev':rows})],
                'ytd': {'curr': rows, 'prev': rows}}
    monkeypatch.setattr(svod, 'collect', collect)
    result = svod.build(date(2026, 10, 1), date(2026, 10, 7), ['4', '4', '1'])
    assert len(calls) == 2 and result['ytd']['curr']['419']['n'] == 8
    files = list(tmp_path.glob('obrashcheniya_*.json'))
    assert len(files) == 2 and not list(tmp_path.glob('.*'))
    assert context.json_read(files[0])['portal_updated'] == '2026-10-19'


@pytest.mark.parametrize('flt', [{'omsu': '../secret'}, {'source': '4/../../secret'}, {'password': 'x'}])
def test_filter_path_injection_rejected(flt):
    with pytest.raises(ValueError):
        svod.load(date(2026, 10, 1), date(2026, 10, 7), flt)


def test_summary_keeps_previous_only_topics_and_excludes_noncompetence():
    current = {'419': {'title': 'Ржавая вода', 'n': 200}, '48': {'title': 'Вне компетенции (МИНЖКХ)', 'n': 10000}}
    previous = {'419': {'title': 'Ржавая вода', 'n': 400}, '221': {'title': 'Мусор в МКД', 'n': 100}}
    data = {'date_from': '2026-01-01', 'date_to': '2026-01-31', 'ytd': {'curr': current, 'prev': previous},
            'months': [{'year': 2026, 'month': 1, 'curr': current, 'prev': previous}]}
    text = svod.render(data)
    assert 'Всего: 0,5 тыс → 0,2 тыс (−60%)' in text
    assert 'Мусор в МКД' in text and 'Вне компетенции' not in text


def test_daily_top_five_plus_other_keeps_total_and_weekday_comparison(monkeypatch):
    calls = []
    def fetch(p, start, end, today, filters):
        calls.append((start, end))
        return {str(i): {'title': f'Тема {i}', 'n': i} for i in range(1, 8)}
    monkeypatch.setattr(svod, 'fetch_period', fetch)
    result = svod.daily(date(2026, 10, 1), date(2026, 10, 8))
    assert calls == [(date(2026, 10, 1), date(2026, 10, 8)), (date(2026, 9, 17), date(2026, 9, 24))]
    assert len(result['rows']) == 6 and result['rows'][-1]['title'] == 'Прочие вопросы'
    assert sum(r['curr'] for r in result['rows']) == 28
    assert 'Поступило: 28 обращений' in svod.render_daily(result)


def mock_http(monkeypatch, handler):
    original = httpx.Client
    def factory(**kwargs):
        assert isinstance(kwargs['verify'], ssl.SSLContext)
        assert kwargs['verify'].verify_mode == ssl.CERT_REQUIRED
        assert kwargs['trust_env'] is False and kwargs['follow_redirects'] is False
        return original(transport=httpx.MockTransport(handler), **kwargs)
    monkeypatch.setattr(httpx, 'Client', factory)


def test_portal_uses_saved_mingkh_access_and_scoped_allowlist(monkeypatch):
    seen = []
    def credentials(name):
        seen.append(name)
        return {'username': 'synthetic', 'password': 'synthetic-secret'}
    monkeypatch.setattr(client, 'credentials', credentials)
    requests = []
    def handler(request):
        requests.append(request)
        return httpx.Response(200, json={'metadata': [{'colName':'title'}], 'resultset':[['Вода']]})
    mock_http(monkeypatch, handler)
    portal = client.Pentaho()
    assert portal.query('q_top_subtopic') == (['title'], [['Вода']])
    assert seen == ['mingkh'] and requests[0].url.host == 'cur.bi.mosreg.ru'
    assert 'synthetic-secret' not in str(requests[0].url)
    with pytest.raises(base_client.PortalError):
        portal.query('arbitrary')
    with pytest.raises(base_client.PortalError):
        base_client.Pentaho('x', 'y').query('q_top_subtopic')


def test_portal_missing_credentials_is_safe(monkeypatch):
    monkeypatch.setattr(client, 'credentials', lambda _: None)
    with pytest.raises(base_client.PortalError, match='Администратору'):
        client.Pentaho()


@pytest.mark.parametrize('bad', [{'curator': '1'}, {'source': '1,4'}, {'omsu': '../secret'}])
def test_portal_filter_override_rejected(bad):
    with pytest.raises(ValueError):
        client.Pentaho.periods(6, 0, 13, 7, **bad)


def test_fallback_501_and_reordered_columns_are_preserved():
    calls = []
    class Portal:
        def query(self, name, params, **kw):
            calls.append(name)
            if name == 'q_download_detalization':
                raise base_client.PortalError('HTTP 501', status=501, retryable=True)
            return ['Внутренний Id', 'Адрес'], [[1, 'Округ, дом 1']]
    assert dataset._fetch_one(Portal(), {}, 6, 0)[1] == [[1, 'Округ, дом 1']]
    assert calls == ['q_download_detalization', 'q_download_detalization_address']
    columns, rows = dataset._merge([
        (['Внутренний Id', 'Адрес'], [[1, 'Первый']]),
        (['Адрес', 'Внутренний Id', 'Новая колонка'], [['Первый', 1, 'x'], ['Второй', 2, 'y']])])
    assert columns == ['Внутренний Id', 'Адрес', 'Новая колонка']
    assert rows == [[1, 'Первый', None], [2, 'Второй', 'y']]


def test_auth_failures_do_not_repeat_or_try_fallback():
    calls = []
    class Portal:
        def query(self, *args, **kwargs):
            calls.append(args)
            raise base_client.PortalError('Нет доступа')
    with pytest.raises(base_client.PortalError):
        dataset._fetch_one(Portal(), {}, 6, 0)
    assert len(calls) == 1


def test_month_cache_is_reused_only_when_complete(monkeypatch):
    payload = {'fields': rez_data.FIELDS, 'rows': [], 'fetched_at': '2026-10-08T10:00:00+03:00'}
    context.atomic_json(rez_data._path(2026, 9), payload)
    assert rez_data._complete(2026, 9)
    payload['fetched_at'] = '2026-10-01T10:00:00+03:00'
    context.atomic_json(rez_data._path(2026, 9), payload)
    assert not rez_data._complete(2026, 9)


def test_rez_scope_filters_only_periods_history_includes_other_sources(monkeypatch):
    monkeypatch.setattr(rez_data, 'ensure', lambda *args: 1)
    monkeypatch.setattr(svod, 'omsu_list', lambda: [('1', 'Округ')])
    def row(ident, when, source='Добродел', theme='Ржавая вода', addr='Округ, дом 1'):
        result = {k: '' for k in rez_data.FIELDS}
        result.update(innid=ident, date=when, src=source, subtopic=theme, addr=addr, omsu='Округ')
        return result
    monkeypatch.setattr(rez_data, 'rows_between', lambda *args: [
        row(1, '2026-10-14'), row(2, '2026-10-07'), row(3, '2026-01-01', 'ПОС'),
        row(4, '2026-10-14', theme='Вне компетенции (МИНЖКХ)'),
        row(5, '2026-01-01', addr='Другой адрес'), row(6, '2026-10-14', 'ПОС')])
    result = rez_data.collect({'from':'2026-10-08','to':'2026-10-14','prev_from':'2026-10-01','prev_to':'2026-10-07',
                               'subtopics':['419'],'sources':['4'],'omsu':'1'})
    ii = rez_data.FIELDS.index('innid')
    assert [row[ii] for row in result['current']] == [1]
    assert [row[ii] for row in result['prev']] == [2]
    assert [row[ii] for row in result['history']] == [3, 6]


def case(key='test', **changes):
    return {'key':key, 'addr':'Тестовый округ, дом 1', 'omsu':'Тестовый округ', 'score':5, **changes}


def test_found_upsert_is_idempotent_and_retains_earliest_and_latest():
    assert rez_data.found_upsert('2026-10-08', '2026-10-14', [case()]) == {'added':1, 'total':1}
    assert rez_data.found_upsert('2026-10-08', '2026-10-14', [case()]) == {'added':0, 'total':1}
    rez_data.found_upsert('2026-10-01', '2026-10-07', [case()])
    result = rez_data.found_load()['test']
    assert result['first_from'] == '2026-10-01' and result['last_to'] == '2026-10-14'


def test_found_concurrent_requests_never_lose_cases():
    def write(index):
        return rez_data.found_upsert('2026-10-08', '2026-10-14', [case(str(index))])
    with ThreadPoolExecutor(max_workers=8) as executor:
        list(executor.map(write, range(32)))
    assert len(rez_data.found_load()) == 32
    assert not list(Path(rez_data.FOUND).parent.glob('.rez_found.json-*'))


@pytest.mark.parametrize('bad', [case(score=float('nan')), case(score=True), case(score=-1), case(score=1000001),
                                  case(key=''), case(addr='x'*1501), case(omsu='x'*301), case(path='/tmp/file')])
def test_found_invalid_payload_keeps_previous_data(bad):
    rez_data.found_upsert('2026-10-08', '2026-10-14', [case()])
    before = Path(rez_data.FOUND).read_bytes()
    with pytest.raises(ValueError):
        rez_data.found_upsert('2026-10-08', '2026-10-14', [case('valid'), bad])
    assert Path(rez_data.FOUND).read_bytes() == before


def test_atomic_storage_keeps_last_snapshot_on_failure(tmp_path, monkeypatch):
    path = tmp_path / 'test.json'
    context.atomic_json(path, {'previous': True})
    before = path.read_bytes()
    monkeypatch.setattr(context.os, 'replace', lambda *args: (_ for _ in ()).throw(OSError('test')))
    with pytest.raises(OSError):
        context.atomic_json(path, {'new': True})
    assert path.read_bytes() == before and not list(tmp_path.glob('.test.json-*'))


def test_nonblocking_process_lock_conflict_releases_after_exit(tmp_path):
    path = tmp_path / 'test.lock'
    with context.process_lock(path):
        with pytest.raises(ValueError, match='обновляются'):
            with context.process_lock(path, timeout=0):
                pass
    with context.process_lock(path, timeout=0):
        pass


def test_rm_uses_current_server_key_without_redirect_or_secret_in_url(monkeypatch):
    monkeypatch.setenv('MGKH_API_KEY', 'synthetic-rm-secret')
    requests = []
    def handler(request):
        requests.append(request)
        assert request.headers['X-Redmine-API-Key'] == 'synthetic-rm-secret'
        return httpx.Response(200, json={'issues':[], 'total_count':0})
    mock_http(monkeypatch, handler)
    assert rm_data.issues() == []
    assert requests[0].url.host == 'mgkh.rm.mosreg.ru'
    assert 'synthetic-rm-secret' not in str(requests[0].url)


def test_rm_redirect_does_not_forward_key(monkeypatch):
    monkeypatch.setenv('MGKH_API_KEY', 'synthetic-rm-secret')
    seen = []
    def handler(request):
        seen.append(request)
        return httpx.Response(302, headers={'location':'https://example.invalid/steal'})
    mock_http(monkeypatch, handler)
    with pytest.raises(rm_data.RedmineError, match='не принял'):
        rm_data.issues()
    assert len(seen) == 1


def test_rm_bad_path_or_missing_key_never_connects(monkeypatch):
    with pytest.raises(ValueError):
        rm_data._get('//other.invalid')
    monkeypatch.delenv('MGKH_API_KEY', raising=False)
    with pytest.raises(rm_data.RedmineError, match='MGKH_API_KEY'):
        rm_data.issues()


def test_rm_response_size_bounded(monkeypatch):
    monkeypatch.setenv('MGKH_API_KEY', 'synthetic-rm-secret')
    monkeypatch.setattr(rm_data, 'MAX_RESPONSE_BYTES', 20)
    mock_http(monkeypatch, lambda request: httpx.Response(200, content=b'x'*21))
    with pytest.raises(rm_data.RedmineError, match='большой'):
        rm_data.issues()


def test_rm_measures_uses_event_dates_not_issue_dates_and_excludes_confirmation(monkeypatch):
    monkeypatch.setattr(rm_data, 'issues', lambda: [{'id':1, 'subject':'Дом 1', 'assigned_to':{'name':'Округ'},
        'status':{'name':'Открыто'}, 'custom_fields':[{'id':43,'value':'2026-10-01'}, {'id':229,'value':'2026-10-07'},
        {'id':51,'value':'2026-10-01'}, {'id':45,'value':'2026-09-01'}, {'id':221,'value':'334'}]}])
    result = rm_data.measures('2026-10-01', '2026-10-07')
    assert len(result['measures']) == 2
    assert [item['group'] for item in result['measures']] == ['quick', 'system']
    assert all(item['url'] == 'https://mgkh.rm.mosreg.ru/issues/1' for item in result['measures'])
    assert result['measures'][0]['type'] == 'Резонансные'


def test_rm_missing_page_is_error_not_partial_total(monkeypatch):
    monkeypatch.setattr(rm_data, '_get', lambda path: {'issues': [{'id':1}], 'total_count':101})
    with pytest.raises(rm_data.RedmineError, match='изменился'):
        rm_data.issues()


def test_omsu_dictionary_has_short_timeout_and_server_cache(monkeypatch):
    observed = []
    class Portal:
        periods = staticmethod(lambda *args: {})
        def query(self, query, params, **kwargs):
            observed.append(kwargs)
            return ['id','title'], [['1','Тестовый округ'], ['-1','Все']]
    monkeypatch.setattr(svod, 'Pentaho', Portal)
    assert svod.omsu_list() == [('1', 'Тестовый округ')]
    assert svod.omsu_list() == [('1', 'Тестовый округ')]
    assert observed == [{'timeout':15}]


def test_found_archive_merge_preserves_earliest_and_latest_observation():
    rez_data.found_upsert('2026-10-08', '2026-10-14', [case()])
    raw = {'test': {'addr':'Тестовый округ, дом 1','omsu':'Тестовый округ','score':5,
                    'first_from':'2026-09-01', 'first_to':'2026-09-07'}}
    assert rez_data.merge_found(raw) == {'total':1}
    assert rez_data.found_load()['test']['first_from'] == '2026-09-01'
    assert rez_data.found_load()['test']['last_to'] == '2026-10-14'
    assert 'last_to' not in raw['test']  # Validation/merge does not mutate the import.


def test_found_archive_validation_rejects_future_dates_and_unknown_fields():
    value = {'x': {'addr':'Тестовый дом','omsu':'Округ','score':5,'first_from':'2026-09-01','first_to':'2026-09-07'}}
    assert rez_data.validate_found_data(value)['x']['score'] == 5
    value['x']['file_path'] = '../other.json'
    with pytest.raises(ValueError):
        rez_data.validate_found_data(value)
    del value['x']['file_path']
    value['x']['first_to'] = '2027-01-01'
    with pytest.raises(ValueError):
        rez_data.validate_found_data(value)
