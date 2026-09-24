"""Contract tests for the supplied MINZHKH dashboard, using synthetic data."""
from datetime import date
import ssl
from unittest.mock import patch
import httpx
import pytest
from services.mingkh import client, dashboard


def test_windows_thursday_leap_and_validation():
    assert dashboard.resolve_periods({'preset':'thucur'},date(2026,9,22))==(5,0,12,7)
    assert dashboard.resolve_periods({'preset':'thuwed'},date(2026,9,22))==(12,6,19,13)
    assert client.year_ago(0,date(2024,2,29))==366
    p=client.periods(6,0,13,7)
    assert p['curr_period_start']=='6 day' and p['start_periods']=='6 day'
    assert p['curator']=='17' and p['spammers']=='0'
    with pytest.raises(ValueError):dashboard.resolve_periods({'preset':'custom','curr_from':'invalid','curr_to':'2026-09-22'})
    with pytest.raises(ValueError):dashboard.resolve_periods({'preset':'custom','curr_from':'2026-09-23','curr_to':'2026-09-22'},date(2026,9,22))


def test_fixed_portal_allowlist_verified_tls_and_no_redirects():
    observed=[]
    def handler(request):
        observed.append(request)
        return httpx.Response(200,json={'metadata':[{'colName':'ОМСУ'}],'resultset':[['Тест']]})
    original=httpx.Client
    def make_client(**kwargs):
        assert isinstance(kwargs['verify'],ssl.SSLContext) and kwargs['verify'].verify_mode==ssl.CERT_REQUIRED
        assert kwargs['follow_redirects'] is False and kwargs['trust_env'] is False
        return original(transport=httpx.MockTransport(handler),**kwargs)
    with patch.object(client.httpx,'Client',side_effect=make_client):
        columns,rows=client.Pentaho('test-user','test-password').query('q_download_detalization',{'curr_period_start':'6 day'})
    assert columns==['ОМСУ'] and rows==[['Тест']]
    url=observed[0].url
    assert url.host=='cur.bi.mosreg.ru' and url.params['paramcurr_period_start']=='6 day'
    assert 'test-password' not in str(url)
    assert url.params['path']=='/public/ЕЦУР_v3/ЕЦУР_api.cda'
    with pytest.raises(client.PortalError):client.Pentaho('x','y').query('arbitrary_query')


def test_redirect_never_receives_credentials():
    requests=[]
    def handler(request):
        requests.append(request)
        return httpx.Response(302,headers={'location':'https://example.invalid/collect'})
    original=httpx.Client
    with patch.object(client.httpx,'Client',side_effect=lambda **kw:original(transport=httpx.MockTransport(handler),**kw)):
        with pytest.raises(client.PortalError):client.Pentaho('x','secret').query('q_map')
    assert len(requests)==1


class Portal:
    def query(self,query,params,**kwargs):
        if query=='q_map':return ['title','population'],[['Тестовый округ',100000]]
        return list(dashboard.COLS.values()),[
            ['Тестовый округ','Портал','Вода','Сети','Утечка','Авария','Тест РСО'],
            ['Тестовый округ','Портал','Вне компетенции Ведомств МО','Сети','Утечка','Авария','Тест РСО']]
    def last_updated(self):return '2026-09-22 10:00',0


def test_dataset_filters_and_retains_all_three_periods():
    result=dashboard.build(Portal(),(6,0,13,7))
    assert result['counts']=={'curr':1,'prev':1,'appg':1}
    assert result['excluded']=={'curr':1,'prev':1,'appg':1}
    assert result['dims']['omsu']==['Тестовый округ']
    assert result['population']['Тестовый округ']==100000
    assert result['errors']=={}


def test_partial_comparison_is_not_presented_as_zero():
    class Partial(Portal):
        def query(self,query,params,**kwargs):
            if query=='q_download_detalization' and params['curr_period_start']=='13 day':raise client.PortalError('Unavailable')
            return super().query(query,params)
    with pytest.raises(client.PortalError,match='сравнения'):dashboard.build(Partial(),(6,0,13,7))


def test_cache_is_invalidated_when_credentials_change():
    dashboard.CACHE.clear()
    with patch.object(dashboard,'build',return_value={'counts':{'curr':2}}) as build:
        a=dashboard.get_dataset({'username':'user','password':'first'}, {})
        a['counts']['curr']=900
        assert dashboard.get_dataset({'username':'user','password':'first'}, {})['counts']['curr']==2
        dashboard.get_dataset({'username':'user','password':'second'}, {})
        assert build.call_count==2
    dashboard.CACHE.clear()


def test_monthly_chunks_cover_every_date_exactly_once():
    spans = list(dashboard.chunks(90, 0))
    assert spans == [(90, 61), (60, 31), (30, 1), (0, 0)]
    assert [day for start, end in spans for day in range(start, end-1, -1)] == list(range(90, -1, -1))


def test_chunk_merge_deduplicates_ids_but_keeps_rows_without_ids():
    columns = ['Внутренний Id', 'ОМСУ']
    result = dashboard.merge_chunks([
        (columns, [[1, 'Тест'], [None, 'Без ID']]),
        (columns[::-1], [['Тест', 1], ['Второй', 2], ['Ещё без ID', None]]),
    ])
    assert result == (columns, [[1, 'Тест'], [None, 'Без ID'], [2, 'Второй'], [None, 'Ещё без ID']])
    with pytest.raises(client.PortalError):
        dashboard.merge_chunks([(columns, [[1, 'Тест']]), (['Новый столбец'], [[2]])])


def test_only_transient_chunk_errors_are_retried():
    from unittest.mock import Mock
    portal = Mock()
    portal.query.side_effect = [client.PortalError('Temporary', retryable=True), (['ID'], [[1]])]
    with patch.object(dashboard.time, 'sleep') as sleep:
        assert dashboard.fetch_chunk(portal, {}, 29, 0) == (['ID'], [[1]])
        assert portal.query.call_count == 2
        sleep.assert_called_once_with(2)
    portal.reset_mock()
    portal.query.side_effect = client.PortalError('Invalid credentials')
    with patch.object(dashboard.time, 'sleep') as sleep:
        with pytest.raises(client.PortalError): dashboard.fetch_chunk(portal, {}, 29, 0)
        assert portal.query.call_count == 1
        sleep.assert_not_called()


def test_failed_chunk_invalidates_entire_period():
    class ChunkPortal:
        def query(self, query, params, **kwargs):
            if params['curr_period_start'] == '30 day':
                raise client.PortalError('Failed chunk')
            return ['Внутренний Id'], [[params['curr_period_start']]]
    data, errors = dashboard.fetch_periods(ChunkPortal(), {}, {'curr': (60, 0), 'prev': (10, 0)})
    assert 'curr' not in data and 'curr' in errors
    assert data['prev'][1] == [['10 day']]
