"""Dobrodel transport and collector tests, with synthetic responses and no network."""
from datetime import date, timedelta
from email.message import Message
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch
from urllib.parse import parse_qs, urlsplit

import requests
from requests.adapters import BaseAdapter

from services.edds import collector, dobrodel, runner


FORM = '<form id="loginForm"><input type="hidden" name="_csrf" value="csrf&amp;token"><input name="j_username"><input type="password" name="j_password"></form>'
REPORT = '<html><select id="curatorSelect"><option>МинЖКХ</option></select></html>'
ACCOUNT = {'username': 'synthetic-user', 'password': 'synthetic-password&with=encoding'}


def reply(body='', status=200, headers=None):
    return {'body': body if isinstance(body, (str, bytes)) else json.dumps(body), 'status': status, 'headers': headers or {}}


class ScriptedAdapter(BaseAdapter):
    """Real Session handles encoding/cookies; this adapter never opens a socket."""
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []
        self.closed = False

    def send(self, request, **options):
        self.calls.append((request, options))
        if not self.responses:
            raise AssertionError('Unexpected HTTP request')
        value = self.responses.pop(0)
        if isinstance(value, Exception):
            raise value
        response = requests.Response()
        response.status_code = value['status']
        response.url = request.url
        response.request = request
        response.headers.update(value['headers'])
        response.encoding = 'utf-8'
        content = value['body']
        response._content = content.encode() if isinstance(content, str) else content
        response._content_consumed = True
        message = Message()
        for key, value in response.headers.items():
            message.add_header(key, value)
        response.raw = SimpleNamespace(_original_response=SimpleNamespace(msg=message), release_conn=lambda: None)
        return response

    def close(self):
        self.closed = True


class DobrodelTests(unittest.TestCase):
    def setUp(self):
        environment = patch.dict(dobrodel.os.environ, {'DOBRODEL_CA_BUNDLE': ''})
        environment.start()
        self.addCleanup(environment.stop)

    def client(self, responses, account=None, *, deadline=None):
        client = dobrodel.DobrodelClient(**(account or ACCOUNT))
        client.deadline = deadline
        # Do not consult the test runner's proxy environment or make real calls.
        client.session.trust_env = False
        adapter = ScriptedAdapter(responses)
        client.session.mount('https://', adapter)
        client.session.mount('http://', adapter)
        self.addCleanup(client.session.close)
        return client, adapter

    def assert_code(self, code, callable, *args, **kwargs):
        with self.assertRaises(dobrodel.DobrodelError) as caught:
            callable(*args, **kwargs)
        self.assertEqual(caught.exception.code, code)
        self.assertNotIn(ACCOUNT['password'], str(caught.exception))
        self.assertNotIn('private-response-content', str(caught.exception))
        return caught.exception

    def test_ajax_login_uses_hidden_fields_and_same_cookie_session_then_fixed_report(self):
        client, adapter = self.client([
            reply(FORM, headers={'Set-Cookie':'JSESSIONID=prelogin; Path=/; Secure'}),
            reply({'show':'https://outside.invalid/private-response-content'}, headers={'Set-Cookie':'JSESSIONID=authenticated; Path=/; Secure'}),
            reply(REPORT), reply([{'cardId':1}]),
        ])
        with client:
            session = client.session
            client.login()
            self.assertEqual(client.fetch_all({'filters.curators':'МинЖКХ'}), [{'cardId':1}])
            self.assertIs(client.session, session)
        self.assertTrue(adapter.closed)
        requests_sent = [request for request, options in adapter.calls]
        self.assertEqual([(r.method,urlsplit(r.url).path) for r in requests_sent], [
            ('GET','/login'),('POST','/login/admin'),('GET','/OperativeReportGenerating'),('GET','/report/operative')])
        post = requests_sent[1]
        self.assertEqual(parse_qs(post.body), {'_csrf':['csrf&token'], 'j_username':[ACCOUNT['username']],
            'j_password':[ACCOUNT['password']], '_spring_security_remember_me':['on']})
        self.assertEqual(post.headers['X-Requested-With'], 'XMLHttpRequest')
        self.assertEqual(post.headers['Accept'], 'application/json')
        self.assertEqual(post.headers['Origin'], dobrodel.BASE)
        self.assertEqual(post.headers['Referer'], dobrodel.BASE+'/login')
        self.assertIn('JSESSIONID=prelogin', post.headers['Cookie'])
        for request in requests_sent[2:]:
            self.assertIn('JSESSIONID=authenticated', request.headers['Cookie'])
        for request, options in adapter.calls:
            self.assertEqual(options['verify'], True)
            self.assertEqual(options['timeout'], (15,240 if urlsplit(request.url).path == '/report/operative' else 60))
            self.assertTrue(options['stream'])
            self.assertNotIn('j_password', request.url)

    def test_custom_ca_bundle_keeps_certificate_verification(self):
        with patch.dict(dobrodel.os.environ, {'DOBRODEL_CA_BUNDLE':'/synthetic/trusted-ca.pem'}):
            client, adapter = self.client([reply('ok')])
        client.request('GET','/login')
        self.assertEqual(adapter.calls[0][1]['verify'], '/synthetic/trusted-ca.pem')

    def test_external_downgrade_userinfo_and_nonstandard_port_redirects_never_receive_credentials(self):
        for target in ['https://outside.invalid/login','//outside.invalid/login','http://admin.vmeste.mosreg.ru/login',
                       'https://admin.vmeste.mosreg.ru:8443/login','https://user:secret@admin.vmeste.mosreg.ru/login']:
            for status in [302,307,308]:
                with self.subTest(target=target, status=status):
                    client, adapter = self.client([reply('',status,{'Location':target})])
                    self.assert_code('redirect',client.request,'POST','/login/admin',data={'password':ACCOUNT['password']})
                    self.assertEqual(len(adapter.calls),1)
                    self.assertEqual(urlsplit(adapter.calls[0][0].url).hostname,'admin.vmeste.mosreg.ru')

    def test_invalid_destination_is_rejected_before_transport(self):
        for target in ['https://outside.invalid/', 'https://admin.vmeste.mosreg.ru:bad/', 'https://[malformed/']:
            with self.subTest(target=target):
                client, adapter = self.client([])
                self.assert_code('redirect', client.request, 'POST', target, data={'password':ACCOUNT['password']})
                self.assertEqual(adapter.calls, [])

    def test_same_origin_303_changes_post_to_get_and_drops_form_body(self):
        client, adapter = self.client([reply('',303,{'Location':'/OperativeReportGenerating'}),reply(REPORT)])
        client.request('POST','/login/admin',data={'password':ACCOUNT['password']})
        self.assertEqual(len(adapter.calls),2)
        second = adapter.calls[1][0]
        self.assertEqual(second.method,'GET')
        self.assertIsNone(second.body)
        self.assertNotIn('password',second.url)

    def test_redirect_loop_has_bounded_requests(self):
        client, adapter = self.client([reply('',302,{'Location':'/login'}) for _ in range(6)])
        self.assert_code('redirect',client.request,'GET','/login')
        self.assertEqual(len(adapter.calls),6)

    def test_missing_credentials_never_start_transport(self):
        for account in [{'username':'','password':'x'},{'username':'x','password':''}]:
            client, adapter = self.client([],account)
            self.assert_code('credentials',client.login)
            self.assertEqual(adapter.calls,[])

    def test_login_requires_expected_form_json_and_confirmed_report_permissions(self):
        cases = [([reply('<html>private-response-content</html>')], 'format'),
                 ([reply(FORM),reply('private-response-content')], 'format'),
                 ([reply(FORM),reply([])], 'format'),
                 ([reply(FORM),reply({'error':'private-response-content'})], 'login'),
                 ([reply(FORM),reply({}),reply('<input type="password">')], 'login'),
                 ([reply(FORM),reply({}),reply('<html>private-response-content</html>')], 'access')]
        for responses, code in cases:
            with self.subTest(code=code,responses=len(responses)):
                client, _ = self.client(responses)
                self.assert_code(code, client.login)

    def test_transport_errors_use_safe_fixed_codes(self):
        for error,code in [(requests.exceptions.SSLError('private-response-content'),'tls'),
                           (requests.exceptions.Timeout('private-response-content'),'timeout'),
                           (requests.exceptions.ConnectionError('private-response-content'),'network'),
                           (OSError('private-response-content'),'network')]:
            with self.subTest(code=code):
                client, _ = self.client([error])
                self.assert_code(code,client.request,'GET','/login')
        for status,code in [(401,'login'),(403,'login'),(503,'http')]:
            client, _ = self.client([reply('private-response-content',status)])
            error=self.assert_code(code,client.request,'GET','/login')
            self.assertEqual(error.status,status)

    def test_large_stream_is_rejected(self):
        client, _ = self.client([reply('123456')])
        with patch.object(dobrodel,'MAX_BYTES',5):
            self.assert_code('size',client.request,'GET','/login')

    def test_report_rejects_non_json_non_list_and_invalid_card_ids(self):
        invalid = ['<html>private-response-content</html>', {}, [1], [{'cardId':None}], [{'cardId':True}],
                   [{'cardId':0}], [{'cardId':-1}], [{'cardId':1.5}], [{'cardId':'１'}], [{'cardId':'١'}],
                   [{'cardId':'1e2'}], [{'cardId':'1/private-response-content'}], [{'cardId':1},{}]]
        for body in invalid:
            with self.subTest(body=body):
                client, _ = self.client([reply(body)])
                self.assert_code('format',client.fetch_all,{})
        client, _ = self.client([reply([{'cardId':'001'},{'cardId':2}])])
        self.assertEqual(len(client.fetch_all({})),2)

    def test_pagination_keeps_filters_and_collects_all_pages(self):
        client, adapter = self.client([reply([{'cardId':1},{'cardId':2}]),reply([{'cardId':3}])])
        filters={'filters.curators':'МинЖКХ','filters.statuses':'32,53','filters.createdBefore':'2026-10-02'}
        with patch.object(dobrodel,'PAGE_SIZE',2):
            self.assertEqual([row['cardId'] for row in client.fetch_all(filters)],[1,2,3])
        for page,(request,_) in enumerate(adapter.calls):
            self.assertEqual(parse_qs(urlsplit(request.url).query),
                {'orderBy':['ID'],'page':[str(page)],'size':['2'],**{key:[value] for key,value in filters.items()}})
        self.assertNotIn('page',filters)

    def test_empty_repeated_and_too_many_pages_are_bounded(self):
        client, adapter = self.client([reply([])])
        self.assertEqual(client.fetch_all({}),[])
        self.assertEqual(len(adapter.calls),1)
        client, _ = self.client([reply([{'cardId':1}]),reply([{'cardId':1}])])
        with patch.object(dobrodel,'PAGE_SIZE',1):
            self.assert_code('format',client.fetch_all,{})
        client, adapter = self.client([reply([{'cardId':i+1}]) for i in range(600)])
        with patch.object(dobrodel,'PAGE_SIZE',1):
            self.assert_code('size',client.fetch_all,{})
        self.assertEqual(len(adapter.calls),600)

    def test_expired_session_reauthenticates_once_and_retries_same_page(self):
        client, adapter = self.client([reply('',401),reply(FORM),reply({}),reply(REPORT),reply([{'cardId':7}])])
        filters={'filters.statuses':'32'}
        self.assertEqual(client.fetch_all(filters),[{'cardId':7}])
        self.assertEqual(adapter.calls[0][0].url,adapter.calls[-1][0].url)
        self.assertEqual(sum(r.method=='POST' for r,_ in adapter.calls),1)
        client, adapter = self.client([reply('',401),reply(FORM),reply({}),reply(REPORT),reply('',401)])
        self.assert_code('login',client.fetch_all,filters)
        self.assertEqual(sum(r.method=='POST' for r,_ in adapter.calls),1)

    def test_optional_deadline_keeps_default_timeout_and_caps_remaining_budget(self):
        with patch('services.edds.dobrodel.time.monotonic', return_value=100):
            for deadline in [None, 103.5]:
                with self.subTest(deadline=deadline):
                    client, adapter = self.client([reply('ok')], deadline=deadline)
                    self.assertEqual(client.request('GET', '/login').text, 'ok')
                    timeouts = adapter.calls[0][1]['timeout']
                    if deadline is None:
                        self.assertEqual(timeouts, (15, 60))
                    else:
                        self.assertEqual(len(timeouts), 2)
                        self.assertTrue(all(0 < value <= 3.5 for value in timeouts))

    def test_expired_deadline_stops_before_sending_credentials(self):
        with patch('services.edds.dobrodel.time.monotonic', return_value=100):
            for deadline in [99, 100]:
                with self.subTest(deadline=deadline):
                    client, adapter = self.client([], deadline=deadline)
                    self.assert_code('timeout', client.request, 'POST', '/login/admin',
                                     data={'j_password': ACCOUNT['password']})
                    self.assertEqual(adapter.calls, [])

    def test_redirect_follow_cannot_restart_expired_budget(self):
        clock = {'now': 100}
        client, adapter = self.client([reply('', 302, {'Location': '/OperativeReportGenerating'})], deadline=105)
        original_send = adapter.send
        def send(request, **options):
            response = original_send(request, **options)
            clock['now'] = 106
            return response
        with patch('services.edds.dobrodel.time.monotonic', side_effect=lambda: clock['now']), \
             patch.object(adapter, 'send', side_effect=send):
            self.assert_code('timeout', client.request, 'POST', '/login/admin',
                             data={'j_password': ACCOUNT['password']})
        self.assertEqual(len(adapter.calls), 1)

    def test_deadline_is_checked_during_streaming_and_response_is_closed(self):
        clock = {'now': 100}
        client, adapter = self.client([reply('')], deadline=105)
        original_send = adapter.send
        closed = []
        def chunks(_chunk_size):
            yield b'first chunk'
            clock['now'] = 106
            yield b'private-response-content'
        def send(request, **options):
            response = original_send(request, **options)
            response.iter_content = chunks
            original_close = response.close
            def close():
                closed.append(True)
                original_close()
            response.close = close
            return response
        with patch('services.edds.dobrodel.time.monotonic', side_effect=lambda: clock['now']), \
             patch.object(adapter, 'send', side_effect=send):
            self.assert_code('timeout', client.request, 'GET', '/report/operative')
        self.assertEqual(len(adapter.calls), 1)
        self.assertEqual(closed, [True])

    def test_login_requests_share_one_deadline(self):
        clock = {'now': 100}
        client, adapter = self.client([reply(FORM), reply({}), reply(REPORT)], deadline=105)
        original_send = adapter.send
        def send(request, **options):
            response = original_send(request, **options)
            clock['now'] += 2
            return response
        with patch('services.edds.dobrodel.time.monotonic', side_effect=lambda: clock['now']), \
             patch.object(adapter, 'send', side_effect=send):
            self.assert_code('timeout', client.login)
        self.assertEqual(len(adapter.calls), 3)
        read_timeouts = [options['timeout'][1] for _, options in adapter.calls]
        self.assertGreater(read_timeouts[0], read_timeouts[1])
        self.assertGreater(read_timeouts[1], read_timeouts[2])
        self.assertLessEqual(read_timeouts[-1], 1)

    def test_card_id_length_limit_returns_format_instead_of_integer_conversion_error(self):
        for length in [33, 5000]:
            with self.subTest(length=length):
                client, _ = self.client([reply([{'cardId': '9' * length}])])
                self.assert_code('format', client.fetch_all, {})
        client, _ = self.client([reply([{'cardId': '9' * 32}])])
        self.assertEqual(client.fetch_all({}), [{'cardId': '9' * 32}])

    def test_runner_diagnostics_never_expose_transport_output(self):
        for code in dobrodel.ERROR_MESSAGES:
            text=runner.failure_message(f'Trace private-response-content {ACCOUNT["password"]}\nDOBRODEL_ERROR:{code}\n')
            self.assertEqual(text,str(dobrodel.DobrodelError(code)))
            self.assertNotIn('private-response-content',text)
            self.assertNotIn(ACCOUNT['password'],text)
        self.assertIn('503',runner.failure_message('DOBRODEL_ERROR:http:503 private-response-content'))
        self.assertNotIn('private-response-content',runner.failure_message('DOBRODEL_ERROR:unknown private-response-content'))


class CollectorTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path=Path(self.temp.name)/'water_daily.json'
        for name,value in [('WATER_JSON',self.path),('KEEP_DAYS',0)]:
            patcher=patch.object(collector,name,value)
            patcher.start();self.addCleanup(patcher.stop)
        logger=patch.object(collector,'log')
        self.log=logger.start()
        self.addCleanup(logger.stop)
        self.windows=[(date(2026,9,30),date(2026,10,1),'test')]

    def run_collector(self,batches,windows=None):
        client=MagicMock()
        client.__enter__.return_value=client
        client.fetch_all.side_effect=batches
        with patch.object(collector,'credentials',return_value=ACCOUNT) as credentials, \
             patch.object(collector,'DobrodelClient',return_value=client) as factory, \
             patch.object(collector,'pull_windows',return_value=self.windows if windows is None else windows), \
             patch.dict(collector.os.environ,{'EDDS_CHROME_EXECUTABLE':'/missing/browser'}):
            collector.main()
        credentials.assert_called_once_with('edds')
        factory.assert_called_once_with(**ACCOUNT)
        client.login.assert_called_once()
        client.__exit__.assert_called_once()
        return client

    def complaint(self, card_id, created):
        return {'cardId':card_id,'created':created,'district':'Тестовый округ','subcategory':'Холодное водоснабжение'}

    def test_requests_client_includes_last_day_and_filters_outside_dates(self):
        rows=[self.complaint(1,'2026-09-29T23:59:59'),self.complaint(2,'2026-09-30T00:00:00'),
              self.complaint(3,'01.10.2026 23:59:59'),self.complaint(4,'2026-10-02T00:00:00'),
              self.complaint(5,None)]
        client=self.run_collector([rows])
        client.fetch_all.assert_called_once_with({'filters.curators':collector.CURATOR,'filters.statuses':collector.STATUSES,
            'filters.createdAfter':'2026-09-30','filters.createdBefore':'2026-10-02'})
        data=json.loads(self.path.read_text())
        self.assertEqual(data['days'],{'2026-09-30':{'Тестовый округ':[1,0,0]},'2026-10-01':{'Тестовый округ':[1,0,0]}})
        self.assertNotIn(ACCOUNT['password'],self.path.read_text())

    def test_empty_all_windows_preserves_previous_summary(self):
        original=json.dumps({'days':{'2026-09-30':{'Тестовый округ':[9,0,0]}}})
        self.path.write_text(original)
        with self.assertRaisesRegex(RuntimeError,'пустой отчёт'):
            self.run_collector([[]])
        self.assertEqual(self.path.read_text(),original)

    def test_later_window_failure_keeps_successfully_updated_interval(self):
        original=json.dumps({'days':{'2026-09-30':{'Тестовый округ':[9,0,0]}}})
        self.path.write_text(original)
        windows=[(date(2026,9,30),date(2026,9,30),'first'),(date(2026,10,1),date(2026,10,1),'second')]
        with self.assertRaises(dobrodel.DobrodelError):
            self.run_collector([[self.complaint(1,'2026-09-30')],dobrodel.DobrodelError('timeout')],windows)
        self.assertEqual(json.loads(self.path.read_text())['days'], {'2026-09-30':{'Тестовый округ':[1,0,0]}})

    def test_successful_empty_window_clears_stale_counts_when_another_window_has_data(self):
        self.path.write_text(json.dumps({'days':{'2026-09-29':{'Old':[3,0,0]},'2026-09-30':{'Old':[9,0,0]}}}))
        windows=[(date(2026,9,30),date(2026,9,30),'first'),(date(2026,10,1),date(2026,10,1),'second')]
        self.run_collector([[],[self.complaint(1,'2026-10-01')]],windows)
        self.assertEqual(json.loads(self.path.read_text())['days'],{'2026-09-29':{'Old':[3,0,0]},'2026-09-30':{},'2026-10-01':{'Тестовый округ':[1,0,0]}})

    def test_successful_window_includes_empty_boundary_days_in_coverage(self):
        windows = [(date(2026,9,28), date(2026,10,1), 'test')]
        self.run_collector([[self.complaint(1, '2026-09-30')]], windows)
        data = json.loads(self.path.read_text())
        self.assertEqual(data['from'], '2026-09-28')
        self.assertEqual(data['to'], '2026-10-01')
        self.assertEqual(data['days'], {
            '2026-09-28': {}, '2026-09-29': {},
            '2026-09-30': {'Тестовый округ': [1,0,0]}, '2026-10-01': {},
        })
        self.assertEqual(data['window'], {'from':'2026-09-28', 'to':'2026-10-01'})

    def test_empty_successful_backfill_advances_the_next_query(self):
        today = date.today()
        first = today - timedelta(days=5)
        back_start = first - timedelta(days=30)
        self.path.write_text(json.dumps({'from':first.isoformat(), 'to':today.isoformat(),
            'days':{first.isoformat():{'Old':[3,0,0]},today.isoformat():{'Old':[9,0,0]}}}))
        windows = [(today, today, 'tail'), (back_start, first-timedelta(days=1), 'backfill')]
        self.run_collector([[self.complaint(1, today.isoformat())], []], windows)
        data = json.loads(self.path.read_text())
        self.assertEqual(data['from'], back_start.isoformat())
        self.assertEqual(data['to'], today.isoformat())
        self.assertEqual(data['days'][first.isoformat()], {'Old':[3,0,0]})
        self.assertEqual(data['days'][back_start.isoformat()], {})
        with patch.object(collector, 'DAYS_BACK', 100), patch.object(collector, 'MAX_CATCHUP', 30):
            following = collector.pull_windows(data)
        older = [(a,b) for a,b,reason in following if reason.startswith('добор назад')]
        self.assertTrue(older)
        self.assertTrue(all(b < back_start for _,b in older))
        self.assertEqual(older[0][1], back_start-timedelta(days=1))

    def test_legacy_sparse_summary_does_not_gain_unqueried_empty_days(self):
        original = {'from':'2026-09-10', 'to':'2026-09-28',
            'window':{'from':'2026-09-01', 'to':'2026-09-29'},
            'days':{'2026-09-10':{'Old':[3,0,0]}, '2026-09-28':{'Old':[9,0,0]}}}
        self.path.write_text(json.dumps(original))
        self.run_collector([[self.complaint(1, '2026-09-30')]])
        data = json.loads(self.path.read_text())
        self.assertEqual(data['from'], '2026-09-10')
        self.assertEqual(data['to'], '2026-10-01')
        self.assertEqual(data['days'], {
            **original['days'], '2026-09-30':{'Тестовый округ':[1,0,0]}, '2026-10-01':{},
        })
        self.assertNotIn('2026-09-01', data['days'])
        self.assertNotIn('2026-09-29', data['days'])

    def test_retention_prunes_empty_coverage_and_nonzero_days_together(self):
        today = date.today()
        start = today-timedelta(days=3)
        rows = collector.build_rows([self.complaint(1, start.isoformat()), self.complaint(2, today.isoformat())])
        with patch.object(collector, 'KEEP_DAYS', 1):
            collector.write_water_daily(rows, start.isoformat(), today.isoformat(), None)
        data = json.loads(self.path.read_text())
        self.assertEqual(data['from'], (today-timedelta(days=1)).isoformat())
        self.assertEqual(data['to'], today.isoformat())
        self.assertEqual(data['days'], {
            (today-timedelta(days=1)).isoformat():{}, today.isoformat():{'Тестовый округ':[1,0,0]},
        })

    def test_missing_or_unreadable_credentials_stop_with_safe_code(self):
        for result,code in [(None,'credentials'),(RuntimeError('private-response-content'),'config')]:
            with patch.object(collector,'credentials',side_effect=result if isinstance(result,Exception) else None,return_value=result) as credentials, \
                 patch.object(collector,'DobrodelClient') as factory:
                with self.assertRaises(dobrodel.DobrodelError) as error:
                    collector.main()
                self.assertEqual(error.exception.code,code)
                self.assertNotIn('private-response-content',str(error.exception))
                credentials.assert_called_once_with('edds')
                factory.assert_not_called()

    def test_first_run_queries_bounded_contiguous_periods(self):
        with patch.object(collector,'DAYS_BACK',730),patch.object(collector,'MAX_CATCHUP',30):
            windows=collector.pull_windows(None)
        self.assertEqual(len(windows),5)
        self.assertEqual(windows[0][1],date.today())
        self.assertEqual((windows[-1][0]-windows[0][1]).days,-29)
        for previous,following in zip(windows,windows[1:]):
            self.assertEqual(following[1].toordinal()+1,previous[0].toordinal())
        self.assertTrue(all((end-start).days<collector.MAX_WINDOW_DAYS for start,end,_ in windows))


if __name__=='__main__':
    unittest.main()
