"""Independent recovery regressions: synthetic HTTP, isolated files, no real portal."""
from datetime import date, timedelta
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch
from urllib.parse import parse_qs, urlsplit

import requests
from requests.adapters import BaseAdapter

from services.edds import collector, dobrodel, runner

ACCOUNT = {'username': 'recovery-fixture-user', 'password': 'fixture-password-never-log'}
PRIVATE = 'private complaint text that must never enter progress'
FORM = '<form id="loginForm"><input name="j_username"><input type="password" name="j_password"></form>'
REPORT = '<select id="curatorSelect"><option>test</option></select>'


def row(cid, created='2026-10-01'):
    return {'cardId': cid, 'created': created, 'district': 'Synthetic district',
            'subcategory': 'Холодное водоснабжение', 'body': PRIVATE}


def filters(start='2026-10-01', stop='2026-10-05'):
    return {'filters.curators': collector.CURATOR, 'filters.statuses': collector.STATUSES,
            'filters.createdAfter': start, 'filters.createdBefore': stop}


def query(request):
    return {key: values[0] for key, values in parse_qs(urlsplit(request.url).query).items()}


class LocalAdapter(BaseAdapter):
    def __init__(self, respond):
        self.respond = respond
        self.calls = []

    def send(self, request, **options):
        self.calls.append((request, options))
        status, body = self.respond(request, len(self.calls))
        response = requests.Response()
        response.status_code = status
        response.url = request.url
        response.request = request
        response.encoding = 'utf-8'
        response._content = (body if isinstance(body, str) else json.dumps(body)).encode()
        response._content_consumed = True
        return response

    def close(self):
        pass


class RecoveryTests(unittest.TestCase):
    def client(self, respond):
        client = dobrodel.DobrodelClient(**ACCOUNT)
        client.session.trust_env = False
        adapter = LocalAdapter(respond)
        # Both schemes intercepted: any accidental destination is still offline.
        client.session.mount('https://', adapter)
        client.session.mount('http://', adapter)
        self.addCleanup(client.session.close)
        return client, adapter

    def test_transient_report_failures_split_complete_intervals_keep_filters_and_deduplicate_boundaries(self):
        for fault in ('timeout', 'size', 502, 503, 504):
            with self.subTest(fault=fault):
                original = filters()
                def respond(request, attempt):
                    params = query(request)
                    self.assertEqual(params['filters.curators'], collector.CURATOR)
                    self.assertEqual(params['filters.statuses'], collector.STATUSES)
                    self.assertEqual(params['orderBy'], 'ID')
                    self.assertEqual(params['page'], '0')
                    span = params['filters.createdAfter'], params['filters.createdBefore']
                    if span == ('2026-10-01', '2026-10-05'):
                        if fault == 'timeout': raise requests.exceptions.ReadTimeout(PRIVATE)
                        if fault == 'size': return 200, 'x' * 1001
                        return fault, PRIVATE
                    if span == ('2026-10-01', '2026-10-03'):
                        return 200, [row(1), row('002', '2026-10-03')]
                    if span == ('2026-10-03', '2026-10-05'):
                        return 200, [row(2, '2026-10-03'), row(3, '2026-10-04')]
                    self.fail(f'Unexpected split: {span}')
                client, adapter = self.client(respond)
                with patch.object(dobrodel, 'MAX_BYTES', 1000):
                    result = client.fetch_all(original)
                self.assertEqual({int(record['cardId']) for record in result}, {1, 2, 3})
                self.assertEqual(len(result), 3)
                self.assertEqual(len(adapter.calls), 3)
                self.assertEqual(original, filters())  # The caller's filter state is immutable.

    def test_streamed_body_timeout_is_wrapped_by_requests_and_recovers_by_splitting(self):
        from urllib3.exceptions import ReadTimeoutError

        class TimedOutRaw:
            def stream(self, chunk_size, decode_content=True):
                yield json.dumps([row(999)]).encode()
                raise ReadTimeoutError(None, '/report/operative', PRIVATE)

            def close(self):
                pass

            def release_conn(self):
                pass

        def interrupted_response(url):
            response = requests.Response()
            response.status_code = 200
            response.url = url
            response.encoding = 'utf-8'
            response.raw = TimedOutRaw()
            return response

        # Exercise Requests itself: iter_content converts this raw read timeout
        # to ConnectionError, so catching requests.Timeout alone is insufficient.
        with interrupted_response(dobrodel.BASE + '/report/operative') as response:
            with self.assertRaises(requests.exceptions.ConnectionError) as wrapped:
                list(response.iter_content(65536))
        self.assertIsInstance(wrapped.exception.args[0], ReadTimeoutError)

        class StreamingAdapter(LocalAdapter):
            def send(self, request, **options):
                if not self.calls:
                    self.calls.append((request, options))
                    response = interrupted_response(request.url)
                    response.request = request
                    return response
                return super().send(request, **options)

        recovered = [row(101, '2026-10-01'), row(102, '2026-10-04')]
        def respond(request, attempt):
            params = query(request)
            return 200, [recovered[0] if params['filters.createdAfter'] == '2026-10-01' else recovered[1]]

        client = dobrodel.DobrodelClient(**ACCOUNT)
        client.session.trust_env = False
        adapter = StreamingAdapter(respond)
        client.session.mount('https://', adapter)
        client.session.mount('http://', adapter)
        self.addCleanup(client.session.close)
        events = []
        client.on_progress = events.append
        result = client.fetch_all(filters())
        self.assertEqual(result, recovered)  # Discard bytes from the failed parent response.
        self.assertEqual([(query(request)['filters.createdAfter'], query(request)['filters.createdBefore'])
                          for request, _ in adapter.calls], [
                              ('2026-10-01', '2026-10-05'),
                              ('2026-10-01', '2026-10-03'),
                              ('2026-10-03', '2026-10-05')])
        self.assertEqual([event['stage'] for event in events].count('split'), 1)
        self.assertTrue(all(query(request)['page'] == '0' for request, _ in adapter.calls))

    def test_single_day_recovery_restarts_smaller_pages_at_zero_after_midstream_failure(self):
        records = [row(i) for i in range(1, 652)]
        failed = False
        def respond(request, attempt):
            nonlocal failed
            params = query(request)
            page, size = int(params['page']), int(params['size'])
            if size == 500 and page == 1 and not failed:
                failed = True
                raise requests.exceptions.ReadTimeout(PRIVATE)
            return 200, records[page * size:(page + 1) * size]
        client, adapter = self.client(respond)
        with patch.object(dobrodel.time, 'sleep') as sleep:
            result = client.fetch_all(filters(stop='2026-10-02'))
        sizes_pages = [(int(query(request)['size']), int(query(request)['page'])) for request, _ in adapter.calls]
        self.assertEqual(sizes_pages[:2], [(500, 0), (500, 1)])
        self.assertLess(sizes_pages[2][0], 500)
        self.assertEqual(sizes_pages[2][1], 0)
        self.assertEqual([int(record['cardId']) for record in result], list(range(1, 652)))
        sleep.assert_called_once()
        for request, _ in adapter.calls:
            self.assertEqual(query(request)['filters.createdBefore'], '2026-10-02')

    def test_single_day_failure_has_only_one_smaller_retry(self):
        def respond(request, attempt):
            raise requests.exceptions.ReadTimeout(PRIVATE)
        client, adapter = self.client(respond)
        with patch.object(dobrodel.time, 'sleep') as sleep, self.assertRaises(dobrodel.DobrodelError) as error:
            client.fetch_all(filters(stop='2026-10-02'))
        self.assertEqual(error.exception.code, 'timeout')
        self.assertEqual(len(adapter.calls), 2)
        self.assertEqual([query(request)['page'] for request, _ in adapter.calls], ['0', '0'])
        self.assertLess(int(query(adapter.calls[1][0])['size']), int(query(adapter.calls[0][0])['size']))
        self.assertNotIn(PRIVATE, str(error.exception))
        sleep.assert_called_once()

    def test_nonretryable_tls_format_and_denied_access_do_not_split(self):
        for fault, expected, attempts in [('tls', 'tls', 1), ('format', 'format', 1), ('login', 'login', 2)]:
            with self.subTest(fault=fault):
                def respond(request, attempt):
                    if fault == 'tls': raise requests.exceptions.SSLError(PRIVATE)
                    if fault == 'format': return 200, '<html>' + PRIVATE + '</html>'
                    return 403, PRIVATE
                client, adapter = self.client(respond)
                with patch.object(client, 'login') as login, patch.object(dobrodel.time, 'sleep') as sleep:
                    with self.assertRaises(dobrodel.DobrodelError) as caught:
                        client.fetch_all(filters())
                self.assertEqual(caught.exception.code, expected)
                self.assertEqual(len(adapter.calls), attempts)
                sleep.assert_not_called()
                self.assertEqual({query(request)['filters.createdBefore'] for request, _ in adapter.calls}, {'2026-10-05'})
                self.assertEqual(login.call_count, 1 if fault == 'login' else 0)

    def test_expired_shared_deadline_prevents_children_and_restores_caller_deadline(self):
        clock = {'now': 100.0}
        def respond(request, attempt):
            clock['now'] = 103.0
            raise requests.exceptions.ReadTimeout(PRIVATE)
        client, adapter = self.client(respond)
        client.deadline = 102.0
        with patch.object(dobrodel.time, 'monotonic', side_effect=lambda: clock['now']), \
             patch.object(dobrodel.time, 'sleep') as sleep:
            with self.assertRaises(dobrodel.DobrodelError) as caught:
                client.fetch_all(filters())
        self.assertEqual(caught.exception.code, 'timeout')
        self.assertEqual(len(adapter.calls), 1)
        self.assertEqual(client.deadline, 102.0)
        self.assertTrue(all(timeout <= 2 for timeout in adapter.calls[0][1]['timeout']))
        sleep.assert_not_called()

    def test_report_has_longer_read_timeout_than_login_but_respects_short_caller_budget(self):
        client, adapter = self.client(lambda request, attempt: (200, []))
        with patch.object(dobrodel.time, 'monotonic', return_value=100):
            client.request('GET', '/login')
            client.request('GET', '/report/operative')
            client.deadline = 109
            client.request('GET', '/report/operative')
        self.assertEqual([options['timeout'][1] for _, options in adapter.calls], [60, 240, 9])
        self.assertTrue(all(options['verify'] is not False for _, options in adapter.calls))

    def test_zero_deadline_is_expired_not_unlimited(self):
        client, adapter = self.client(lambda request, attempt: self.fail('Deadline already expired'))
        client.deadline = 0
        with self.assertRaises(dobrodel.DobrodelError) as caught:
            client.fetch_all(filters())
        self.assertEqual(caught.exception.code, 'timeout')
        self.assertEqual(client.deadline, 0)
        self.assertEqual(adapter.calls, [])

    def test_progress_is_fixed_metadata_never_credentials_filters_or_complaint_body(self):
        def respond(request, attempt):
            if attempt == 1: return 504, PRIVATE
            return 200, [row(attempt)]
        client, _ = self.client(respond)
        events = []
        client.on_progress = events.append
        source = {**filters(), 'private-extra-filter': ACCOUNT['password']}
        client.fetch_all(source)
        self.assertIn('split', [event['stage'] for event in events])
        for event in events:
            self.assertEqual(set(event), {'stage', 'from', 'to', 'page', 'count'})
            self.assertIn(event['stage'], {'report', 'split', 'retry'})
            self.assertGreaterEqual(event['page'], 1)
            self.assertGreaterEqual(event['count'], 0)
            self.assertLessEqual(event['from'], event['to'])
        serialized = json.dumps(events)
        for private in (PRIVATE, ACCOUNT['username'], ACCOUNT['password'], 'private-extra-filter'):
            self.assertNotIn(private, serialized)

    def test_probe_is_one_small_request_even_when_a_full_probe_page_is_returned(self):
        client, adapter = self.client(lambda request, attempt: (200, [row(1)]))
        self.assertEqual(client.probe_report(filters(stop='2026-10-02')), 1)
        self.assertEqual(len(adapter.calls), 1)
        self.assertEqual(query(adapter.calls[0][0])['size'], '1')
        self.assertEqual(query(adapter.calls[0][0])['page'], '0')

    def test_admin_probe_reaches_actual_report_api_and_report_failure_cannot_confirm_access(self):
        from routers.users import integration_edds_check
        for fail in (False, True):
            with self.subTest(fail=fail):
                def respond(request, attempt):
                    path = urlsplit(request.url).path
                    if path == '/login': return 200, FORM
                    if path == '/login/admin': return 200, {}
                    if path == '/OperativeReportGenerating': return 200, REPORT
                    if path == '/report/operative': return (504, PRIVATE) if fail else (200, [])
                    self.fail(f'Unexpected route: {path}')
                client, adapter = self.client(respond)
                with patch('services.auth.integrations.credentials', return_value=ACCOUNT), \
                     patch.object(dobrodel, 'DobrodelClient', return_value=client):
                    result = integration_edds_check()
                payload = json.loads(result.body) if hasattr(result, 'body') else result
                self.assertEqual(payload['ok'], not fail)
                if fail:
                    self.assertEqual(result.status_code, 502)
                    self.assertIn('Вход подтверждён', payload['message'])
                query_params = query(adapter.calls[-1][0])
                self.assertEqual(query_params['size'], '1')
                self.assertEqual(query_params['filters.createdAfter'], (date.today() - timedelta(days=1)).isoformat())
                self.assertEqual(query_params['filters.createdBefore'], date.today().isoformat())
                self.assertEqual(len(adapter.calls), 4)
                self.assertLessEqual(adapter.calls[-1][1]['timeout'][1], 120)
                for private in (PRIVATE, ACCOUNT['password'], ACCOUNT['username']):
                    self.assertNotIn(private, json.dumps(payload))


class CollectorRecoveryTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.path = Path(temp.name) / 'water_daily.json'
        for owner, name, value in [(collector, 'WATER_JSON', self.path), (collector, 'KEEP_DAYS', 0)]:
            patcher = patch.object(owner, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        logger = patch.object(collector, 'log')
        logger.start()
        self.addCleanup(logger.stop)
        self.original = {'from': '2026-09-01', 'to': '2026-10-01', 'days': {
            '2026-09-01': {'Old': [8, 0, 0]}, '2026-09-24': {'Old': [9, 0, 0]},
            '2026-10-01': {'Old': [7, 0, 0]}}}
        self.path.write_text(json.dumps(self.original))
        self.windows = [(date(2026, 9, 25), date(2026, 10, 1), 'fresh'),
                        (date(2026, 9, 18), date(2026, 9, 24), 'older')]

    def run_collector(self, responses, windows=None):
        client = MagicMock()
        client.__enter__.return_value = client
        if callable(responses):
            client.fetch_all.side_effect = responses
        else:
            pending = iter(responses)
            def deliver(*args, **kwargs):
                value = next(pending)
                if isinstance(value, Exception):
                    raise value
                return value(*args, **kwargs) if callable(value) else value
            client.fetch_all.side_effect = deliver
        with patch.object(collector, 'get_credentials', return_value=ACCOUNT), \
             patch.object(collector, 'DobrodelClient', return_value=client), \
             patch.object(collector, 'pull_windows', return_value=windows or self.windows):
            collector.main()
        return client

    def test_fresh_complete_window_survives_failure_of_older_window(self):
        def fail_after_complete_window(*args):
            saved = json.loads(self.path.read_text())['days']
            self.assertEqual(saved['2026-10-01'], {'Synthetic district': [1, 0, 0]})
            self.assertEqual(saved['2026-09-25'], {})
            raise dobrodel.DobrodelError('timeout')
        with self.assertRaises(dobrodel.DobrodelError):
            self.run_collector([[row(1)], fail_after_complete_window])
        saved = json.loads(self.path.read_text())['days']
        self.assertEqual(saved['2026-10-01'], {'Synthetic district': [1, 0, 0]})
        self.assertEqual(saved['2026-09-24'], {'Old': [9, 0, 0]})
        self.assertEqual(saved['2026-09-01'], {'Old': [8, 0, 0]})

    def test_failed_first_window_preserves_original_bytes_and_coverage(self):
        before = self.path.read_bytes()
        with self.assertRaises(dobrodel.DobrodelError):
            self.run_collector([dobrodel.DobrodelError('timeout')])
        self.assertEqual(self.path.read_bytes(), before)

    def test_empty_fresh_window_waits_for_nonempty_evidence_before_clearing(self):
        def later_evidence(*args):
            self.assertEqual(json.loads(self.path.read_text()), self.original)
            return [row(1, '2026-09-24')]
        # MagicMock returns callable items rather than calling them: use a single dispatcher.
        index = iter([[], later_evidence])
        def next_response(*args):
            value = next(index)
            return value(*args) if callable(value) else value
        self.run_collector(next_response)
        saved = json.loads(self.path.read_text())['days']
        self.assertEqual(saved['2026-10-01'], {})
        self.assertEqual(saved['2026-09-24'], {'Synthetic district': [1, 0, 0]})
        self.assertEqual(saved['2026-09-01'], {'Old': [8, 0, 0]})

    def test_all_empty_or_unconfirmed_empty_then_error_preserves_previous_summary(self):
        for responses in ([[], []], [[], dobrodel.DobrodelError('timeout')]):
            with self.subTest(responses=len(responses)):
                before = self.path.read_bytes()
                with self.assertRaises((RuntimeError, dobrodel.DobrodelError)):
                    self.run_collector(responses)
                self.assertEqual(self.path.read_bytes(), before)

    def test_initial_windows_are_newest_first_and_cover_every_requested_day_once(self):
        with patch.object(collector, 'DAYS_BACK', 730), patch.object(collector, 'MAX_CATCHUP', 30):
            windows = collector.pull_windows(None)
        self.assertEqual(windows[0][1], date.today())
        self.assertTrue(all(0 <= (end - start).days < 7 for start, end, _ in windows))
        for newer, older in zip(windows, windows[1:]):
            self.assertEqual(older[1] + timedelta(days=1), newer[0])
        covered = [start + timedelta(days=i) for start, end, _ in windows for i in range((end-start).days + 1)]
        expected = {date.today() - timedelta(days=i) for i in range(30)}
        self.assertEqual(set(covered), expected)
        self.assertEqual(len(covered), len(expected))

    def test_resume_after_only_newest_window_was_saved_has_no_duplicate_or_reordered_windows(self):
        today = date.today()
        days = {(today - timedelta(days=i)).isoformat(): {} for i in range(7)}
        existing = {'from': (today - timedelta(days=6)).isoformat(), 'to': today.isoformat(), 'days': days}
        with patch.object(collector, 'DAYS_BACK', 730), patch.object(collector, 'MAX_CATCHUP', 30):
            windows = collector.pull_windows(existing)
        covered = [start + timedelta(days=i) for start, end, _ in windows for i in range((end-start).days + 1)]
        self.assertEqual(len(covered), len(set(covered)), 'Gap recovery and historical backfill query the same days twice')
        for newer, older in zip(windows, windows[1:]):
            self.assertLess(older[1], newer[0], 'All windows should remain newest first after resuming')
        self.assertTrue(all((end-start).days < 7 for start, end, _ in windows))

    def test_progress_file_contains_only_local_stage_and_version_and_is_not_required_for_save(self):
        self.run_collector([[row(1)], [row(2, '2026-09-24')]])
        progress = json.loads(self.path.with_name('progress.json').read_text())
        self.assertEqual(set(progress), {'updated_at', 'message', 'collector_version'})
        self.assertIsInstance(progress['updated_at'], (int, float))
        self.assertTrue(progress['collector_version'])
        for private in (PRIVATE, ACCOUNT['password'], ACCOUNT['username']):
            self.assertNotIn(private, json.dumps(progress))
        with patch.object(Path, 'write_text', side_effect=OSError('test permission denied')):
            collector.progress('Local safe progress')  # Progress failure must not raise.

    def test_runner_ignores_old_progress_and_accepts_only_current_job_progress(self):
        conn = MagicMock()
        conn.execute.return_value.fetchone.return_value = {'running': 1, 'started_at': 1000, 'message': 'initial'}
        manager = MagicMock()
        manager.__enter__.return_value = conn
        with patch.object(runner, 'DATA', self.path.parent), patch.object(runner, 'get_db_connection', return_value=manager), \
             patch.object(runner.time, 'time', return_value=1010):
            for timestamp, expected in [(999, 'initial'), (1000, 'current'), (1005, 'current')]:
                with self.subTest(timestamp=timestamp):
                    self.path.with_name('progress.json').write_text(json.dumps({'updated_at': timestamp, 'message': 'current'}))
                    self.assertEqual(runner.status()['message'], expected)
            conn.execute.return_value.fetchone.return_value = {'running': 0, 'started_at': 1000, 'message': 'finished'}
            self.assertEqual(runner.status()['message'], 'finished')


if __name__ == '__main__':
    unittest.main()
