"""Survey contract regressions using the real router and an isolated SQLite DB.

Run from the repository root with ``python -m pytest -q test/test_ux_survey.py``.
The production app is intentionally not imported: startup owns unrelated network
services and the production database. Every test uses a fresh temporary database.
"""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.testclient import TestClient

from routers import ux_survey as survey_router
from services import ux_survey as survey
from utils import db


API = '/api/ux-survey'
QUESTION_IDS = [f'q{i}' for i in range(1, 7)]


def scores(value=0):
    return dict.fromkeys(QUESTION_IDS, value)


class SurveyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        for attribute, value in [('DB_FILE', Path(self.temp.name) / 'survey.sqlite'),
                                 ('_SCHEMA_READY', False)]:
            patcher = patch.object(db, attribute, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        self.now = datetime(2026, 10, 13, 20, 59, tzinfo=timezone.utc).timestamp()
        clock = patch.object(survey, '_now', lambda: self.now)
        clock.start()
        self.addCleanup(clock.stop)
        with db.get_db_connection() as conn:
            conn.executemany(
                'INSERT INTO users(id,username,password_hash,role,modules) VALUES (?,?,?,?,?)',
                [(1, 'member', 'private-hash', 'Пользователь', '[]'),
                 (2, 'manager', 'private-hash', 'Управление пользователями', '[]'),
                 (3, 'another', 'private-hash', 'Пользователь', '[]'),
                 (4, 'delegated', 'private-hash', 'Пользователь', '[]')],
            )
            conn.execute('UPDATE account_control SET manager_user_id=2, grants_migrated=1')
        self.actors = {
            str(i): {'id': i, 'username': name, 'role': role, 'modules': []}
            for i, name, role in [(1, 'member', 'Пользователь'),
                                  (2, 'manager', 'Управление пользователями'),
                                  (3, 'another', 'Пользователь'),
                                  (4, 'delegated', 'Пользователь')]
        }
        self.actors['external'] = {'id': 1, 'kc_sub': 'external-1', 'username': 'member', 'modules': []}
        self.app = FastAPI()
        self.app.mount('/static', StaticFiles(directory=Path(survey.__file__).resolve().parents[1] / 'static'), name='static')

        @self.app.middleware('http')
        async def test_identity(request, call_next):
            # Keep production authorization dependencies: this only replaces the
            # token-decoding middleware, with the same request.state contract.
            request.state.user = self.actors.get(request.headers.get('x-test-actor'))
            return await call_next(request)

        self.app.include_router(survey_router.router)
        self.member = self.client('1')
        self.other_browser = self.client('1')
        self.manager = self.client('2')

    def client(self, actor=None):
        client = TestClient(self.app, headers={'x-test-actor': actor} if actor else {})
        self.addCleanup(client.close)
        return client

    def state(self, client=None):
        response = (client or self.member).get(API + '/state')
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()

    def analytics(self, client=None):
        response = (client or self.manager).get(API + '/admin')
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()

    def send(self, client=None, ratings=None, comment=''):
        return (client or self.member).post(API + '/submit', json={
            'ratings': scores() if ratings is None else ratings, 'comment': comment,
        })

    def claim(self, client=None):
        response = (client or self.member).post(API + '/invitation/claim')
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()

    def acknowledge(self, token, client=None):
        response = (client or self.member).post(API + '/invitation/shown', json={'token': token})
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()

    def test_fresh_state_has_six_unanswered_questions_and_no_cached_private_data(self):
        response = self.member.get(API + '/state')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers.get('cache-control'), 'no-store')
        data = response.json()
        self.assertEqual(data['survey']['id'], 'service-ux-v1')
        self.assertTrue(data['survey']['title'])
        self.assertEqual([q['id'] for q in data['survey']['questions']], QUESTION_IDS)
        for question in data['survey']['questions']:
            self.assertTrue(question['label'])
            self.assertTrue(question['low'])
            self.assertTrue(question['high'])
        self.assertEqual(data['draft']['ratings'], scores(None))
        self.assertEqual(data['draft']['comment'], '')
        self.assertFalse(data['completed'])
        self.assertFalse(data['has_draft'])
        self.assertIsNone(data['submitted_at'])
        self.assertTrue(data['eligible'])
        self.assertEqual(self.analytics()['completed_count'], 0)
        self.assertIsNone(self.analytics()['completion_rate'])

    def test_all_six_zero_scores_are_valid_answers(self):
        response = self.send()
        self.assertEqual(response.status_code, 200, response.text)
        self.assertTrue(response.json()['completed'])
        self.assertFalse(response.json()['already_submitted'])
        self.assertTrue(self.state()['completed'])
        for question in self.analytics()['questions']:
            self.assertEqual(question['count'], 1)
            self.assertEqual(question['average'], 0)
            self.assertEqual(question['distribution'], [
                {'score': value, 'count': int(value == 0)} for value in range(6)
            ])

    def test_missing_null_boolean_fraction_string_and_out_of_range_are_rejected(self):
        invalid = [None, True, False, -1, 6, 0.0, 1.5, '0', [], {}]
        for value in invalid:
            with self.subTest(value=value):
                response = self.send(ratings={**scores(), 'q3': value})
                self.assertEqual(response.status_code, 422, response.text)
        for missing in QUESTION_IDS:
            with self.subTest(missing=missing):
                response = self.send(ratings={key: 0 for key in QUESTION_IDS if key != missing})
                self.assertEqual(response.status_code, 422, response.text)
        self.assertEqual(self.member.post(API + '/submit', json={}).status_code, 422)
        self.assertEqual(self.send(ratings={**scores(), 'q7': 0}).status_code, 422)
        self.assertFalse(self.state()['completed'])
        self.assertEqual(self.analytics()['completed_count'], 0)

    def test_every_participant_endpoint_requires_authentication(self):
        anonymous = self.client()
        requests = [
            ('get', '/state', {}),
            ('post', '/invitation/claim', {}),
            ('post', '/invitation/shown', {'json': {'token': 'invalid'}}),
            ('post', '/later', {}), ('post', '/open', {}),
            ('put', '/draft', {'json': {'ratings': scores(None), 'comment': '', 'draft_version': 0}}),
            ('post', '/submit', {'json': {'ratings': scores(), 'comment': ''}}),
        ]
        for method, path, kwargs in requests:
            with self.subTest(path=path):
                self.assertEqual(getattr(anonymous, method)(API + path, **kwargs).status_code, 401)

    def test_manager_can_participate_with_no_working_modules(self):
        self.assertTrue(self.claim(self.manager)['show'])
        response = self.send(self.manager, scores(5), 'manager feedback')
        self.assertEqual(response.status_code, 200, response.text)
        self.assertTrue(self.state(self.manager)['completed'])
        self.assertFalse(self.state()['completed'])
        self.assertEqual(self.analytics()['completed_count'], 1)

    def test_survey_page_allows_participants_and_results_page_only_allows_managers(self):
        for client in [self.member, self.manager]:
            response = client.get('/feedback/survey')
            self.assertEqual(response.status_code, 200, response.text)
            self.assertEqual(response.headers.get('cache-control'), 'no-store')
            self.assertIn('id="uxSurveyForm"', response.text)
        response = self.manager.get('/users/feedback')
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.headers.get('cache-control'), 'no-store')
        self.assertIn('id="feedbackAdmin"', response.text)
        self.assertEqual(self.member.get('/users/feedback').status_code, 403)
        self.assertEqual(self.client().get('/users/feedback').status_code, 403)

    def test_inactive_session_is_rejected_and_disabled_manager_loses_statistics_access(self):
        self.actors['1']['is_active'] = False
        self.assertEqual(self.member.get(API + '/state').status_code, 401)
        self.assertEqual(self.send().status_code, 401)
        self.assertEqual(self.member.get('/feedback/survey').status_code, 401)
        with db.get_db_connection() as conn:
            conn.execute('UPDATE users SET is_active=0 WHERE id=2')
        self.assertEqual(self.manager.get(API + '/admin').status_code, 403)
        self.assertEqual(self.manager.get('/users/feedback').status_code, 403)

    def test_server_invitation_cooldown_is_shared_by_two_clients(self):
        first = self.claim()
        self.assertTrue(first['show'])
        self.assertTrue(first['token'])
        self.assertFalse(self.claim(self.other_browser)['show'])
        self.assertFalse(self.state(self.other_browser)['eligible'])
        self.now += 24 * 60 * 60 - 1
        self.assertFalse(self.claim(self.other_browser)['show'])
        self.now += 1
        self.assertTrue(self.state(self.other_browser)['eligible'])
        self.assertTrue(self.claim(self.other_browser)['show'])

    def test_open_and_later_restart_server_cooldown(self):
        self.claim()
        self.now += 100
        for route in ['/open', '/later']:
            with self.subTest(route=route):
                response = self.other_browser.post(API + route)
                self.assertEqual(response.status_code, 200, response.text)
                self.assertFalse(response.json()['eligible'])
                self.now += 86400 - 1
                self.assertFalse(self.claim()['show'])
                self.now += 1
                self.assertTrue(self.state()['eligible'])

    def test_completed_account_never_becomes_eligible_again(self):
        self.send()
        self.now += 365 * 86400
        self.assertFalse(self.state(self.other_browser)['eligible'])
        self.assertFalse(self.claim(self.other_browser)['show'])
        for path in ['/open', '/later']:
            response = self.member.post(API + path)
            self.assertEqual(response.status_code, 200)
            self.assertTrue(response.json()['completed'])
            self.assertFalse(response.json()['eligible'])

    def test_concurrent_claims_have_one_winner_and_only_acknowledged_impressions_count(self):
        clients = [self.client('1') for _ in range(8)]
        with ThreadPoolExecutor(max_workers=8) as pool:
            claims = list(pool.map(self.claim, clients))
        winners = [claim for claim in claims if claim['show']]
        self.assertEqual(len(winners), 1)
        self.assertEqual(self.analytics()['invited_count'], 0)
        token = winners[0]['token']
        self.acknowledge(token)
        self.acknowledge(token, self.other_browser)
        self.assertEqual(self.analytics()['invited_count'], 1)
        self.now += 86400
        next_claim = self.claim()
        self.acknowledge(next_claim['token'])
        self.assertEqual(self.analytics()['invited_count'], 1)

    def test_other_accounts_cannot_acknowledge_an_invitation_token(self):
        claim = self.claim()
        foreign = self.client('3')
        foreign.post(API + '/invitation/shown', json={'token': claim['token']})
        self.assertEqual(self.analytics()['invited_count'], 0)
        self.member.post(API + '/invitation/shown', json={'token': 'not-a-real-claim'})
        self.assertEqual(self.analytics()['invited_count'], 0)
        self.acknowledge(claim['token'])
        self.assertEqual(self.analytics()['invited_count'], 1)

    def test_draft_is_persisted_with_zero_and_null_and_excluded_from_analytics(self):
        start = self.state()
        ratings = {**scores(None), 'q1': 0, 'q6': 5}
        response = self.member.put(API + '/draft', json={
            'ratings': ratings, 'comment': 'unfinished', 'draft_version': start['draft_version'],
        })
        self.assertEqual(response.status_code, 200, response.text)
        saved = self.state(self.other_browser)
        self.assertEqual(saved['draft'], {'ratings': ratings, 'comment': 'unfinished'})
        self.assertTrue(saved['has_draft'])
        self.assertGreater(saved['draft_version'], start['draft_version'])
        report = self.analytics()
        self.assertEqual(report['completed_count'], 0)
        self.assertEqual(report['comments'], [])
        self.assertEqual(report['daily'], [])
        for question in report['questions']:
            self.assertEqual(question['count'], 0)
            self.assertIsNone(question['average'])
            self.assertEqual(sum(item['count'] for item in question['distribution']), 0)

    def test_stale_draft_cannot_overwrite_another_browser(self):
        original = self.state()['draft_version']
        first = {'ratings': {**scores(None), 'q1': 0}, 'comment': 'new', 'draft_version': original}
        response = self.member.put(API + '/draft', json=first)
        self.assertEqual(response.status_code, 200, response.text)
        stale = self.other_browser.put(API + '/draft', json={**first, 'comment': 'stale'})
        self.assertEqual(stale.status_code, 409, stale.text)
        error = stale.json()
        self.assertEqual(error['code'], 'draft_conflict')
        self.assertEqual(error['state']['draft']['comment'], 'new')
        self.assertEqual(self.state()['draft']['comment'], 'new')

    def test_concurrent_draft_saves_accept_only_one_matching_version(self):
        version = self.state()['draft_version']
        clients = [self.client('1') for _ in range(6)]

        def save(candidate):
            index, client = candidate
            return client.put(API + '/draft', json={
                'ratings': {**scores(None), 'q1': index},
                'comment': f'draft-{index}', 'draft_version': version,
            })

        with ThreadPoolExecutor(max_workers=6) as pool:
            responses = list(pool.map(save, enumerate(clients)))
        winners = [response for response in responses if response.status_code == 200]
        self.assertEqual(len(winners), 1)
        for response in responses:
            if response.status_code != 200:
                self.assertEqual(response.status_code, 409, response.text)
                self.assertEqual(response.json()['code'], 'draft_conflict')
                self.assertEqual(response.json()['state']['draft'], winners[0].json()['draft'])
        self.assertEqual(self.state()['draft'], winners[0].json()['draft'])

    def test_submission_prevents_late_draft_write(self):
        version = self.state()['draft_version']
        self.send(comment='final')
        response = self.other_browser.put(API + '/draft', json={
            'ratings': scores(5), 'comment': 'late autosave', 'draft_version': version,
        })
        self.assertEqual(response.status_code, 409, response.text)
        self.assertEqual(response.json()['code'], 'already_submitted')
        self.assertTrue(response.json()['state']['completed'])
        self.assertEqual([item['text'] for item in self.analytics()['comments']], ['final'])

    def test_duplicate_submission_is_idempotent_and_preserves_original_answers(self):
        first = self.send(comment='original')
        self.assertEqual(first.status_code, 200, first.text)
        self.now += 10
        second = self.send(self.other_browser, scores(5), 'replacement')
        self.assertEqual(second.status_code, 200, second.text)
        self.assertTrue(second.json()['already_submitted'])
        self.assertEqual(second.json()['submitted_at'], first.json()['submitted_at'])
        report = self.analytics()
        self.assertEqual(report['completed_count'], 1)
        self.assertEqual([row['text'] for row in report['comments']], ['original'])
        self.assertTrue(all(q['average'] == 0 for q in report['questions']))

    def test_concurrent_submissions_commit_one_immutable_response(self):
        clients = [self.client('1') for _ in range(8)]
        payloads = [(client, i % 6, f'candidate-{i}') for i, client in enumerate(clients)]

        def submit_candidate(payload):
            client, value, comment = payload
            return value, comment, self.send(client, scores(value), comment)

        with ThreadPoolExecutor(max_workers=8) as pool:
            results = list(pool.map(submit_candidate, payloads))
        for _, _, response in results:
            self.assertEqual(response.status_code, 200, response.text)
        winners = [(value, comment) for value, comment, response in results
                   if not response.json()['already_submitted']]
        self.assertEqual(len(winners), 1)
        value, comment = winners[0]
        report = self.analytics()
        self.assertEqual(report['completed_count'], 1)
        self.assertEqual([row['text'] for row in report['comments']], [comment])
        self.assertTrue(all(q['count'] == 1 and q['average'] == value for q in report['questions']))

    def test_analytics_counts_distributions_unique_invite_conversion_and_moscow_days(self):
        first_claim = self.claim()
        self.acknowledge(first_claim['token'])
        self.send(comment='before Moscow midnight')
        # 21:01 UTC is 00:01 on the following calendar day in Europe/Moscow.
        self.now += 120
        self.send(self.manager, scores(5), 'after Moscow midnight')
        third = self.client('3')
        self.acknowledge(self.claim(third)['token'], third)
        report = self.analytics()
        self.assertEqual(report['completed_count'], 2)
        self.assertEqual(report['invited_count'], 2)
        self.assertEqual(report['completed_invited_count'], 1)
        self.assertEqual(report['completion_rate'], 50)
        self.assertEqual(report['daily'], [
            {'date': '2026-10-13', 'count': 1}, {'date': '2026-10-14', 'count': 1},
        ])
        self.assertEqual([row['text'] for row in report['comments']],
                         ['after Moscow midnight', 'before Moscow midnight'])
        for question in report['questions']:
            self.assertEqual(question['count'], 2)
            self.assertEqual(question['average'], 2.5)
            self.assertEqual(question['distribution'], [
                {'score': value, 'count': int(value in {0, 5})} for value in range(6)
            ])

    def test_analytics_requires_current_manager_grant_and_revocation_takes_effect(self):
        self.assertEqual(self.member.get(API + '/admin').status_code, 403)
        self.assertEqual(self.member.get('/users/feedback').status_code, 403)
        self.assertEqual(self.client().get('/feedback/survey').status_code, 401)
        self.actors['1']['role'] = 'admin'
        self.assertEqual(self.member.get(API + '/admin').status_code, 403)
        self.assertEqual(self.client().get(API + '/admin').status_code, 403)
        with db.get_db_connection() as conn:
            conn.execute('INSERT INTO account_managers(user_id) VALUES (4)')
        delegated = self.client('4')
        self.assertEqual(delegated.get(API + '/admin').status_code, 200)
        with db.get_db_connection() as conn:
            conn.execute('DELETE FROM account_managers WHERE user_id=4')
        self.assertEqual(delegated.get(API + '/admin').status_code, 403)
        self.assertEqual(delegated.get('/users/feedback').status_code, 403)
        self.assertEqual(self.manager.get(API + '/admin').status_code, 200)

    def test_an_existing_admin_client_refresh_sees_new_submissions(self):
        response = self.manager.get(API + '/admin')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers.get('cache-control'), 'no-store')
        self.assertEqual(response.json()['completed_count'], 0)
        self.send()
        self.assertEqual(self.analytics()['completed_count'], 1)
        self.send(self.client('3'), scores(4))
        self.assertEqual(self.analytics()['completed_count'], 2)

    def test_empty_comment_and_two_thousand_characters_are_valid_but_2001_is_rejected(self):
        response = self.send(comment='x' * 2001)
        self.assertEqual(response.status_code, 422, response.text)
        self.assertFalse(self.state()['completed'])
        self.assertEqual(self.send(comment='x' * 2000).status_code, 200)
        self.assertEqual(self.send(self.manager, comment='').status_code, 200)
        self.assertEqual([row['text'] for row in self.analytics()['comments']], ['x' * 2000])

    def test_comment_is_optional_and_untrusted_html_round_trips_as_data(self):
        comment = '<script>alert("survey")</script><img src=x onerror=alert(1)>'
        response = self.send(comment=comment)
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(self.analytics()['comments'][0]['text'], comment)
        response = self.manager.post(API + '/submit', json={'ratings': scores(3)})
        self.assertEqual(response.status_code, 200, response.text)

    def test_draft_comment_limit_and_rating_validation_match_submission(self):
        version = self.state()['draft_version']
        payload = {'ratings': scores(None), 'comment': 'x' * 2001, 'draft_version': version}
        self.assertEqual(self.member.put(API + '/draft', json=payload).status_code, 422)
        payload['comment'] = 'x' * 2000
        self.assertEqual(self.member.put(API + '/draft', json=payload).status_code, 200)
        payload['draft_version'] = self.state()['draft_version']
        for invalid in [True, False, -1, 6, 2.5, '0']:
            with self.subTest(invalid=invalid):
                payload['ratings'] = {**scores(None), 'q1': invalid}
                self.assertEqual(self.member.put(API + '/draft', json=payload).status_code, 422)

    def test_external_identity_has_separate_storage_from_same_numeric_local_id(self):
        external = self.client('external')
        self.assertEqual(self.send(external, scores(5)).status_code, 200)
        self.assertTrue(self.state(external)['completed'])
        self.assertFalse(self.state()['completed'])
        self.assertEqual(self.send().status_code, 200)
        self.assertEqual(self.analytics()['completed_count'], 2)

    def test_idempotent_migration_preserves_response_draft_and_account_data(self):
        self.send(comment='keep completed response')
        version = self.state(self.manager)['draft_version']
        response = self.manager.put(API + '/draft', json={
            'ratings': {**scores(None), 'q2': 0}, 'comment': 'keep draft', 'draft_version': version,
        })
        self.assertEqual(response.status_code, 200, response.text)
        before_state = self.state(self.manager)
        before_report = self.analytics()
        with db.get_db_connection() as conn:
            before_users = [tuple(row) for row in conn.execute('SELECT * FROM users ORDER BY id')]
            db._init_schema(conn)
            db._init_schema(conn)
            self.assertEqual([tuple(row) for row in conn.execute('SELECT * FROM users ORDER BY id')], before_users)
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM app_schema_migrations WHERE name='20261013_ux_survey'"
            ).fetchone()[0], 1)
            self.assertEqual(conn.execute('PRAGMA foreign_key_check').fetchall(), [])
        self.assertEqual(self.state(self.manager), before_state)
        self.assertEqual(self.analytics(), before_report)

    def test_database_rejects_missing_out_of_range_and_fractional_response_ratings(self):
        survey.defer('local:1')
        for key in QUESTION_IDS:
            for invalid in [None, -1, 6, 1.5]:
                with self.subTest(question=key, invalid=invalid):
                    values = {**scores(), key: invalid}
                    with self.assertRaises(sqlite3.IntegrityError):
                        with db.get_db_connection() as conn:
                            conn.execute(
                                'INSERT INTO ux_survey_responses'
                                '(account_key,survey_id,q1,q2,q3,q4,q5,q6,comment,submitted_at) '
                                'VALUES (?,?,?,?,?,?,?,?,?,?)',
                                ('local:1', survey.SURVEY_ID, *[values[q] for q in QUESTION_IDS], '', self.now),
                            )
        self.assertEqual(self.analytics()['completed_count'], 0)

    def test_database_rejects_duplicate_responses_and_updates_to_completed_answers(self):
        self.send(comment='immutable')
        with self.assertRaises(sqlite3.IntegrityError):
            with db.get_db_connection() as conn:
                conn.execute(
                    'INSERT INTO ux_survey_responses '
                    'SELECT * FROM ux_survey_responses WHERE account_key=? AND survey_id=?',
                    ('local:1', survey.SURVEY_ID),
                )
        for column, replacement in [('q1', 5), ('comment', 'changed'), ('submitted_at', self.now + 100)]:
            with self.subTest(column=column):
                with self.assertRaises(sqlite3.IntegrityError):
                    with db.get_db_connection() as conn:
                        conn.execute(
                            f'UPDATE ux_survey_responses SET {column}=? WHERE account_key=? AND survey_id=?',
                            (replacement, 'local:1', survey.SURVEY_ID),
                        )
        self.assertEqual(self.analytics()['completed_count'], 1)
        self.assertEqual(self.analytics()['comments'][0]['text'], 'immutable')
        self.assertEqual(self.analytics()['questions'][0]['average'], 0)

    def test_database_enforces_draft_score_and_comment_constraints(self):
        survey.defer('local:1')
        for key in QUESTION_IDS:
            for invalid in [-1, 6, 1.5]:
                with self.subTest(question=key, invalid=invalid):
                    with self.assertRaises(sqlite3.IntegrityError):
                        with db.get_db_connection() as conn:
                            conn.execute(
                                f'UPDATE ux_survey_state SET draft_{key}=? WHERE account_key=? AND survey_id=?',
                                (invalid, 'local:1', survey.SURVEY_ID),
                            )
        with self.assertRaises(sqlite3.IntegrityError):
            with db.get_db_connection() as conn:
                conn.execute('UPDATE ux_survey_state SET draft_comment=?', ('x' * 2001,))
        with self.assertRaises(sqlite3.IntegrityError):
            with db.get_db_connection() as conn:
                conn.execute(
                    'INSERT INTO ux_survey_responses'
                    '(account_key,survey_id,q1,q2,q3,q4,q5,q6,comment,submitted_at) '
                    'VALUES (?,?,?,?,?,?,?,?,?,?)',
                    ('local:1', survey.SURVEY_ID, *[0] * 6, 'x' * 2001, self.now),
                )


if __name__ == '__main__':
    unittest.main()
