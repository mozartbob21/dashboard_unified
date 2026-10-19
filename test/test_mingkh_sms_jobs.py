"""Queue/validation tests use the real SQLite helper with isolated temporary storage."""
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path
import json
import threading
import time

import httpx
import pytest

from utils import db
from services.mingkh.sms import context, jobs, operations, rm_data

REAL_KICK = jobs.kick
MIGRATION = Path(db.__file__).resolve().parents[1] / 'migrations/20261019_mingkh_sms.sql'


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'dashboard.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(context, 'DATA', tmp_path / 'sms')
    monkeypatch.setattr(context, 'today', lambda: date(2026, 10, 19))
    monkeypatch.setattr(jobs, 'MIGRATION', MIGRATION)
    monkeypatch.setattr(jobs, '_thread', None)
    monkeypatch.setattr(jobs, '_wake', threading.Event())
    monkeypatch.setattr(jobs, 'kick', lambda: None)
    monkeypatch.setattr(httpx.HTTPTransport, 'handle_request', lambda *a, **kw: pytest.fail('Live network is forbidden'))
    yield
    worker = jobs._thread
    if worker is not None:
        worker.join(timeout=3)
        assert not worker.is_alive(), 'Test worker must finish before temporary files disappear'


def payload(**changes):
    return {'date':'2026-10-14', 'from':'2026-10-08', **changes}


def build_payload(**changes):
    return payload(**{'sources':['4'], 'subtopics':['419'], 'omsu':'-1', 'dept':'-1', **changes})


def submit(index=0, owner='local:1', **changes):
    return jobs.submit(owner, 'daily', payload(**changes), f'job-key-{index:016d}')


def rows():
    with db.get_db_connection() as conn:
        return [dict(row) for row in conn.execute('SELECT * FROM mingkh_sms_jobs ORDER BY created_at,id')]


@pytest.mark.parametrize('user', [None, {}, {'id':0}, {'id':False}, {'id':'1'}, {'kc_sub':True}, {'kc_sub':''}, {'kc_sub':'a\nsecret'}])
def test_owner_key_rejects_unknown_or_invalid_identity(user):
    with pytest.raises(ValueError):
        jobs.owner_key(user)


def test_owner_namespaces_and_keycloak_precedence():
    assert jobs.owner_key({'id':1}) == 'local:1'
    assert jobs.owner_key({'id':1, 'kc_sub':'1'}) == 'keycloak:1'
    assert jobs.owner_key({'kc_sub':'test-subject'}) == 'keycloak:test-subject'


@pytest.mark.parametrize('key', [None, 0, True, [], 'short', '../'+'a'*20, 'x'*101])
def test_invalid_request_key_rejected_before_database_write(key):
    with pytest.raises(ValueError):
        jobs.submit('local:1', 'daily', payload(), key)
    assert not db.DB_FILE.exists()


@pytest.mark.parametrize('bad', [0, False, '0000-00-00', '2026-10-99', '2026-12-01'])
def test_date_zero_and_invalid_never_become_default(bad):
    with pytest.raises(ValueError):
        operations.validate('daily', payload(**{'from':bad}))
    with pytest.raises(ValueError):
        operations.validate('daily', payload(date=bad))


def test_omitted_dates_default_but_inverted_and_long_dates_reject():
    assert operations.validate('daily', {'date':'2026-10-14'}) == {'from':'2026-10-14', 'date':'2026-10-14'}
    for request in (payload(**{'from':'2026-10-15'}), payload(**{'from':'2020-01-01'})):
        with pytest.raises(ValueError):
            operations.validate('daily', request)


@pytest.mark.parametrize('name,value', [('sources',[]),('sources',['../']),('sources',[4]),('subtopics',['48']),
                                       ('subtopics',None),('omsu',0),('omsu',False),('dept',0),('dept','unknown')])
def test_sources_themes_municipality_and_department_validated(name,value):
    request = build_payload()
    request[name] = value
    with pytest.raises(ValueError):
        operations.validate('build', request)


def test_validation_deduplicates_filters_and_ignores_untrusted_owner():
    result = operations.validate('build', build_payload(sources=['4','4','1'], subtopics=['419','419'], user_id=99))
    assert result['sources'] == ['4','1'] and result['subtopics'] == ['419']
    assert 'user_id' not in result


def test_resonance_comparison_too_far_away_is_rejected_before_queue():
    with pytest.raises(ValueError, match='далеко'):
        operations.validate('rez/data', {'from':'2026-10-08', 'to':'2026-10-14', 'prev_from':'2020-01-01',
                                        'prev_to':'2020-01-07', 'sources':['4'], 'subtopics':['419']})


def test_bad_action_and_request_shapes_rejected():
    for action, request in ((None, {}), ([], {}), ('missing', {}), ('daily', [])):
        with pytest.raises(ValueError):
            operations.validate(action, request)
    with pytest.raises(ValueError):
        jobs.submit('local:1', 'missing', {}, 'valid-request-key-1')
    assert not db.DB_FILE.exists()


def test_unknown_omsu_never_starts_expensive_collection(monkeypatch):
    monkeypatch.setattr(operations.svod, 'omsu_list', lambda: [('1','Тестовый округ')])
    monkeypatch.setattr(operations.svod, 'build', lambda *a, **kw: pytest.fail('Unknown OMSU must not collect'))
    with pytest.raises(ValueError, match='муниципалитет'):
        operations.execute('build', build_payload(omsu='2'))


def test_options_fallback_never_leaks_source_exception(monkeypatch):
    monkeypatch.setattr(operations.svod, 'omsu_list', lambda: (_ for _ in ()).throw(RuntimeError('super-secret-path')))
    result = operations.options()
    assert result['omsu'] == [] and result['warning']
    assert 'secret' not in json.dumps(result)
    assert result['week'] == {'from':'2026-10-08', 'to':'2026-10-14'}
    assert all(item['id'] != '-1' for item in result['sources'])
    assert '48' not in [item['id'] for item in result['subtopics']]


def test_concurrent_same_key_creates_one_job_and_migration_is_idempotent():
    with ThreadPoolExecutor(max_workers=12) as executor:
        tickets = list(executor.map(lambda _: submit(), range(24)))
    assert len({ticket['job_id'] for ticket in tickets}) == 1
    assert len(rows()) == 1
    jobs.initialize()
    assert len(rows()) == 1


def test_same_key_different_payload_conflicts_and_owner_isolated():
    original = submit()
    with pytest.raises(jobs.JobConflict):
        submit(date='2026-10-13')
    other = submit(owner='local:2')
    assert original['job_id'] != other['job_id']
    assert jobs.status('local:2', original['job_id']) is None
    assert jobs.status('local:1', original['job_id'])['status'] == 'queued'
    assert jobs.status('local:1', '../result.json') is None
    assert jobs.status('local:1', None) is None


def test_per_owner_and_global_queue_limits():
    for index in range(8):
        submit(index)
    with pytest.raises(jobs.JobConflict, match='Очередь'):
        submit(8)
    for owner in range(2,18):
        submit(owner= f'local:{owner}')
    assert len(rows()) == 24
    with pytest.raises(jobs.JobConflict):
        submit(owner='local:19')


def test_worker_saves_result_only_for_owner_and_clears_payload(monkeypatch):
    ticket = submit()
    monkeypatch.setattr(operations, 'execute', lambda action, req: {'text':'Готовый отчёт', 'count':0})
    jobs._drain()
    assert jobs.status('local:2', ticket['job_id']) is None
    result = jobs.status('local:1', ticket['job_id'])
    assert result['status'] == 'success' and result['result']['count'] == 0
    assert result['result']['text'] == 'Готовый отчёт'
    assert rows()[0]['payload'] == '{}'
    assert submit()['job_id'] == ticket['job_id']  # Network retry after completion is still idempotent.


def test_worker_continues_queue_after_failure_and_returns_safe_rm_reason(monkeypatch):
    first, second = submit(), submit(1)
    calls = []
    def execute(action, req):
        calls.append(req)
        if len(calls) == 1:
            raise rm_data.RedmineError('РМ МинЖКХ не принял ключ доступа.')
        return {'ok':True}
    monkeypatch.setattr(operations, 'execute', execute)
    jobs._drain()
    assert jobs.status('local:1', first['job_id'])['error'] == 'РМ МинЖКХ не принял ключ доступа.'
    assert jobs.status('local:1', second['job_id'])['status'] == 'success'


@pytest.mark.parametrize('failure', [RuntimeError('api-key-secret'), ValueError('source-secret'), json.JSONDecodeError('secret', 'secret', 0)])
def test_worker_unknown_errors_do_not_leak_details(monkeypatch, failure):
    ticket = submit()
    monkeypatch.setattr(operations, 'execute', lambda *a: (_ for _ in ()).throw(failure))
    jobs._drain()
    result = jobs.status('local:1', ticket['job_id'])
    assert result['status'] == 'error' and 'secret' not in result['error']
    assert rows()[0]['payload'] == '{}'


def test_restart_marks_interrupted_job_and_processes_remaining_queue(monkeypatch):
    interrupted, queued = submit(), submit(1)
    with db.get_db_connection() as conn:
        conn.execute("UPDATE mingkh_sms_jobs SET status='running' WHERE id=?", (interrupted['job_id'],))
    monkeypatch.setattr(operations, 'execute', lambda *a: {'ok':True})
    jobs._drain()
    assert 'перезапущен' in jobs.status('local:1', interrupted['job_id'])['error']
    assert jobs.status('local:1', queued['job_id'])['status'] == 'success'
    assert all(row['payload'] == '{}' for row in rows())


def test_locked_live_worker_never_marked_as_interrupted():
    ticket = submit()
    with db.get_db_connection() as conn:
        conn.execute("UPDATE mingkh_sms_jobs SET status='running' WHERE id=?", (ticket['job_id'],))
    with context.process_lock(context.DATA / '.collector.lock'):
        jobs._drain()
    assert rows()[0]['status'] == 'running'


def test_two_workers_never_execute_collectors_at_same_time(monkeypatch):
    for index in range(3):
        submit(index)
    lock = threading.Lock()
    counts = {'active':0, 'max':0, 'calls':0}
    def execute(*args):
        with lock:
            counts['active'] += 1
            counts['max'] = max(counts['max'], counts['active'])
            counts['calls'] += 1
        time.sleep(0.02)
        with lock:
            counts['active'] -= 1
        return {'ok':True}
    monkeypatch.setattr(operations, 'execute', execute)
    with ThreadPoolExecutor(max_workers=2) as executor:
        list(executor.map(lambda _: jobs._drain(), range(2)))
    assert counts['calls'] == 3 and counts['max'] == 1
    assert all(row['status'] == 'success' for row in rows())


def test_enqueue_during_worker_exit_does_not_lose_wakeup(monkeypatch):
    started, finish = threading.Event(), threading.Event()
    calls = []
    def drain():
        calls.append(1)
        if len(calls) == 1:
            started.set()
            assert finish.wait(timeout=2)
    monkeypatch.setattr(jobs, '_drain', drain)
    REAL_KICK()
    worker = jobs._thread
    assert started.wait(timeout=2)
    REAL_KICK()
    finish.set()
    worker.join(timeout=2)
    assert not worker.is_alive() and len(calls) == 2


def test_missing_result_is_error_and_cannot_read_another_path(monkeypatch):
    ticket = submit()
    monkeypatch.setattr(operations, 'execute', lambda *a: {'ok':True})
    jobs._drain()
    (context.DATA / 'jobs' / (ticket['job_id'] + '.json')).unlink()
    result = jobs.status('local:1', ticket['job_id'])
    assert result['status'] == 'error' and 'недоступен' in result['error']


def test_disk_failure_and_oversize_results_not_marked_success(monkeypatch):
    ticket = submit()
    monkeypatch.setattr(operations, 'execute', lambda *a: {'ok':True})
    monkeypatch.setattr(context, 'atomic_json', lambda *a: (_ for _ in ()).throw(OSError('secret-path')))
    jobs._drain()
    assert jobs.status('local:1', ticket['job_id'])['status'] == 'error'
    assert 'secret' not in jobs.status('local:1', ticket['job_id'])['error']
    second = submit(1)
    monkeypatch.setattr(jobs, 'MAX_RESULT', 2)
    jobs._drain()
    assert 'большой' in jobs.status('local:1', second['job_id'])['error']


def test_prune_removes_only_expired_finished_jobs(monkeypatch):
    old, current = submit(), submit(1)
    monkeypatch.setattr(operations, 'execute', lambda *a: {'ok':True})
    jobs._drain()
    queued = submit(2)
    with db.get_db_connection() as conn:
        conn.execute('UPDATE mingkh_sms_jobs SET updated_at=? WHERE id IN (?,?)',
                     (time.time()-jobs.RETENTION-1, old['job_id'], queued['job_id']))
    jobs._prune()
    assert jobs.status('local:1', old['job_id']) is None
    assert not (context.DATA / 'jobs' / (old['job_id']+'.json')).exists()
    assert jobs.status('local:1', current['job_id'])['status'] == 'success'
    assert jobs.status('local:1', queued['job_id'])['status'] == 'queued'
