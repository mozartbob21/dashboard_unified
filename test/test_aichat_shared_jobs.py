"""Shared task races and access checks with temporary SQLite, never live services."""
from concurrent.futures import ThreadPoolExecutor
import asyncio
import json
import threading
import time
from types import SimpleNamespace
from unittest.mock import Mock

from fastapi import HTTPException
import pytest

from services.aichat import jobs, job_api, private_storage as storage
from utils import db

FIRST = {'id': 41, 'username': 'first', 'role': 'user', 'modules': ['edo', 'cds']}
SECOND = {'id': 42, 'username': 'second', 'role': 'user', 'modules': ['edo', 'cds']}
DENIED = {'id': 43, 'username': 'denied', 'role': 'user', 'modules': []}


@pytest.fixture(autouse=True)
def temporary_database(tmp_path, monkeypatch):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'shared.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(storage, 'LEGACY_FILE', tmp_path / 'absent-legacy.json')
    jobs.initialize()
    with db.get_db_connection() as conn:
        for user in (FIRST, SECOND, DENIED):
            conn.execute('INSERT INTO users(id,username,password_hash,role,modules,is_active) VALUES (?,?,?,?,?,1)',
                         (user['id'], user['username'], 'test-only', user['role'], json.dumps(user['modules'])))


def complete(ticket, **extra):
    return jobs._finish(ticket['job_id'], 'succeeded',
                        {'ok': True, 'answer': 'Фактический результат', 'required_modules': [], **extra})


def confirm(proposal, user):
    class EmptyRequest:
        async def json(self):
            return {}
    response = asyncio.run(job_api.confirm(proposal['job_id'], EmptyRequest(), user))
    return json.loads(response.body)


def proposal(user, key):
    ticket, _ = jobs.submit(user, key, {'text': 'Запусти блок'})
    complete(ticket, action={'module_id': 'edo', 'params': {}, 'summary': 'Заполненность данных'})
    return ticket


def test_simultaneous_accounts_share_one_active_module_and_durable_receipts():
    barrier = threading.Barrier(2)
    def request(user):
        barrier.wait()
        return jobs.submit(user, 'shared_request_123456', {'requested_period': user['id']}, module_id='edo')
    with ThreadPoolExecutor(max_workers=2) as pool:
        replies = list(pool.map(request, (FIRST, SECOND)))
    assert sum(fresh for _, fresh in replies) == 1
    assert len({ticket['job_id'] for ticket, _ in replies}) == 1
    attached = next(ticket for ticket, fresh in replies if not fresh)
    assert attached['already_running'] and 'Новый запуск не создан' in attached['notice']
    assert 'owner' not in attached
    complete(attached, module_id='edo')
    # A retry after completion still means the old request, including the one
    # that attached with different parameters. Only a NEW request can start anew.
    for user in (FIRST, SECOND):
        repeated, fresh = jobs.submit(user, 'shared_request_123456', {'requested_period': user['id']}, module_id='edo')
        assert not fresh and repeated['job_id'] == attached['job_id']
        assert not repeated['already_running']
    with db.get_db_connection() as conn:
        assert conn.execute('SELECT count(*) FROM aichat_jobs').fetchone()[0] == 1
        assert conn.execute('SELECT count(*) FROM aichat_request_receipts').fetchone()[0] == 2


def test_same_shared_proposal_confirmed_concurrently_by_accounts_runs_once(monkeypatch):
    parent = proposal(FIRST, 'one_proposal_request')
    start = Mock()
    monkeypatch.setattr(jobs, 'start', start)
    barrier = threading.Barrier(2)
    def request(user):
        barrier.wait()
        return confirm(parent, user)
    with ThreadPoolExecutor(max_workers=2) as pool:
        replies = list(pool.map(request, (FIRST, SECOND)))
    assert start.call_count == 1
    assert replies[0]['job_id'] == replies[1]['job_id']
    complete(replies[0], module_id='edo')
    for user in (FIRST, SECOND):
        assert confirm(parent, user)['job_id'] == replies[0]['job_id']
        assert jobs.get(user, parent['job_id'])['result']['action']['confirmed_job_id'] == replies[0]['job_id']
    assert start.call_count == 1


def test_different_proposals_attach_to_existing_module_and_never_replay(monkeypatch):
    parents = [proposal(FIRST, 'first_proposal_12345'), proposal(SECOND, 'second_proposal_12345')]
    start = Mock()
    monkeypatch.setattr(jobs, 'start', start)
    first = confirm(parents[0], FIRST)
    second = confirm(parents[1], SECOND)
    assert second['already_running'] and second['job_id'] == first['job_id']
    complete(first, module_id='edo')
    assert confirm(parents[1], FIRST)['job_id'] == first['job_id']
    assert confirm(parents[0], SECOND)['job_id'] == first['job_id']
    assert start.call_count == 1
    with db.get_db_connection() as conn:
        assert conn.execute('SELECT count(*) FROM aichat_confirmations').fetchone()[0] == 2


def test_shared_dialog_collision_is_global_and_does_not_create_duplicate():
    first, _ = jobs.submit(FIRST, 'first_dialog_request', {}, required_modules=[])
    with pytest.raises(HTTPException) as error:
        jobs.submit(SECOND, 'second_dialog_request', {}, dialog_id=first['dialog_id'], required_modules=[])
    assert error.value.status_code == 409
    assert 'общем чате' in error.value.detail
    assert len(jobs.list_jobs(SECOND)) == 1
    complete(first)
    second, fresh = jobs.submit(SECOND, 'second_dialog_request', {}, dialog_id=first['dialog_id'], required_modules=[])
    assert fresh and second['dialog_id'] == first['dialog_id']


def test_scope_hides_all_metadata_and_permissions_can_change():
    ticket, _ = jobs.submit(FIRST, 'scoped_job_request', {}, module_id='edo')
    jobs.progress(ticket['job_id'], {'id': 'secret', 'kind': 'tool', 'label': 'Restricted stage', 'status': 'running'})
    complete(ticket, answer='Restricted answer', module_id='edo', evidence={'secret': 'Restricted evidence'})
    with pytest.raises(HTTPException) as error:
        jobs.get(DENIED, ticket['job_id'])
    assert error.value.status_code == 403
    assert 'Restricted' not in str(error.value.detail)
    assert jobs.list_jobs(DENIED) == []
    granted = {**DENIED, 'modules': ['edo']}
    assert jobs.get(granted, ticket['job_id'])['result']['answer'] == 'Restricted answer'
    assert jobs.get(SECOND, ticket['job_id'])['job_id'] == ticket['job_id']


def test_pending_model_scope_is_conservative_then_source_free_result_is_shared():
    ticket, _ = jobs.submit(FIRST, 'model_scope_request', {})
    with pytest.raises(HTTPException):
        jobs.get(DENIED, ticket['job_id'])
    complete(ticket)
    assert jobs.get(DENIED, ticket['job_id'])['result']['answer'] == 'Фактический результат'


def test_manual_or_scheduler_busy_is_normal_notice_not_failure(monkeypatch):
    from services import run_history
    ticket, _ = jobs.submit(FIRST, 'busy_module_request', {}, module_id='edo')
    monkeypatch.setattr(job_api.block_execution, 'execute', lambda *args, **kwargs:
                        {'status': 'error', 'error_code': 'BLOCK_BUSY', 'url': '/edo'})
    model, runner, progress = Mock(), Mock(), Mock()
    result = job_api.run_block(FIRST, 'edo', {}, ticket, progress, runner=runner, model=model)
    assert result['ok'] and result['already_running']
    assert result['response_kind'] == 'clarification'
    assert 'Повторный запуск не выполнялся' in result['answer']
    model.assert_not_called()
    assert progress.call_args.args[0]['status'] == 'done'
    entry = next(item for item in run_history.get_all() if item['run_id'] == result['run_id'])
    assert entry['status'] == 'skipped'


@pytest.mark.parametrize('extra', [{}, {'report_scope': {'modules': ['edo']}}])
def test_old_unscoped_rows_are_not_exposed_when_becoming_shared(extra):
    ticket, _ = jobs.submit(FIRST, 'legacy_scope_request', {})
    complete(ticket)
    with db.get_db_connection() as conn:
        conn.execute('UPDATE aichat_jobs SET required_modules=NULL,result=? WHERE id=?',
                     (json.dumps({'ok': True, 'answer': 'Unknown legacy sources', **extra}), ticket['job_id']))
    with pytest.raises(HTTPException) as error:
        jobs.get(FIRST, ticket['job_id'])
    assert error.value.status_code == 403
    assert jobs.list_jobs(FIRST) == []


def test_previously_confirmed_request_from_old_release_is_not_replayed(monkeypatch):
    parent = proposal(FIRST, 'legacy_proposal_request')
    old, _ = jobs.submit(FIRST, 'confirm_' + parent['job_id'], {'proposal': parent['job_id']}, module_id='edo')
    complete(old, module_id='edo')
    with db.get_db_connection() as conn:
        conn.execute('DELETE FROM aichat_request_receipts')
        conn.execute('DELETE FROM aichat_confirmations')
    start = Mock()
    monkeypatch.setattr(jobs, 'start', start)
    assert confirm(parent, SECOND)['job_id'] == old['job_id']
    start.assert_not_called()


def test_shared_jobs_still_require_authentication():
    with pytest.raises(HTTPException) as error:
        jobs.list_jobs(None)
    assert error.value.status_code == 401


def test_edds_existing_manual_queue_is_reported_busy_atomically(tmp_path, monkeypatch):
    from services.edds import runner
    from services.auth import integrations
    from services.aichat import block_execution
    credentials = lambda service: {'username': 'test', 'password': 'test'}
    monkeypatch.setattr(integrations, 'credentials', credentials)
    monkeypatch.setattr(runner, 'credentials', credentials)
    monkeypatch.setattr(block_execution, 'BASE', tmp_path)
    subprocess = Mock()
    monkeypatch.setattr(runner.subprocess, 'run', subprocess)
    with db.get_db_connection() as conn:
        conn.execute("UPDATE edds_job SET running=1,started_at=?,message='Existing manual collector' WHERE id=1", (time.time(),))
    # The legacy caller contract is unchanged; the typed outcome is opt-in.
    assert runner.run() is False
    result = block_execution.execute({**FIRST, 'modules': ['edds']}, 'edds', {}, run_id='edds_manual_busy_12345')
    assert result['error_code'] == 'BLOCK_BUSY'
    subprocess.assert_not_called()
    with db.get_db_connection() as conn:
        row = conn.execute('SELECT running,message FROM edds_job WHERE id=1').fetchone()
        assert row['running'] == 1 and row['message'] == 'Existing manual collector'


def test_edds_true_collector_failure_is_not_misreported_busy(tmp_path, monkeypatch):
    from services.edds import runner
    from services.auth import integrations
    from services.aichat import block_execution
    credentials = lambda service: {'username': 'test', 'password': 'test'}
    monkeypatch.setattr(integrations, 'credentials', credentials)
    monkeypatch.setattr(runner, 'credentials', credentials)
    monkeypatch.setattr(runner, 'DATA', tmp_path / 'edds')
    monkeypatch.setattr(block_execution, 'BASE', tmp_path)
    subprocess = Mock(return_value=SimpleNamespace(returncode=1, stdout='Synthetic collector failure'))
    monkeypatch.setattr(runner.subprocess, 'run', subprocess)
    result = block_execution.execute({**FIRST, 'modules': ['edds']}, 'edds', {}, run_id='edds_failed_run_12345')
    assert result['error_code'] == 'BLOCK_FAILED'
    subprocess.assert_called_once()
    with db.get_db_connection() as conn:
        assert conn.execute('SELECT running FROM edds_job WHERE id=1').fetchone()[0] == 0


def test_map_existing_store_lock_maps_to_busy_without_collection(tmp_path, monkeypatch):
    from services.mingkh import water_map, water_map_runner
    from services.auth import integrations
    from services.aichat import block_execution
    credentials = lambda service: {'username': 'test', 'password': 'test'}
    monkeypatch.setattr(integrations, 'credentials', credentials)
    monkeypatch.setattr(water_map_runner, 'credentials', credentials)
    monkeypatch.setattr(block_execution, 'BASE', tmp_path)
    refresh = Mock(side_effect=water_map.StoreBusy('Synthetic existing lock'))
    monkeypatch.setattr(water_map, 'run_refresh', refresh)
    with db.get_db_connection() as conn:
        conn.execute('UPDATE account_control SET manager_user_id=? WHERE id=1', (FIRST['id'],))
    result = block_execution.execute({**FIRST, 'modules': ['mingkh']}, 'mingkh-water-map', {}, run_id='map_existing_run_12345')
    assert result['error_code'] == 'BLOCK_BUSY'
    assert 'Synthetic' not in result['answer']
