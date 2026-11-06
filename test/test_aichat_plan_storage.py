"""Agent task durability on real temporary SQLite; no model or portal requests."""
import json
import threading
import time
from concurrent.futures import ThreadPoolExecutor

import pytest
from fastapi import HTTPException

from services.aichat import jobs, private_storage as storage
from services.auth.home_preferences import account_key
from utils import db


USER = {'id': 9001, 'username': 'planner-test', 'role': 'user', 'modules': ['mingkh']}
COLLEAGUE = {'id': 9002, 'username': 'planner-colleague', 'role': 'user', 'modules': ['mingkh']}


@pytest.fixture(autouse=True)
def database(tmp_path, monkeypatch):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'tasks.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(storage, 'LEGACY_FILE', tmp_path / 'absent.json')
    jobs.initialize()
    with db.get_db_connection() as conn:
        for user in (USER, COLLEAGUE):
            conn.execute('INSERT INTO users(id,username,password_hash,role,modules,is_active) VALUES (?,?,?,?,?,1)',
                         (user['id'], user['username'], 'synthetic hash', 'user', json.dumps(user['modules'])))
    yield
    with jobs._lock:
        workers = list(jobs._workers)
    for worker in workers:
        worker.join(timeout=3)


def submit(key='agent_plan_request_001'):
    return jobs.submit(USER, key, {'text': 'Сравни два периода'}, required_modules=['mingkh'])[0]


def running(ticket):
    with db.get_db_connection() as conn:
        conn.execute("UPDATE aichat_jobs SET status='running' WHERE id=?", (ticket['job_id'],))


def partial():
    return {'answer': 'Первый период подтверждён. Жалоб 4.', 'required_modules': ['mingkh'],
            'agent_plan': {'version': 1, 'status': 'running', 'steps': [
                {'id': 'one', 'status': 'succeeded', 'result': {'count': 4}},
                {'id': 'two', 'status': 'running'}]}}


@pytest.mark.parametrize('reason,status', [('heartbeat', 'interrupted'), ('deadline', 'timed_out')])
def test_partial_plan_survives_restart_without_success_or_replay(reason, status):
    ticket = submit()
    running(ticket)
    assert jobs.plan_checkpoint(ticket['job_id'], partial())
    with db.get_db_connection() as conn:
        conn.execute('UPDATE aichat_jobs SET ' + reason + '=? WHERE id=?', (time.time() - 360, ticket['job_id']))
    jobs.recover()
    restored = jobs.get(USER, ticket['job_id'])
    assert restored['status'] == status
    result = restored['result']
    assert result['ok'] is False and result['response_kind'] == 'error'
    assert result['agent_plan']['steps'][0]['result'] == {'count': 4}
    assert result['agent_plan']['steps'][1]['status'] == status
    assert 'Первый период подтверждён' in result['answer']
    assert 'Автоматический повтор не выполнялся' in result['answer']
    assert '_agent_plan_state' not in result
    again, fresh = jobs.submit(USER, 'agent_plan_request_001', {'text': 'Сравни два периода'}, required_modules=['mingkh'])
    assert not fresh and again['job_id'] == ticket['job_id']
    assert not jobs.plan_checkpoint(ticket['job_id'], partial())
    assert not jobs._finish(ticket['job_id'], 'succeeded', {'ok': True, 'answer': 'Late success'})
    dialog = storage.get_dialog(ticket['dialog_id'], account_key(USER))
    assert len(dialog['messages']) == 1 and dialog['messages'][0]['agent_plan'] == result['agent_plan']


def test_partial_checkpoint_never_inherits_success_flag_and_keeps_provenance():
    ticket = submit()
    running(ticket)
    value = {**partial(), 'ok': True, 'block_confirmed': True, 'required_modules': []}
    assert jobs.plan_checkpoint(ticket['job_id'], value)
    stored = jobs.get(USER, ticket['job_id'])['result']
    assert stored['ok'] is False and 'block_confirmed' not in stored
    assert stored['required_modules'] == ['mingkh']
    with pytest.raises(HTTPException) as denied:
        jobs.get({**COLLEAGUE, 'modules': []}, ticket['job_id'])
    assert denied.value.status_code == 403


def test_completed_plan_can_be_checkpointed_before_narration():
    ticket = submit()
    running(ticket)
    assert jobs.plan_checkpoint(ticket['job_id'], partial())
    final = {'ok': True, 'answer': 'Оба периода подтверждены.', 'required_modules': ['mingkh'],
             'agent_plan': {'version': 1, 'status': 'succeeded', 'steps': []}}
    jobs.checkpoint(ticket['job_id'], final)
    with db.get_db_connection() as conn:
        conn.execute('UPDATE aichat_jobs SET heartbeat=? WHERE id=?', (time.time() - 360, ticket['job_id']))
    restored = jobs.get(USER, ticket['job_id'])
    assert restored['status'] == 'succeeded' and restored['result']['ok']
    assert 'Оба периода подтверждены' in restored['result']['answer']


def test_worker_exception_after_revocation_preserves_partial_data_behind_rights():
    ticket = submit()
    def work(publish):
        assert jobs.plan_checkpoint(ticket['job_id'], partial())
        with db.get_db_connection() as conn:
            conn.execute("UPDATE users SET modules='[]' WHERE id=?", (USER['id'],))
        raise HTTPException(403, 'Права изменились. Выполнение остановлено.')
    jobs.start(ticket, work)
    deadline = time.monotonic() + 3
    while time.monotonic() < deadline:
        with db.get_db_connection() as conn:
            status = conn.execute('SELECT status FROM aichat_jobs WHERE id=?', (ticket['job_id'],)).fetchone()[0]
        if status not in jobs.ACTIVE:
            break
        time.sleep(.01)
    assert status == 'failed'
    with pytest.raises(HTTPException) as denied:
        jobs.get(jobs.current_user(USER), ticket['job_id'])
    assert denied.value.status_code == 403
    permitted = jobs.get(COLLEAGUE, ticket['job_id'])['result']
    assert permitted['ok'] is False and permitted['agent_plan']['status'] == 'failed'
    assert permitted['agent_plan']['steps'][0]['result'] == {'count': 4}
    assert permitted['agent_plan']['steps'][1]['status'] == 'failed'
    assert 'Права изменились' in permitted['answer']
    assert 'Первый период подтверждён' in permitted['answer']
    history = storage.visible_dialog(jobs.current_user(USER), storage.get_dialog(ticket['dialog_id'], account_key(USER)))
    assert 'agent_plan' not in history['messages'][-1]


def test_total_budget_is_bounded_from_acceptance_and_never_revives():
    ticket = submit()
    assert jobs.set_budget(ticket['job_id'], jobs.PLAN_SECONDS)
    with db.get_db_connection() as conn:
        row = conn.execute('SELECT created_at,deadline FROM aichat_jobs WHERE id=?', (ticket['job_id'],)).fetchone()
    assert row['deadline'] - row['created_at'] == jobs.PLAN_SECONDS
    for value in (0, -1, True, 1.5, '7200', jobs.PLAN_SECONDS + 1):
        with pytest.raises(ValueError):
            jobs.set_budget(ticket['job_id'], value)
    with db.get_db_connection() as conn:
        conn.execute('UPDATE aichat_jobs SET deadline=? WHERE id=?', (time.time() - 1, ticket['job_id']))
    assert not jobs.set_budget(ticket['job_id'], jobs.PLAN_SECONDS)
    assert jobs.get(USER, ticket['job_id'])['status'] == 'timed_out'
    assert not jobs.set_budget(ticket['job_id'], jobs.PLAN_SECONDS)


def test_clarification_is_shared_but_not_visible_without_current_rights():
    ticket = submit()
    pending = {'version': 1, 'tool': 'water', 'missing': ['municipality'], 'params': {'year': 2025}}
    assert jobs._finish(ticket['job_id'], 'succeeded', {'ok': True, 'answer': 'Какой округ?',
        'response_kind': 'clarification', 'agent_pending': pending, 'required_modules': ['mingkh']})
    assert jobs.latest_pending(USER, ticket['dialog_id']) == pending
    assert jobs.latest_pending(COLLEAGUE, ticket['dialog_id']) == pending
    assert jobs.latest_pending({**COLLEAGUE, 'modules': []}, ticket['dialog_id']) is None
    changed = jobs.latest_pending(USER, ticket['dialog_id'])
    changed['params']['year'] = 1900
    assert jobs.latest_pending(USER, ticket['dialog_id'])['params']['year'] == 2025
    redacted = storage.visible_dialog({**COLLEAGUE, 'modules': []},
                                     storage.get_dialog(ticket['dialog_id'], account_key(USER)))
    assert 'agent_pending' not in redacted['messages'][-1]


def test_new_answer_or_expired_clarification_cannot_resume_old_task():
    ticket = submit()
    jobs._finish(ticket['job_id'], 'succeeded', {'ok': True, 'answer': 'Какой округ?',
        'response_kind': 'clarification', 'agent_pending': {'version': 1}, 'required_modules': ['mingkh']})
    def expire(dialog):
        dialog['messages'][-1]['ts'] = time.time() - jobs.PENDING_SECONDS - 1
    storage._change(ticket['dialog_id'], account_key(USER), expire)
    assert jobs.latest_pending(USER, ticket['dialog_id']) is None
    storage.append_message(ticket['dialog_id'], 'assistant', 'Другой ответ', owner=account_key(USER),
                           required_modules=[], response_kind='answer')
    assert jobs.latest_pending(USER, ticket['dialog_id']) is None


def test_new_restricted_answer_cannot_fall_back_to_old_permitted_pending():
    ticket = submit()
    jobs._finish(ticket['job_id'], 'succeeded', {'ok': True, 'answer': 'Какой округ?',
        'response_kind': 'clarification', 'agent_pending': {'version': 1}, 'required_modules': ['mingkh']})
    storage.append_message(ticket['dialog_id'], 'assistant', 'Другой закрытый вопрос', owner=account_key(USER),
                           required_modules=['cameras'], response_kind='clarification', agent_pending={'version': 2})
    assert jobs.latest_pending(USER, ticket['dialog_id']) is None


def test_long_plan_worker_disconnect_and_duplicate_receipts_use_existing_queue():
    ticket = submit()
    entered, release = threading.Event(), threading.Event()
    calls = []
    def work(publish):
        calls.append(True)
        publish({'id': 'first', 'kind': 'tool', 'label': 'Первый аналитический запуск', 'status': 'running'})
        assert jobs.plan_checkpoint(ticket['job_id'], partial())
        entered.set()
        release.wait(2)
        return {'ok': True, 'answer': 'Подтверждённое сравнение', 'required_modules': ['mingkh'],
                'agent_plan': {'version': 1, 'status': 'succeeded', 'steps': []}}
    jobs.start(ticket, work)
    assert entered.wait(1)
    try:
        def repeat(_):
            return jobs.submit(USER, 'agent_plan_request_001', {'text': 'Сравни два периода'}, required_modules=['mingkh'])
        with ThreadPoolExecutor(max_workers=4) as executor:
            receipts = list(executor.map(repeat, range(4)))
        assert all(not fresh and receipt['job_id'] == ticket['job_id'] for receipt, fresh in receipts)
        assert jobs.get(COLLEAGUE, ticket['job_id'])['status'] == 'running'
    finally:
        release.set()
    deadline = time.monotonic() + 3
    while time.monotonic() < deadline and jobs.get(USER, ticket['job_id'])['status'] in jobs.ACTIVE:
        time.sleep(.01)
    assert calls == [True]
    restored = jobs.get(USER, ticket['job_id'])
    assert restored['status'] == 'succeeded' and restored['result']['answer'] == 'Подтверждённое сравнение'
