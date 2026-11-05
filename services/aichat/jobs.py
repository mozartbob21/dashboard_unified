"""Durable chat receipts; same SQLite/thread pattern as the SMS job service.

Transport disconnects never cancel an accepted job. We deliberately do not replay
jobs after a process dies: external effects might already have happened.
"""
import asyncio
import hashlib
import json
import logging
from pathlib import Path
import re
import threading
import time
import uuid

from fastapi import HTTPException
from utils.db import get_db_connection
from services.auth.home_preferences import account_key
from . import private_storage as storage

ACTIVE = ('queued', 'running')
HEARTBEAT_SECONDS = 10
STALE_SECONDS = 90
MODEL_SECONDS = 65 * 60  # Includes two bounded model attempts and preparation.
BLOCK_SECONDS = 2 * 60 * 60
_workers = set()
_lock = threading.Lock()


def current_user(user):
    """Recheck local grants at execution time, without loading passwords."""
    if user.get('kc_sub'):
        return dict(user)
    with get_db_connection() as conn:
        row = conn.execute('SELECT id,username,role,modules,is_active FROM users WHERE id=?', (user.get('id'),)).fetchone()
    if not row or not row['is_active']:
        raise HTTPException(403, 'Аккаунт отключён или недоступен.')
    result = dict(row)
    result['modules'] = json.loads(result['modules'])
    return result


def initialize():
    storage.initialize()
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        columns = {row[1] for row in conn.execute('PRAGMA table_info(aichat_jobs)')}
        if 'required_modules' not in columns:
            conn.execute('ALTER TABLE aichat_jobs ADD COLUMN required_modules TEXT')
    migration = Path(__file__).resolve().parents[2] / 'migrations/20261104_shared_chat.sql'
    with get_db_connection() as conn:
        conn.executescript(migration.read_text(encoding='utf-8'))


def _requirements(value):
    modules = set(storage.required_modules(value))
    action = value.get('action') or {}
    if isinstance(action, dict) and action.get('module_id'):
        modules.update(storage.required_modules({'module_id': action['module_id']}))
    return sorted(modules)


def _row_requirements(row):
    if row['required_modules'] is not None:
        return json.loads(row['required_modules'])
    # Old tasks had no provenance field. A shared dialog cannot automatically
    # expose old module results to accounts without appropriate grants.
    if row['module_id']:
        return _requirements({'module_id': row['module_id']})
    result = json.loads(row['result']) if row['result'] else {}
    if result.get('required_modules') is not None:
        return _requirements(result)
    # report_scope alone described the current report, not earlier history used
    # by the model. It cannot prove the scope of a legacy model answer.
    from core.roles import ALL_MODULE_IDS
    return list(ALL_MODULE_IDS)


def _public(row):
    return {'job_id': row['id'], 'dialog_id': row['dialog_id'], 'status': row['status'],
            'status_url': '/aichat/api/jobs/' + row['id'], 'stage': row['stage'],
            'steps': json.loads(row['steps']), 'created_at': row['created_at'], 'updated_at': row['updated_at'],
            'result': json.loads(row['result']) if row['result'] else None}


def _visible(row, user):
    account_key(user)
    if not storage.can_view(user, {'required_modules': _row_requirements(row), 'module_id': row['module_id']}):
        raise HTTPException(403, 'Недостаточно прав для просмотра данных этой задачи.')
    result = _public(row)
    value = result.get('result')
    if value and not storage.can_view(user, {**value, 'required_modules': _requirements(value)}):
        raise HTTPException(403, 'Недостаточно прав для просмотра результата этой задачи.')
    return result


def _receipt(row, user, *, existing=False):
    result = _visible(row, user)
    if existing:
        result['already_running'] = row['status'] in ACTIVE
        result['notice'] = ('Проверка этого блока уже запущена. Новый запуск не создан; показываем её состояние. '
                            'Параметры могут отличаться — они указаны в исходном запросе общего чата.'
                            if row['status'] in ACTIVE else 'Этот запуск уже обработан. Показываем сохранённый результат.')
        result['message'] = result['notice']
    return result


def _finish(ident, status, result, *, stale_check=False):
    """Persist completion and the assistant message in a single transaction."""
    now = time.time()
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        row = conn.execute('SELECT * FROM aichat_jobs WHERE id=?', (ident,)).fetchone()
        if not row or row['status'] not in ACTIVE:
            return False
        if stale_check and row['heartbeat'] >= now - STALE_SECONDS and row['deadline'] >= now:
            return False
        checkpoint = json.loads(row['result']) if row['result'] else None
        if stale_check and checkpoint and checkpoint.get('block_confirmed'):
            result = checkpoint
            result['answer'] += '\n\nПодготовка вывода ИИ была прервана. Результат блока сохранён.'
            status = 'succeeded' if result.get('ok') else 'failed'
        requirements = _requirements(result) if 'required_modules' in result else sorted(set(_row_requirements(row)) | set(_requirements(result)))
        result = {**result, 'dialog_id': row['dialog_id'], 'required_modules': requirements}
        conn.execute('UPDATE aichat_jobs SET status=?,stage=?,result=?,updated_at=?,required_modules=? WHERE id=?',
                     (status, 'Завершено' if status == 'succeeded' else 'Выполнение остановлено',
                      json.dumps(result, ensure_ascii=False), now, json.dumps(requirements), ident))
        dialog = conn.execute('SELECT body FROM aichat_dialogs WHERE id=?', (row['dialog_id'],)).fetchone()
        if dialog:
            body = json.loads(dialog['body'])
            body['messages'].append({'role': 'assistant', 'content': result.get('answer', ''),
                                     'ts': now, 'event_id': ident, 'team': result.get('team'),
                                     'response_kind': result.get('response_kind', 'answer'),
                                     'skill_run': result.get('skill_run'), 'module_id': row['module_id'],
                                     'report_scope': result.get('report_scope'), 'required_modules': requirements})
            body['updated_at'] = now
            conn.execute('UPDATE aichat_dialogs SET body=?,updated_at=? WHERE id=?', (json.dumps(body, ensure_ascii=False), now, row['dialog_id']))
    return True


def recover():
    initialize()
    now = time.time()
    with get_db_connection() as conn:
        rows = conn.execute("SELECT id,deadline,heartbeat FROM aichat_jobs WHERE status IN ('queued','running') AND (heartbeat<? OR deadline<?)",
                            (now - STALE_SECONDS, now)).fetchall()
    for row in rows:
        timeout = row['deadline'] < now
        message = ('Достигнут предел времени задачи. Результат не подтверждён.' if timeout else
                   'Обработка прервана: сервер был перезапущен или исполнитель перестал отвечать. Результат не подтверждён.')
        _finish(row['id'], 'timed_out' if timeout else 'interrupted',
                {'ok': False, 'response_kind': 'error', 'answer': message +
                 ' Автоматический повтор не выполнялся. Проверьте состояние блока перед новым запуском.'}, stale_check=True)


def checkpoint(ident, result):
    with get_db_connection() as conn:
        row = conn.execute('SELECT * FROM aichat_jobs WHERE id=?', (ident,)).fetchone()
        if not row:
            return
        required = _requirements(result) if 'required_modules' in result else sorted(set(_row_requirements(row)) | set(_requirements(result)))
        conn.execute("UPDATE aichat_jobs SET result=?,updated_at=?,required_modules=? WHERE id=? AND status='running'",
                     (json.dumps({**result, 'block_confirmed': True, 'required_modules': required}, ensure_ascii=False),
                      time.time(), json.dumps(required), ident))


def submit(user, key, payload, *, dialog_id='', title='Новый чат', team=None, module_id=None,
           required_modules=None, confirmation_id=None):
    owner = account_key(user)
    if not isinstance(key, str) or not re.fullmatch(r'[A-Za-z0-9_-]{16,100}', key):
        raise HTTPException(400, 'Нужен идентификатор запроса. Обновите страницу чата.')
    if confirmation_id is not None and (not isinstance(confirmation_id, str) or not re.fullmatch(r'[a-f0-9]{32}', confirmation_id)):
        raise HTTPException(400, 'Некорректный идентификатор подтверждения.')
    from core.roles import effective_modules
    required = _requirements({'required_modules': required_modules if required_modules is not None else
                             ([module_id] if module_id else effective_modules(user)), 'module_id': module_id})
    if not storage.can_view(user, {'required_modules': required}):
        raise HTTPException(403, 'Недостаточно прав для выполнения этого запроса.')
    encoded = json.dumps(payload, sort_keys=True, ensure_ascii=False, allow_nan=False)
    fingerprint = hashlib.sha256(encoded.encode()).hexdigest()
    recover()
    now = time.time()
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        def save_receipt(row):
            conn.execute('INSERT OR IGNORE INTO aichat_request_receipts VALUES (?,?,?,?)', (owner, key, fingerprint, row['id']))
            if confirmation_id:
                conn.execute('INSERT OR IGNORE INTO aichat_confirmations VALUES (?,?)', (confirmation_id, row['id']))
        old = conn.execute('''SELECT j.*, r.fingerprint AS receipt_fingerprint FROM aichat_request_receipts r
                              JOIN aichat_jobs j ON j.id=r.job_id WHERE r.owner=? AND r.request_key=?''', (owner, key)).fetchone()
        if old:
            if old['receipt_fingerprint'] != fingerprint:
                raise HTTPException(409, 'Идентификатор уже использован для другого сообщения.')
            return _receipt(old, user, existing=bool(module_id)), False
        old = conn.execute('SELECT * FROM aichat_jobs WHERE owner=? AND request_key=?', (owner, key)).fetchone()
        if old:
            if old['fingerprint'] != fingerprint:
                raise HTTPException(409, 'Идентификатор уже использован для другого сообщения.')
            save_receipt(old)
            return _receipt(old, user, existing=bool(module_id)), False
        if confirmation_id:
            old = conn.execute('SELECT j.* FROM aichat_confirmations c JOIN aichat_jobs j ON j.id=c.job_id WHERE c.proposal_id=?', (confirmation_id,)).fetchone()
            if not old:
                old = conn.execute('SELECT * FROM aichat_jobs WHERE request_key=? ORDER BY created_at LIMIT 1', ('confirm_' + confirmation_id,)).fetchone()
            if old:
                save_receipt(old)
                return _receipt(old, user, existing=True), False
        if module_id:
            running = conn.execute("SELECT * FROM aichat_jobs WHERE module_id=? AND status IN ('queued','running') ORDER BY created_at LIMIT 1", (module_id,)).fetchone()
            if running:
                save_receipt(running)
                return _receipt(running, user, existing=True), False
        if dialog_id:
            if not conn.execute('SELECT 1 FROM aichat_dialogs WHERE id=?', (dialog_id,)).fetchone():
                raise HTTPException(404, 'Диалог не найден.')
            if conn.execute("SELECT 1 FROM aichat_jobs WHERE dialog_id=? AND status IN ('queued','running')", (dialog_id,)).fetchone():
                raise HTTPException(409, 'В этом общем чате уже выполняется запрос. Дождитесь результата или откройте другой чат.')
        own = conn.execute("SELECT count(*) FROM aichat_jobs WHERE owner=? AND status IN ('queued','running')", (owner,)).fetchone()[0]
        total = conn.execute("SELECT count(*) FROM aichat_jobs WHERE status IN ('queued','running')").fetchone()[0]
        with _lock:
            physical = sum(thread.is_alive() for thread in _workers)
        if own >= 3 or total >= 16 or physical >= 16:
            raise HTTPException(429, 'Сейчас выполняется слишком много задач. Дождитесь завершения одной из них.')
        if not dialog_id:
            dialog_id = uuid.uuid4().hex
            body = {'id': dialog_id, 'title': title, 'team': team, 'messages': [], 'created_at': now, 'updated_at': now}
            conn.execute('INSERT INTO aichat_dialogs VALUES (?,?,?,?)', (dialog_id, owner, json.dumps(body, ensure_ascii=False), now))
        ident = uuid.uuid4().hex
        conn.execute('''INSERT INTO aichat_jobs(id,owner,request_key,fingerprint,dialog_id,status,stage,module_id,created_at,updated_at,heartbeat,deadline,required_modules)
                        VALUES (?,?,?,?,?,'queued',?,?,?,?,?,?,?)''',
                     (ident, owner, key, fingerprint, dialog_id, 'Задача принята', module_id, now, now, now,
                      now + (BLOCK_SECONDS if module_id else MODEL_SECONDS), json.dumps(required)))
        row = conn.execute('SELECT * FROM aichat_jobs WHERE id=?', (ident,)).fetchone()
        save_receipt(row)
    return _public(row), True


def _followup(row, user, conn):
    result = _visible(row, user)
    if (result.get('result') or {}).get('action'):
        followup = conn.execute('SELECT job_id AS id FROM aichat_confirmations WHERE proposal_id=?', (row['id'],)).fetchone()
        if not followup:
            followup = conn.execute('SELECT id FROM aichat_jobs WHERE request_key=? ORDER BY created_at LIMIT 1', ('confirm_' + row['id'],)).fetchone()
        if followup:
            result['result']['action']['confirmed_job_id'] = followup['id']
    return result


def get(user, ident):
    account_key(user)
    recover()
    with get_db_connection() as conn:
        row = conn.execute('SELECT * FROM aichat_jobs WHERE id=?', (ident,)).fetchone()
        if not row:
            raise HTTPException(404, 'Задача не найдена.')
        return _followup(row, user, conn)


def list_jobs(user, dialog_id=None):
    account_key(user)
    recover()
    with get_db_connection() as conn:
        rows = conn.execute('SELECT * FROM aichat_jobs' + (' WHERE dialog_id=?' if dialog_id else '') +
                            ' ORDER BY created_at DESC LIMIT 100', (dialog_id,) if dialog_id else ()).fetchall()
        result = []
        for row in rows:
            try:
                result.append(_followup(row, user, conn))
            except HTTPException as exc:
                if exc.status_code != 403:
                    raise
    return result


def alive(ident):
    with get_db_connection() as conn:
        row = conn.execute('SELECT status,deadline FROM aichat_jobs WHERE id=?', (ident,)).fetchone()
    return bool(row and row['status'] in ACTIVE and row['deadline'] > time.time())


def progress(ident, event):
    from .skill_trace import step
    safe = step(event)
    if not safe:
        return
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        row = conn.execute('SELECT steps,status FROM aichat_jobs WHERE id=?', (ident,)).fetchone()
        if not row or row['status'] not in ACTIVE:
            return
        steps = {s['id']: s for s in json.loads(row['steps'])}
        if len(steps) < 40 or safe['id'] in steps:
            steps[safe['id']] = safe
        conn.execute('UPDATE aichat_jobs SET steps=?,stage=?,updated_at=? WHERE id=?',
                     (json.dumps(list(steps.values()), ensure_ascii=False), safe['label'], time.time(), ident))


def start(ticket, work):
    ident = ticket['job_id']
    def run():
        stop = threading.Event()
        def heartbeat():
            while not stop.wait(HEARTBEAT_SECONDS):
                recover()
                with get_db_connection() as conn:
                    conn.execute("UPDATE aichat_jobs SET heartbeat=? WHERE id=? AND status='running'", (time.time(), ident))
        with get_db_connection() as conn:
            changed = conn.execute("UPDATE aichat_jobs SET status='running',heartbeat=?,updated_at=? WHERE id=? AND status='queued'", (time.time(), time.time(), ident)).rowcount
        if not changed:
            with _lock:
                _workers.discard(threading.current_thread())
            return
        threading.Thread(target=heartbeat, daemon=True, name='chat-heartbeat').start()
        try:
            result = work(lambda event: progress(ident, event))
            if hasattr(result, '__await__'):
                result = asyncio.run(result)
            _finish(ident, 'succeeded' if result.get('ok', True) else 'failed', result)
        except Exception as exc:
            logging.getLogger('aichat').warning('Background request failed job=%s type=%s', ident, type(exc).__name__)
            message = str(exc.detail) if isinstance(exc, HTTPException) else 'Не удалось завершить задачу. Обратитесь к администратору с её номером.'
            _finish(ident, 'failed', {'ok': False, 'response_kind': 'error', 'answer': message})
        finally:
            stop.set()
            with _lock:
                _workers.discard(threading.current_thread())
    thread = threading.Thread(target=run, daemon=True, name='aichat-job-' + ident[:8])
    with _lock:
        _workers.add(thread)
    try:
        thread.start()
    except Exception:
        with _lock:
            _workers.discard(thread)
        _finish(ident, 'failed', {'ok': False, 'response_kind': 'error', 'answer': 'Исполнитель не запущен. Повторите запрос позже.'})
        raise
