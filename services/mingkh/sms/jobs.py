"""Persistent, account-scoped SMS jobs; one collector across worker processes."""
import hashlib
import json
from pathlib import Path
import re
import threading
import time
import uuid

from utils.db import get_db_connection
from . import context

MIGRATION = Path(__file__).resolve().parents[3] / 'migrations/20261019_mingkh_sms.sql'
_thread_lock = threading.Lock()
_thread = None
_wake = threading.Event()
MAX_RESULT = 80 * 1024 * 1024
RETENTION = 7 * 86400


class JobConflict(ValueError):
    pass


def initialize():
    with get_db_connection() as conn:
        conn.executescript(MIGRATION.read_text(encoding='utf-8'))


def owner_key(user):
    if not isinstance(user, dict):
        raise ValueError('Не удалось определить аккаунт. Войдите в систему ещё раз.')
    if user.get('kc_sub') is not None:
        subject = user['kc_sub']
        if type(subject) is not str or not 1 <= len(subject) <= 200 or any(ord(c) < 32 for c in subject):
            raise ValueError('Не удалось определить аккаунт. Войдите в систему ещё раз.')
        return 'keycloak:' + subject
    if type(user.get('id')) is int and user['id'] > 0:
        return 'local:' + str(user['id'])
    raise ValueError('Не удалось определить аккаунт. Войдите в систему ещё раз.')


def _ticket(row):
    return {'job_id': row['id'], 'status_url': '/mingkh/sms/api/jobs/' + row['id'], 'status': row['status']}


def submit(owner, action, payload, key):
    if not isinstance(key, str) or not re.fullmatch(r'[A-Za-z0-9_-]{16,100}', key):
        raise ValueError('Не удалось определить запрос. Повторите действие.')
    from .operations import validate
    payload = validate(action, payload)
    initialize()
    body = json.dumps(payload, ensure_ascii=False, sort_keys=True, allow_nan=False)
    digest = hashlib.sha256((action + '\n' + body).encode()).hexdigest()
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        old = conn.execute('SELECT * FROM mingkh_sms_jobs WHERE owner=? AND request_key=?', (owner, key)).fetchone()
        if old:
            if old['fingerprint'] != digest:
                raise JobConflict('Этот запрос уже использован для другого расчёта. Запустите расчёт заново.')
            ticket = _ticket(old)
        else:
            own = conn.execute("SELECT count(*) FROM mingkh_sms_jobs WHERE owner=? AND status IN ('queued','running')", (owner,)).fetchone()[0]
            total = conn.execute("SELECT count(*) FROM mingkh_sms_jobs WHERE status IN ('queued','running')").fetchone()[0]
            if own >= 8 or total >= 24:
                raise JobConflict('Очередь расчётов заполнена. Дождитесь завершения текущих задач.')
            ident = uuid.uuid4().hex
            conn.execute('INSERT INTO mingkh_sms_jobs(id,owner,action,request_key,fingerprint,payload,status,created_at,updated_at) VALUES (?,?,?,?,?,?,?, ?,?)',
                         (ident, owner, action, key, digest, body, 'queued', time.time(), time.time()))
            ticket = _ticket({'id': ident, 'status': 'queued'})
    kick()
    return ticket


def status(owner, ident):
    if not isinstance(ident, str) or not re.fullmatch('[a-f0-9]{32}', ident):
        return None
    initialize()
    with get_db_connection() as conn:
        row = conn.execute('SELECT * FROM mingkh_sms_jobs WHERE id=? AND owner=?', (ident, owner)).fetchone()
    if not row:
        return None
    result = {**_ticket(row), 'created_at': row['created_at'], 'updated_at': row['updated_at']}
    if row['status'] == 'success':
        try:
            result['result'] = context.json_read(context.DATA / 'jobs' / (ident + '.json'), MAX_RESULT)
        except (OSError, ValueError):
            result.update(status='error', error='Сохранённый результат недоступен. Повторите расчёт.')
    elif row['status'] == 'error':
        result['error'] = row['error']
    else:
        kick()
    return result


def kick():
    global _thread
    with _thread_lock:
        _wake.set()
        if _thread and _thread.is_alive():
            return
        _thread = threading.Thread(target=_run_worker, daemon=True, name='mingkh-sms-jobs')
        _thread.start()


def _run_worker():
    global _thread
    try:
        while True:
            _wake.clear()
            _drain()
            with _thread_lock:
                if not _wake.is_set():
                    _thread = None
                    return
    finally:
        with _thread_lock:
            if _thread is threading.current_thread():
                _thread = None


def _error(exc):
    from services.mingkh.client import PortalError
    from .rm_data import RedmineError
    from fastapi import HTTPException
    # Only application-authored validation and portal errors may reach the UI.
    if isinstance(exc, (PortalError, RedmineError)):
        return str(exc)[:500]
    if isinstance(exc, ValueError) and not isinstance(exc, (json.JSONDecodeError, UnicodeError)):
        text = str(exc)
        if re.match(r'^[А-Яа-яЁё]', text) and '\n' not in text:
            return text[:500]
    if isinstance(exc, HTTPException):
        return 'Администратору нужно проверить сохранённые настройки доступа к источнику.'
    return 'Не удалось завершить расчёт. Сохранённые реестры не удалены. Повторите попытку или обратитесь к администратору.'


def _drain():
    try:
        with context.process_lock(context.DATA / '.collector.lock', timeout=0):
            initialize()
            # Holding the OS lock proves that a previously running owner is gone.
            with get_db_connection() as conn:
                conn.execute("UPDATE mingkh_sms_jobs SET status='error',error=?,payload='{}',updated_at=? WHERE status='running'", ('Сервер был перезапущен во время расчёта. Запустите расчёт повторно.', time.time()))
            while True:
                with get_db_connection() as conn:
                    conn.execute('BEGIN IMMEDIATE')
                    row = conn.execute("SELECT * FROM mingkh_sms_jobs WHERE status='queued' ORDER BY created_at,id LIMIT 1").fetchone()
                    if row is None:
                        break
                    conn.execute("UPDATE mingkh_sms_jobs SET status='running',updated_at=? WHERE id=?", (time.time(), row['id']))
                try:
                    from .operations import execute
                    result = execute(row['action'], json.loads(row['payload']))
                    encoded = json.dumps(result, ensure_ascii=False, allow_nan=False).encode()
                    if len(encoded) > MAX_RESULT:
                        raise ValueError('Результат слишком большой. Сократите период.')
                    context.atomic_json(context.DATA / 'jobs' / (row['id'] + '.json'), result)
                    with get_db_connection() as conn:
                        conn.execute("UPDATE mingkh_sms_jobs SET status='success',payload='{}',updated_at=? WHERE id=?", (time.time(), row['id']))
                except Exception as exc:
                    with get_db_connection() as conn:
                        conn.execute("UPDATE mingkh_sms_jobs SET status='error',error=?,payload='{}',updated_at=? WHERE id=?", (_error(exc), time.time(), row['id']))
            _prune()
    except ValueError:
        pass  # Another worker or the administrator's import owns the collector.


def _prune():
    with get_db_connection() as conn:
        old = conn.execute("SELECT id FROM mingkh_sms_jobs WHERE status IN ('success','error') AND updated_at<?", (time.time() - RETENTION,)).fetchall()
        for row in old:
            path = context.DATA / 'jobs' / (row['id'] + '.json')
            path.unlink(missing_ok=True)
        conn.execute("DELETE FROM mingkh_sms_jobs WHERE status IN ('success','error') AND updated_at<?", (time.time() - RETENTION,))
