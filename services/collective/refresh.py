"""Shared durable refresh jobs. Browser disconnects never cancel a source import."""
from contextlib import contextmanager
from pathlib import Path
import os
import logging
import re
import threading
import time
import uuid

from utils import db
from services.auth.integrations import credentials
from .client import CollectivePortal, PortalError, period
from . import store

MIGRATION = Path(__file__).resolve().parents[2] / 'migrations/20261019_collective_refresh.sql'
_thread = None
_mutex = threading.Lock()
_wake = threading.Event()
_logger = logging.getLogger(__name__)


class BusyError(ValueError):
    pass


def initialize():
    with db.get_db_connection() as conn:
        conn.executescript(MIGRATION.read_text(encoding='utf-8'))


@contextmanager
def _collector_lock():
    """OS lock is released on process death, including on the Windows server."""
    path = Path(db.DB_FILE).resolve().parent / '.collective-refresh.lock'
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('a+b') as handle:
        handle.seek(0, os.SEEK_END)
        if not handle.tell():
            handle.write(b'0')
            handle.flush()
        handle.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            raise BusyError('Обновление уже выполняется.') from None
        try:
            yield
        finally:
            handle.seek(0)
            if os.name == 'nt':
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def _ticket(row):
    return {'job_id': row['id'], 'status': row['status'],
            'status_url': '/mingkh/collective/api/jobs/' + row['id']}


def submit(start, end, request_key):
    if not isinstance(request_key, str) or not re.fullmatch(r"[A-Za-z0-9_-]{16,100}", request_key):
        raise ValueError("Не удалось определить запрос. Повторите обновление.")
    a, b = period(start, end)
    start, end = a.isoformat(), b.isoformat()
    initialize()
    with db.get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        old = conn.execute('SELECT j.* FROM collective_refresh_requests r JOIN collective_refresh_jobs j ON j.id=r.job_id WHERE r.request_key=?', (request_key,)).fetchone()
        if old:
            if (old['start'], old['end']) != (start, end):
                raise BusyError('Этот запрос относится к другому периоду. Запустите обновление заново.')
            return _ticket(old)
        row = conn.execute("SELECT * FROM collective_refresh_jobs WHERE status IN ('queued','running')").fetchone()
        if row:
            if (row['start'], row['end']) != (start, end):
                kick()
                raise BusyError('Сейчас обновляется другой период. Дождитесь завершения и повторите запрос.')
            ticket = _ticket(row)
        else:
            ident = uuid.uuid4().hex
            now = time.time()
            conn.execute('INSERT INTO collective_refresh_jobs(id,start,end,status,created_at,updated_at) VALUES (?,?,?,?,?,?)',
                         (ident, start, end, 'queued', now, now))
            ticket = _ticket({'id': ident, 'status': 'queued'})
        conn.execute('INSERT INTO collective_refresh_requests(request_key,job_id) VALUES (?,?)', (request_key, ticket['job_id']))
    kick()
    return ticket


def status(ident):
    if not isinstance(ident, str) or not re.fullmatch('[a-f0-9]{32}', ident):
        return None
    initialize()
    with db.get_db_connection() as conn:
        row = conn.execute('SELECT * FROM collective_refresh_jobs WHERE id=?', (ident,)).fetchone()
    if not row:
        return None
    result = {**_ticket(row), 'start': row['start'], 'end': row['end'],
              'created_at': row['created_at'], 'updated_at': row['updated_at']}
    if row['status'] == 'error':
        result['error'] = row['error']
    if row['status'] in ('queued', 'running'):
        kick()
    return result


def kick():
    global _thread
    with _mutex:
        _wake.set()
        if _thread and _thread.is_alive():
            return
        _thread = threading.Thread(target=_worker, daemon=True, name='collective-refresh')
        _thread.start()


def _worker():
    global _thread
    try:
        while True:
            _wake.clear()
            _drain()
            with _mutex:
                if not _wake.is_set():
                    _thread = None
                    return
    finally:
        with _mutex:
            if _thread is threading.current_thread():
                _thread = None


def _drain():
    try:
        with _collector_lock():
            initialize()
            with db.get_db_connection() as conn:
                conn.execute("UPDATE collective_refresh_jobs SET status='error',error=?,updated_at=? WHERE status='running'",
                             ('Сервер был перезапущен во время обновления. Сохранённые данные доступны; повторите обновление.', time.time()))
            while True:
                with db.get_db_connection() as conn:
                    conn.execute('BEGIN IMMEDIATE')
                    row = conn.execute("SELECT * FROM collective_refresh_jobs WHERE status='queued' LIMIT 1").fetchone()
                    if row is None:
                        break
                    conn.execute("UPDATE collective_refresh_jobs SET status='running',updated_at=? WHERE id=?", (time.time(), row['id']))
                error = ''
                try:
                    access = credentials('mingkh')
                    if not access:
                        raise PortalError('Администратору нужно сохранить доступ к МинЖКХ в разделе «Пароли».')
                    a, b = period(row['start'], row['end'])
                    kurators, rows = CollectivePortal(**access).fetch(a, b)
                    store.save_snapshot(row['start'], row['end'], kurators, rows)
                except PortalError as exc:
                    error = str(exc)[:500]
                    _logger.warning('Collective refresh failed type=%s code=COLLECTIVE_SOURCE', type(exc).__name__)
                except store.ValidationError as exc:
                    # ValidationError messages are fixed, reviewed user-facing text;
                    # neither source rows nor manual reports are embedded in them.
                    error = str(exc)[:500]
                    _logger.warning('Collective refresh failed type=%s code=COLLECTIVE_VALIDATION', type(exc).__name__)
                except Exception as exc:
                    error = 'Не удалось сохранить обновление. Прежние данные и доклады сохранены. Повторите попытку или обратитесь к администратору.'
                    # Never log exception messages/tracebacks: upstream failures can
                    # contain portal content, connection parameters or credentials.
                    _logger.warning('Collective refresh failed type=%s code=COLLECTIVE_UNEXPECTED', type(exc).__name__)
                with db.get_db_connection() as conn:
                    conn.execute('UPDATE collective_refresh_jobs SET status=?,error=?,updated_at=? WHERE id=?',
                                 ('error' if error else 'success', error, time.time(), row['id']))
            with db.get_db_connection() as conn:
                conn.execute("DELETE FROM collective_refresh_jobs WHERE status IN ('success','error') AND updated_at<?", (time.time() - 30 * 86400,))
    except BusyError:
        pass
