"""Versioned server-side history of actual checks. Existing run_history is the identity ledger.

Retention: keep until the operator sets an approved archival/deletion schedule.
No automatic eviction. Each run keeps only fixed finite counters, fixed codes and
bounded enumerated inputs. A read returns at most 100 rows and no raw source output.
"""
from __future__ import annotations

import calendar
import hashlib
import re
import json
import threading
import uuid
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from fastapi import HTTPException
from core.roles import ALL_MODULE_IDS, MODULE_ALIASES, check_module_access
from services import check_metrics as projection
from services.mingkh.sms.context import process_lock
from utils import db

MIGRATION = Path(__file__).resolve().parents[1] / 'migrations/20261105_check_results.sql'
MOSCOW = ZoneInfo('Europe/Moscow')
STATES = {'running','success','partial','error','interrupted','skipped'}
CODES = {'running','measured','unconfirmed','failed','interrupted','skipped','legacy','partial'}
LONG_RUNNING = timedelta(hours=4)  # Diagnostic only: an old timestamp cannot revoke a live worker.
_guards = {}
_guard_lock = threading.Lock()


def initialize(conn):
    conn.executescript(MIGRATION.read_text(encoding='utf-8'))


def _lock_path(ident):
    # Database run IDs must never become caller-controlled filesystem paths.
    return db.DB_FILE.parent / '.check-leases' / (str(uuid.UUID(ident)) + '.lock')


def claim(ident):
    guard = process_lock(_lock_path(ident), timeout=0)
    guard.__enter__()
    with _guard_lock:
        _guards[ident] = guard


def release(ident, *, orphan=False):
    with _guard_lock:
        guard = _guards.pop(ident, None)
    if guard is None and not orphan:
        return
    if guard is not None:
        guard.__exit__(None, None, None)
    try:
        _lock_path(ident).unlink(missing_ok=True)
    except (ValueError, OSError):
        pass


def subject(owner):
    if not isinstance(owner, str) or not (re.fullmatch(r'local:[1-9][0-9]{0,18}',owner) or
            (owner.startswith('keycloak:') and 8 < len(owner) <= 208 and all(c.isprintable() for c in owner))):
        raise ValueError('Недопустимый владелец закрытой проверки.')
    return hashlib.sha256(('neurona-checks-v1:'+owner).encode('utf-8')).hexdigest()


def begin(conn, ident, module, params, owner):
    scope = MODULE_ALIASES.get(module, module)
    conn.execute('''INSERT INTO check_results(run_id,version,scope,private_subject,params,before_signature,result_code)
                    VALUES (?,?,?,?,?,?,'running')''',
                 (ident, projection.VERSION, scope, subject(owner) if owner is not None else None, json.dumps(projection.safe_params(params)),
                  json.dumps(projection.signature(module))))


def finish(conn, ident, status, metrics, code, result_params):
    row = conn.execute('''SELECT r.started_at,r.module_id,r.status,c.params,c.before_signature FROM run_history r
                      JOIN check_results c USING(run_id) WHERE run_id=?''',(ident,)).fetchone()
    if not row or row['status'] != 'running':
        return status
    params = json.loads(row['params'])
    for key, value in projection.safe_params(result_params).items():
        if key in params and params[key] != value:
            params['unknown'] = True
        else:
            params[key] = value
    clean = {}
    if status in ('success','partial'):
        if status == 'partial':
            code = 'partial'
        if metrics is None:
            code, clean, derived = projection.captured(row['module_id'], json.loads(row['before_signature']), row['started_at'], params)
            params.update(derived)
        else:
            clean = projection.safe_metrics(metrics)
            code = code if code in ('partial','unconfirmed','failed') else ('measured' if clean else 'unconfirmed')
        if code == 'failed':
            status = 'error'
        elif code == 'partial':
            status = 'partial'
    else:
        code = {'error':'failed','skipped':'skipped','interrupted':'interrupted'}.get(status,'unconfirmed')
    if code not in ('measured','partial'):
        clean = {}
    conn.execute('UPDATE check_results SET params=?,metrics=?,result_code=?,before_signature=NULL WHERE run_id=?',
                 (json.dumps(params),json.dumps(clean),code,ident))
    return status


def recover():
    """An orphaned OS lock is proof that the process died. Never replay work."""
    now = datetime.now(timezone.utc)
    with db.get_db_connection() as conn:
        initialize(conn)
        rows = conn.execute("""SELECT r.run_id,r.started_at FROM run_history r JOIN check_results c USING(run_id)
                              WHERE r.status='running'""").fetchall()
    for row in rows:
        ident = row['run_id']
        try:
            with process_lock(_lock_path(ident), timeout=0):
                pass
        except (ValueError, OSError):
            # Another worker still holds this specific lease. Check duration is
            # informational and never authorizes us to finish someone else's run.
            continue
        with db.get_db_connection() as conn:
            conn.execute('BEGIN IMMEDIATE')
            changed = conn.execute("""UPDATE run_history SET status='interrupted',finished_at=?,error_message=?
                        WHERE run_id=? AND status='running'""",
                         (now.isoformat(),'Выполнение прервано. Результат не подтверждён.',ident)).rowcount
            if changed:
                conn.execute("UPDATE check_results SET result_code='interrupted',metrics='{}',before_signature=NULL WHERE run_id=?", (ident,))
        release(ident, orphan=True)


def _authorize(user, module):
    if not user:
        raise HTTPException(401,'Требуется авторизация.')
    scope = MODULE_ALIASES.get(module,module)
    if scope not in ALL_MODULE_IDS or not check_module_access(user,scope):
        raise HTTPException(403,'Нет доступа к истории этого блока.')
    return scope


def _module_ids(module):
    # Spelling aliases share an identity; distinct collectors under the same
    # permission (e.g. camera inspections and prescriptions) keep separate runs.
    if module in ('mgkh-rm','mgkh_rm'):
        return ['mgkh-rm','mgkh_rm']
    if module in ('water-dashboard','water_dashboard'):
        return ['water-dashboard','water_dashboard']
    return [module]



def _record(row):
    v = row['version'] or 0
    supported = v == projection.VERSION
    code = row['result_code'] if supported and row['result_code'] in CODES else 'legacy'
    # A pre-migration unfinished row has no executor ownership proof.
    status = row['status'] if v or row['status'] != 'running' else 'unconfirmed'
    common = {'run_id':row['run_id'],'module_id':row['module_id'],'scope':row['scope'] or MODULE_ALIASES.get(row['module_id'],row['module_id']),
              'at':row['started_at'],'finished_at':row['finished_at'],'status':status,
              'version':v,'visibility':'private' if row['is_private'] else 'shared','result_code':code,'params':{},'metrics':{}}
    if status == 'running':
        started = datetime.fromisoformat(row['started_at'])
        common['long_running'] = datetime.now(timezone.utc) - started > LONG_RUNNING
    if supported:
        try:
            raw_params, raw_metrics = json.loads(row['params']), json.loads(row['metrics'])
            if not isinstance(raw_params, dict) or not isinstance(raw_metrics, dict):
                raise ValueError('invalid projection')
            common['params'] = projection.safe_params(raw_params)
            common['metrics'] = projection.safe_metrics(raw_metrics)
        except (TypeError, ValueError):
            common.update(params={'unknown':True},metrics={},result_code='unconfirmed')
    return common


def _date_filter(value, end=False):
    if isinstance(value, datetime):
        dt = value
    else:
        try:
            d = date.fromisoformat(value) if isinstance(value,str) else value
            dt = datetime.combine(d,time.min,MOSCOW) + (timedelta(days=1) if end else timedelta(0))
        except (ValueError,TypeError):
            raise HTTPException(400,'Некорректная дата истории.') from None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=MOSCOW)
    return dt.astimezone(timezone.utc).isoformat()


def _select(user, module, *, start=None, end=None, status=None, params=None, limit=20, oldest=False, public_only=False):
    ids = _module_ids(module)
    where = ['r.module_id IN (' + ','.join('?' for _ in ids) + ')']
    from services.auth.home_preferences import account_key
    if public_only:
        where.append('c.private_subject IS NULL')
        args = list(ids)
    else:
        where.append('(c.private_subject IS NULL OR c.private_subject=?)')
        args = [*ids, subject(account_key(user))]

    first = _date_filter(start) if start is not None else None
    last = _date_filter(end, True) if end is not None else None
    if first and last and first >= last:
        raise HTTPException(400, 'Начало периода должно предшествовать окончанию.')
    if first:
        where.append('r.started_at >= ?')
        args.append(first)
    if last:
        where.append('r.started_at < ?')
        args.append(last)
    if status is not None:
        if status not in STATES:
            raise HTTPException(400,'Неизвестный статус проверки.')
        where.append('r.status=?')
        args.append(status)
    if params is not None:
        for key, value in projection.safe_params(params).items():
            where.append("json_extract(CASE WHEN json_valid(c.params) THEN c.params ELSE '{}' END,?)=?")
            args.extend(['$.' + key, value])
    query = '''SELECT r.run_id,r.module_id,r.status,r.started_at,r.finished_at,c.version,c.scope,c.params,c.metrics,c.result_code,c.private_subject IS NOT NULL AS is_private
               FROM run_history r LEFT JOIN check_results c USING(run_id) WHERE ''' + ' AND '.join(where)
    order = 'ASC' if oldest else 'DESC'
    with db.get_db_connection() as conn:
        rows = conn.execute(query + f' ORDER BY r.started_at {order},r.rowid {order} LIMIT ?', [*args, max(1,min(100,int(limit)))]).fetchall()
    return [_record(row) for row in rows]


def read_checks(user,module_id,*,limit=20,start=None,end=None,status=None,params=None,public_only=False):
    _authorize(user,module_id)
    recover()
    return _select(user,module_id,start=start,end=end,status=status,limit=limit,params=params,public_only=public_only)


def compare(user,module_id,*,window='last_two',as_of=None,params=None,public_only=False):
    _authorize(user,module_id)
    if window not in ('last_two','month'):
        raise HTTPException(400,'Неизвестный период сравнения.')
    recover()
    now = as_of or datetime.now(timezone.utc)
    if isinstance(now,date) and not isinstance(now,datetime):
        now = datetime.combine(now,time.max,MOSCOW)
    if now.tzinfo is None:
        now = now.replace(tzinfo=MOSCOW)
    start = None
    if window == 'month':
        local = now.astimezone(MOSCOW)
        previous = local.replace(day=1) - timedelta(days=1)
        start = local.replace(year=previous.year, month=previous.month, day=min(local.day,calendar.monthrange(previous.year,previous.month)[1]))
    records = _select(user,module_id,start=start,end=now+timedelta(microseconds=1),params=params,limit=2,public_only=public_only)
    latest = records[0] if records else None
    baseline = records[1] if len(records)>1 else None
    if window == 'month' and baseline:
        baseline = _select(user,module_id,start=start,end=now+timedelta(microseconds=1),params=params,limit=1,oldest=True,public_only=public_only)[0]
    result = {'module_id':module_id,'window':window,'status':'insufficient','reason':'Для сравнения нужны две завершённые проверки.',
              'start':start.isoformat() if start else None,'end':now.isoformat(),'latest':latest,'baseline':baseline,'deltas':[]}
    if latest and latest['status'] != 'success':
        result.update(status='unavailable',reason='Последняя проверка не подтвердила успешное завершение.')
        return result
    if not baseline:
        return result
    if any(r['result_code'] != 'measured' or r['status'] != 'success' for r in (latest,baseline)):
        result.update(status='unavailable',reason='В одной из проверок нет подтверждённых показателей.')
        return result
    if (latest['visibility'] != baseline['visibility'] or _module_ids(latest['module_id']) != _module_ids(baseline['module_id']) or latest['params'] != baseline['params']
            or latest['params'].get('unknown') or latest['version'] != baseline['version']):
        result.update(status='incompatible',reason='Отличаются проверки, параметры или версия результата.')
        return result
    common = sorted(latest['metrics'].keys() & baseline['metrics'].keys())
    for key in common:
        previous, current = baseline['metrics'][key],latest['metrics'][key]
        delta = current - previous
        label, unit, polarity = projection.metric_info(key)
        assessment = 'unchanged' if not delta else 'neutral'
        if delta and polarity:
            assessment = 'improved' if (delta > 0) == (polarity == 'higher') else 'worsened'
        result['deltas'].append({'key':key,'label':label,'unit':unit,'previous':previous,'current':current,'delta':delta,
                                 'direction':'up' if delta>0 else 'down' if delta<0 else 'same','assessment':assessment})
    if result['deltas']:
        result.update(status='compared',reason='Сопоставлены показатели этих двух проверок.')
    else:
        result.update(status='incompatible',reason='Нет общих сопоставимых показателей.')
    return result
