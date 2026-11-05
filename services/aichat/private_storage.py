"""Shared chat storage; the historical module name is kept for compatibility.

The owner records authorship, not visibility. Block provenance is enforced on
server projections and on the history supplied to the model.
"""
import json
import time
import uuid
from datetime import datetime
from pathlib import Path

from utils.db import get_db_connection
from .teams import normalize_team

MIGRATION = Path(__file__).resolve().parents[2] / 'migrations/20261104_aichat_jobs.sql'
LEGACY_FILE = Path(__file__).resolve().parents[2] / 'data/aichat/dialogs.json'


def required_modules(message):
    from core.roles import ALL_MODULE_IDS, MODULE_ALIASES
    values = []
    for value in (message.get('required_modules'), (message.get('report_scope') or {}).get('modules')):
        if isinstance(value, list):
            values.extend(value)
    scope = message.get('report_scope') or {}
    if scope.get('module') and not scope.get('modules'):
        values.append(scope['module'])
    if message.get('module_id'):
        values.append(message['module_id'])
    # Unknown provenance must never silently become unrestricted access.
    if any(not isinstance(value, str) or value in ('all', '*') or
           MODULE_ALIASES.get(value, value) not in ALL_MODULE_IDS for value in values):
        return sorted(ALL_MODULE_IDS)
    return sorted({MODULE_ALIASES.get(value, value) for value in values})


def can_view(user, message):
    from core.roles import check_module_access
    return bool(user) and all(check_module_access(user, module) for module in required_modules(message))


def _legacy_provenance(dialog):
    """Normalize only in memory; the original JSON/old DB body is preserved."""
    from core.roles import ALL_MODULE_IDS
    from .engine import GREETING
    inherited = set()
    messages = []
    for original in dialog.get('messages', []):
        msg = dict(original)
        if 'required_modules' not in msg:
            own = set(required_modules(msg))
            if msg.get('role') == 'assistant' and msg.get('content') != GREETING:
                # Old follow-ups could reuse any earlier source context. An
                # unlabelled legacy answer cannot be proved source-free.
                own.update(inherited)
                if not own:
                    own.update(ALL_MODULE_IDS)
            msg['required_modules'] = sorted(own)
        inherited.update(required_modules(msg))
        messages.append(msg)
    return {**dialog, 'messages': messages}


def _legacy_import(conn):
    conn.execute('CREATE TABLE IF NOT EXISTS aichat_legacy_imports (name TEXT PRIMARY KEY, imported_at REAL NOT NULL)')
    if conn.execute("SELECT 1 FROM aichat_legacy_imports WHERE name='shared-json-v1'").fetchone() or not LEGACY_FILE.is_file():
        return
    try:
        data = json.loads(LEGACY_FILE.read_text(encoding='utf-8-sig'))
        if not isinstance(data, list):
            return
        for value in data:
            if not isinstance(value, dict) or not isinstance(value.get('id'), str):
                return
            if not isinstance(value.get('messages', []), list) or any(not isinstance(m, dict) or
                    m.get('role') not in ('user', 'assistant') or not isinstance(m.get('content'), str)
                    for m in value.get('messages', [])):
                return
    except (OSError, ValueError):
        return  # Keep the original file; an unreadable archive is never erased.
    conn.execute('BEGIN IMMEDIATE')
    if conn.execute("SELECT 1 FROM aichat_legacy_imports WHERE name='shared-json-v1'").fetchone():
        return
    for value in data:
        try:
            updated = datetime.fromisoformat(str(value.get('updated_at', ''))).timestamp()
        except (ValueError, OverflowError):
            updated = 0
        # Keep every original field. Provenance is a read-time projection.
        conn.execute('INSERT OR IGNORE INTO aichat_dialogs VALUES (?,?,?,?)',
                     (value['id'], 'legacy:shared', json.dumps(value, ensure_ascii=False), updated))
    conn.execute("INSERT INTO aichat_legacy_imports VALUES ('shared-json-v1',?)", (time.time(),))


def initialize():
    with get_db_connection() as conn:
        conn.executescript(MIGRATION.read_text(encoding='utf-8'))
        _legacy_import(conn)


def list_dialogs(owner):
    initialize()
    with get_db_connection() as conn:
        rows = conn.execute('SELECT body FROM aichat_dialogs ORDER BY updated_at DESC').fetchall()
    result = []
    for row in rows:
        d = json.loads(row['body'])
        msgs = d.get('messages', [])
        result.append({**{k: d.get(k) for k in ('id', 'title', 'team', 'updated_at')},
                       'preview': msgs[-1]['content'][:80] if msgs else 'Нет сообщений', 'count': len(msgs)})
    return result


def get_dialog(did, owner):
    initialize()
    with get_db_connection() as conn:
        row = conn.execute('SELECT body FROM aichat_dialogs WHERE id=?', (did,)).fetchone()
    return _legacy_provenance(json.loads(row['body'])) if row else None


def create_dialog(title='Новый чат', team=None, *, owner, ident=None):
    initialize()
    d = {'id': ident or uuid.uuid4().hex, 'title': title or 'Новый чат', 'team': normalize_team(team),
         'messages': [], 'created_at': time.time(), 'updated_at': time.time()}
    with get_db_connection() as conn:
        conn.execute('INSERT INTO aichat_dialogs VALUES (?,?,?,?)', (d['id'], owner, json.dumps(d, ensure_ascii=False), time.time()))
    return d


def _change(did, owner, change):
    initialize()
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        row = conn.execute('SELECT body FROM aichat_dialogs WHERE id=?', (did,)).fetchone()
        if not row:
            return None
        d = _legacy_provenance(json.loads(row['body']))
        change(d)
        d['updated_at'] = time.time()
        conn.execute('UPDATE aichat_dialogs SET body=?,updated_at=? WHERE id=?',
                     (json.dumps(d, ensure_ascii=False), time.time(), did))
        return d


def set_team(did, team, *, owner):
    selected = normalize_team(team)
    return _change(did, owner, lambda d: d.update(team=selected))


def delete_dialog(did, *, owner):
    initialize()
    with get_db_connection() as conn:
        conn.execute('BEGIN IMMEDIATE')
        if conn.execute("SELECT 1 FROM aichat_jobs WHERE dialog_id=? AND status IN ('queued','running')", (did,)).fetchone():
            from fastapi import HTTPException
            raise HTTPException(409, 'Дождитесь завершения задачи перед удалением общего диалога.')
        return bool(conn.execute('DELETE FROM aichat_dialogs WHERE id=?', (did,)).rowcount)


def append_message(did, role, content, *, owner, event_id=None, **metadata):
    def change(d):
        messages = d.setdefault('messages', [])
        if event_id and any(m.get('event_id') == event_id for m in messages):
            return
        msg = {'role': role, 'content': content, 'ts': time.time(), **metadata}
        msg['required_modules'] = required_modules(msg)
        if event_id:
            msg['event_id'] = event_id
        messages.append(msg)
    return _change(did, owner, change)


def visible_message(user, msg):
    if not can_view(user, msg):
        return {**{k: msg[k] for k in ('role', 'ts', 'event_id') if k in msg},
                'content': 'Данные этого сообщения доступны только пользователям с правами на соответствующие блоки.',
                'restricted': True}
    return dict(msg)


def visible_dialog(user, dialog):
    if dialog:
        messages = [visible_message(user, msg) for msg in _legacy_provenance(dialog).get('messages', [])]
        first = next((msg for msg in messages if msg.get('role') == 'user'), {})
        return {**dialog, 'title': 'Общий диалог' if first.get('restricted') else dialog.get('title'),
                'team': dialog.get('team'), 'messages': messages}
    return dialog
