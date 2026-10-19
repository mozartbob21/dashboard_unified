"""Shared Collective workspace: immutable source snapshots and separate human work.

All callers must enforce the Collective grant; directory changes/import additionally
require the account-management grant. Values of ``actor`` come from that session.
Source refresh never writes report, assignee or theme. Optimistic versions and SQLite
write transactions prevent simultaneous editors from silently overwriting each other.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
import hashlib
import json
import math
from pathlib import Path
import re
import threading

from utils.db import get_db_connection
from .client import HEADERS, period

MIGRATION = Path(__file__).resolve().parents[2] / 'migrations/20261019_collective_workspace.sql'
MAX_REPORT = 20_000
MAX_NAME = 120
MAX_PEOPLE = 200
MAX_HISTORY = 20
MAX_SOURCE_BYTES = 40 * 1024 * 1024
MAX_IMPORT_BYTES = 32 * 1024 * 1024
MAX_ROWS = 300_000
THEMES = frozenset(('ii', 'kr', 'pr'))
_schema_lock = threading.RLock()
# Moscow is UTC+3 for the dates handled by this workspace; no OS tzdata
# dependency, so the Windows server can import the module without extra packages.
_MOSCOW = timezone(timedelta(hours=3))


class ValidationError(ValueError):
    """Safe, user-facing validation error (never embeds source/report content)."""


class NotFoundError(LookupError):
    pass


class ConflictError(ValueError):
    def __init__(self, message, current=None):
        super().__init__(message)
        self.current = current


def now():
    return datetime.now(timezone.utc)


def today():
    return now().astimezone(_MOSCOW).date()


def _stamp():
    return now().isoformat(timespec='microseconds').replace('+00:00', 'Z')


def _json(value):
    return json.dumps(value, ensure_ascii=False, separators=(',', ':'), allow_nan=False)


def initialize():
    """Idempotent lazy migration; the first shared access fixes the cutoff date."""
    with _schema_lock, get_db_connection() as conn:
        conn.execute('PRAGMA busy_timeout=15000')
        has_migrations = conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='app_schema_migrations'").fetchone()
        ready = has_migrations and conn.execute("SELECT 1 FROM app_schema_migrations WHERE name='20261019_collective_workspace'").fetchone()
        if not ready:
            conn.executescript(MIGRATION.read_text(encoding='utf-8'))
        if not conn.execute('SELECT 1 FROM collective_settings WHERE id=1').fetchone():
            conn.execute('INSERT OR IGNORE INTO collective_settings(id,since,created_at) VALUES(1,?,?)',
                         (today().isoformat(), _stamp()))


@contextmanager
def _db(write=False):
    initialize()
    with get_db_connection() as conn:
        conn.execute('PRAGMA busy_timeout=15000')
        conn.execute('BEGIN IMMEDIATE' if write else 'BEGIN')
        yield conn


def _version(value):
    if type(value) is not int or value < 0:
        raise ValidationError('Нужна актуальная версия записи. Обновите данные и повторите действие.')
    return value


def _text(value, limit, label, *, strip=False, multiline=True):
    if not isinstance(value, str):
        raise ValidationError(f'{label}: нужен текст.')
    value = value.replace('\r\n', '\n').replace('\r', '\n')
    if strip:
        value = value.strip()
    if len(value) > limit:
        raise ValidationError(f'{label}: не более {limit} символов.')
    if any((ord(c) < 32 and c not in ('\n', '\t')) or 0xD800 <= ord(c) <= 0xDFFF for c in value):
        raise ValidationError(f'{label}: недопустимый служебный символ.')
    if not multiline and ('\n' in value or '\t' in value):
        raise ValidationError(f'{label}: значение должно быть в одной строке.')
    return value


def _actor(value):
    return _text(value, 240, 'Имя редактора', strip=True, multiline=False)


def _name(value):
    value = _text(value, MAX_NAME, 'Имя ответственного', strip=True, multiline=False)
    value = ' '.join(value.split())
    if not value:
        raise ValidationError('Укажите имя ответственного.')
    return value


def _ident(value):
    if type(value) is not int or value <= 0:
        raise ValidationError('Не удалось определить ответственного.')
    return value


def _key(value):
    if not isinstance(value, str) or re.fullmatch(r'c_[a-f0-9]{64}', value) is None:
        raise ValidationError('Не удалось определить обращение. Обновите данные.')
    return value


def _dates(start, end):
    if not isinstance(start, str) or not isinstance(end, str):
        raise ValidationError('Укажите начальную и конечную даты.')
    try:
        first, last = period(start, end, today=today())
    except (ValueError, TypeError) as exc:
        raise ValidationError('Укажите корректный период: начало не позже окончания, окончание не в будущем.') from exc
    if first.isoformat() != start or last.isoformat() != end:
        raise ValidationError('Даты должны быть в формате ГГГГ-ММ-ДД.')
    return first, last


def _assignee(row):
    return {name: (bool(row[name]) if name == 'archived' else row[name])
            for name in ('id', 'name', 'archived', 'version', 'created_at', 'updated_at')}


def _assignees(conn):
    rows = conn.execute('SELECT * FROM collective_assignees ORDER BY archived,name_key,id').fetchall()
    return [_assignee(row) for row in rows]


def list_assignees():
    with _db() as conn:
        return _assignees(conn)


def create_assignee(name, actor=''):
    name, actor = _name(name), _actor(actor)
    with _db(True) as conn:
        if conn.execute('SELECT 1 FROM collective_assignees WHERE name_key=?', (name.casefold(),)).fetchone():
            raise ValidationError('Такой ответственный уже есть в справочнике. При необходимости верните его из архива.')
        if conn.execute('SELECT count(*) FROM collective_assignees').fetchone()[0] >= MAX_PEOPLE:
            raise ValidationError(f'В справочнике может быть не более {MAX_PEOPLE} ответственных вместе с архивом.')
        stamp = _stamp()
        cur = conn.execute('INSERT INTO collective_assignees(name,name_key,created_at,updated_at,updated_by) VALUES(?,?,?,?,?)',
                           (name, name.casefold(), stamp, stamp, actor))
        return _assignee(conn.execute('SELECT * FROM collective_assignees WHERE id=?', (cur.lastrowid,)).fetchone())


def update_assignee(ident, expected_version, patch, actor=''):
    ident, version, actor = _ident(ident), _version(expected_version), _actor(actor)
    if not isinstance(patch, dict) or not patch or set(patch) - {'name', 'archived', 'active'}:
        raise ValidationError('Укажите имя или состояние ответственного.')
    if 'active' in patch and 'archived' in patch:
        raise ValidationError('Укажите только одно состояние ответственного.')
    cleaned = {}
    if 'name' in patch:
        cleaned['name'] = _name(patch['name'])
    if 'archived' in patch or 'active' in patch:
        raw = patch.get('archived') if 'archived' in patch else patch['active']
        if type(raw) is not bool:
            raise ValidationError('Состояние ответственного должно быть логическим значением.')
        cleaned['archived'] = raw if 'archived' in patch else not raw
    with _db(True) as conn:
        row = conn.execute('SELECT * FROM collective_assignees WHERE id=?', (ident,)).fetchone()
        if not row:
            raise NotFoundError('Ответственный не найден.')
        if row['version'] != version:
            raise ConflictError('Справочник уже изменён другим пользователем. Проверьте актуальные данные.', _assignee(row))
        name, archived = cleaned.get('name', row['name']), cleaned.get('archived', bool(row['archived']))
        if name == row['name'] and archived == bool(row['archived']):
            return _assignee(row)
        if conn.execute('SELECT 1 FROM collective_assignees WHERE name_key=? AND id<>?', (name.casefold(), ident)).fetchone():
            raise ValidationError('Такой ответственный уже есть в справочнике.')
        conn.execute('UPDATE collective_assignees SET name=?,name_key=?,archived=?,version=version+1,updated_at=?,updated_by=? WHERE id=?',
                     (name, name.casefold(), int(archived), _stamp(), actor, ident))
        return _assignee(conn.execute('SELECT * FROM collective_assignees WHERE id=?', (ident,)).fetchone())


def _source_rows(rows):
    if not isinstance(rows, list) or len(rows) > MAX_ROWS:
        raise ValidationError('Источник вернул некорректный или слишком большой список обращений.')
    result = []
    for row in rows:
        if not isinstance(row, (list, tuple)) or len(row) != len(HEADERS):
            raise ValidationError('Структура обращений источника изменилась. Предыдущие данные сохранены.')
        converted = []
        for cell in row:
            if cell is None:
                converted.append('')
                continue
            if type(cell) not in (str, int, float, bool) or (isinstance(cell, float) and not math.isfinite(cell)):
                raise ValidationError('Источник вернул некорректные значения. Предыдущие данные сохранены.')
            text = str(cell)
            if len(text) > 200_000 or '\x00' in text or any(0xD800 <= ord(c) <= 0xDFFF for c in text):
                raise ValidationError('Источник вернул слишком длинное или некорректное поле. Предыдущие данные сохранены.')
            converted.append(text)
        result.append(converted)
    body = _json(result)
    if len(body.encode('utf-8')) > MAX_SOURCE_BYTES:
        raise ValidationError('Выгрузка слишком велика. Уменьшите период.')
    return result, body


def _norm(value):
    return ' '.join(str(value or '').split()).casefold()


def _aliases(row):
    primary, secondary, source = _norm(row[5]), _norm(row[6]), _norm(row[0])
    if len(primary) > 300 or len(secondary) > 300:
        raise ValidationError('Источник вернул некорректный номер обращения.')
    aliases = []
    if primary:
        aliases.append('primary:' + primary)
    if primary and secondary:
        aliases.append('pair:' + _json([primary, secondary]))
    if secondary and (source or not primary):
        aliases.append('secondary:' + _json([source, secondary]))
    if not aliases:
        # Without source IDs only identical stable content can be matched safely.
        # Status, answer and comments are deliberately excluded from this fallback.
        aliases.append('fallback:' + _json([_norm(row[i]) for i in (0, 1, 2, 8, 9)]))
    return aliases


def _resolve(conn, row, stamp, *, legacy=False):
    aliases = _aliases(row)
    extra = ['legacy-secondary:' + _norm(row[6])] if row[6] else []
    matches = set()
    for alias in aliases:
        found = conn.execute('SELECT complaint_key FROM collective_identity_aliases WHERE alias=?', (alias,)).fetchone()
        if found:
            matches.add(found[0])
    for alias in extra:
        found = conn.execute('SELECT complaint_key FROM collective_identity_aliases WHERE alias=?', (alias,)).fetchone()
        if found:
            previous = json.loads(conn.execute('SELECT source_json FROM collective_complaints WHERE key=?', (found[0],)).fetchone()[0])
            known_primary = {r[0][8:] for r in conn.execute("SELECT alias FROM collective_identity_aliases WHERE complaint_key=? AND alias LIKE 'primary:%'", (found[0],))}
            incompatible = ((row[5] and known_primary and _norm(row[5]) not in known_primary)
                            or (row[0] and previous[0] and _norm(row[0]) != _norm(previous[0])))
            if not incompatible:
                matches.add(found[0])
    if legacy and row[6] and not row[5]:
        matches.update(found[0] for found in conn.execute(
            'SELECT complaint_key FROM collective_secondary_ids WHERE secondary_id=?', (_norm(row[6]),)))
    if len(matches) > 1:
        raise ValidationError('Источник вернул противоречащие друг другу номера обращений. Предыдущие данные сохранены.')
    if matches and row[5]:
        known_primary = {r[0][8:] for r in conn.execute(
            "SELECT alias FROM collective_identity_aliases WHERE complaint_key=? AND alias LIKE 'primary:%'", (next(iter(matches)),))}
        if known_primary and _norm(row[5]) not in known_primary:
            raise ValidationError('Источник вернул противоречащие друг другу номера обращений. Предыдущие данные сохранены.')
    key = next(iter(matches)) if matches else 'c_' + hashlib.sha256(aliases[0].encode()).hexdigest()
    conn.execute('INSERT OR IGNORE INTO collective_complaints(key,source_json,first_seen_at,last_seen_at) VALUES(?,?,?,?)',
                 (key, _json(row), stamp, stamp))
    for alias in aliases + (extra if legacy and not row[5] else []):
        conn.execute('INSERT OR IGNORE INTO collective_identity_aliases(alias,complaint_key) VALUES(?,?)', (alias, key))
    if row[6]:
        conn.execute('INSERT OR IGNORE INTO collective_secondary_ids(secondary_id,complaint_key) VALUES(?,?)', (_norm(row[6]), key))
    conn.execute('INSERT OR IGNORE INTO collective_workflows(complaint_key) VALUES(?)', (key,))
    if not legacy:
        conn.execute('UPDATE collective_complaints SET source_json=?,last_seen_at=? WHERE key=?', (_json(row), stamp, key))
    return key


def _received_date(value):
    value = str(value or '').strip()
    match = re.search(r'\b(\d{2})\.(\d{2})\.(\d{4})\b', value)
    try:
        if match:
            return date(int(match[3]), int(match[2]), int(match[1])).isoformat()
        if re.match(r'^\d{4}-\d{2}-\d{2}(?:$|[ T])', value):
            return date.fromisoformat(value[:10]).isoformat()
    except ValueError:
        pass
    return None


def _can_assign(source_row, since, assigned=False):
    status = str(source_row[14] or '').strip().casefold()
    received = _received_date(source_row[1])
    return bool(assigned or not status or 'в работе' in status or 'в процессе подготовки' in status
                or (received is not None and received >= since))


_WORKFLOW_QUERY = '''SELECT w.*,c.source_json,a.name AS assignee_name,a.archived AS assignee_archived
    FROM collective_workflows w JOIN collective_complaints c ON c.key=w.complaint_key
    LEFT JOIN collective_assignees a ON a.id=w.assignee_id'''


def _workflow(row, since):
    result = {name: row[name] for name in ('version', 'assignee_id', 'report', 'theme', 'assigned_at', 'assigned_by',
                                         'report_at', 'report_by', 'updated_at', 'updated_by')}
    result.update(key=row['complaint_key'], assignee_name=row['assignee_name'] or '',
                  assignee_archived=bool(row['assignee_archived']),
                  can_assign=_can_assign(json.loads(row['source_json']), since, row['assignee_id'] is not None))
    return result


def _get_workflow(conn, key):
    row = conn.execute(_WORKFLOW_QUERY + ' WHERE w.complaint_key=?', (key,)).fetchone()
    if not row:
        raise NotFoundError('Обращение не найдено. Сначала загрузите данные источника.')
    since = conn.execute('SELECT since FROM collective_settings WHERE id=1').fetchone()[0]
    return _workflow(row, since)


def get_workflow(key):
    key = _key(key)
    with _db() as conn:
        return _get_workflow(conn, key)


def _history(conn, key, old, actor, stamp):
    conn.execute('INSERT OR IGNORE INTO collective_workflow_history(complaint_key,version,body_json,recorded_at,recorded_by) VALUES(?,?,?,?,?)',
                 (key, old['version'], _json(old), stamp, actor))
    conn.execute('DELETE FROM collective_workflow_history WHERE complaint_key=? AND id NOT IN '
                 '(SELECT id FROM collective_workflow_history WHERE complaint_key=? ORDER BY version DESC,id DESC LIMIT ?)',
                 (key, key, MAX_HISTORY))


def workflow_history(key):
    key = _key(key)
    with _db() as conn:
        _get_workflow(conn, key)
        return [{'workflow': json.loads(row['body_json']), 'recorded_at': row['recorded_at'], 'recorded_by': row['recorded_by']}
                for row in conn.execute('SELECT * FROM collective_workflow_history WHERE complaint_key=? ORDER BY version DESC,id DESC', (key,))]


def update_workflow(key, expected_version, patch, actor=''):
    key, version, actor = _key(key), _version(expected_version), _actor(actor)
    if not isinstance(patch, dict) or not patch or set(patch) - {'assignee_id', 'report', 'theme'}:
        raise ValidationError('Укажите ответственного, доклад или тему обращения.')
    clean = {}
    if 'report' in patch:
        clean['report'] = _text(patch['report'], MAX_REPORT, 'Доклад', strip=True)
    if 'theme' in patch:
        theme = patch['theme']
        if theme is not None and (not isinstance(theme, str) or theme not in THEMES):
            raise ValidationError('Неизвестная тема обращения.')
        clean['theme'] = theme
    if 'assignee_id' in patch:
        clean['assignee_id'] = None if patch['assignee_id'] is None else _ident(patch['assignee_id'])
    with _db(True) as conn:
        old = _get_workflow(conn, key)
        if old['version'] != version:
            raise ConflictError('Обращение уже изменено другим пользователем. Сверьте данные и повторите сохранение.', old)
        changed = {field: value for field, value in clean.items() if old[field] != value}
        if not changed:
            return old
        if 'assignee_id' in changed:
            if not old['can_assign']:
                raise ValidationError('Закрытое обращение, поступившее до начала распределения, нельзя назначить.')
            if changed['assignee_id'] is not None:
                person = conn.execute('SELECT archived FROM collective_assignees WHERE id=?', (changed['assignee_id'],)).fetchone()
                if not person:
                    raise ValidationError('Ответственный не найден в справочнике.')
                if person['archived']:
                    raise ValidationError('Ответственный в архиве. Выберите действующего сотрудника.')
        stamp = _stamp()
        _history(conn, key, old, actor, stamp)
        updates = dict(changed)
        for field, prefix in (('assignee_id', 'assigned'), ('report', 'report'), ('theme', 'theme')):
            if field in changed:
                updates[prefix + '_at'] = stamp
                updates[prefix + '_by'] = actor
        updates.update(version=version + 1, updated_at=stamp, updated_by=actor)
        conn.execute('UPDATE collective_workflows SET ' + ','.join(field + '=?' for field in updates) + ' WHERE complaint_key=?',
                     (*updates.values(), key))
        return _get_workflow(conn, key)


def _payload(conn, snapshot):
    keys = json.loads(snapshot['keys_json'])
    since = conn.execute('SELECT since FROM collective_settings WHERE id=1').fetchone()[0]
    workflows = {}
    unique = list(dict.fromkeys(keys))
    for offset in range(0, len(unique), 400):
        chunk = unique[offset:offset + 400]
        for row in conn.execute(_WORKFLOW_QUERY + ' WHERE w.complaint_key IN (' + ','.join('?' for _ in chunk) + ')', chunk):
            workflows[row['complaint_key']] = _workflow(row, since)
    stamp = snapshot['fetched_at']
    return {'head': list(HEADERS), 'rows': json.loads(snapshot['rows_json']), 'keys': keys, 'workflows': workflows,
            'assignees': _assignees(conn), 'since': since, 'fetched_at': stamp,
            'snapshot': {'version': snapshot['version'], 'fetched_at': stamp},
            'meta': snapshot['meta'], 'start': snapshot['start'], 'end': snapshot['end'],
            'updated': datetime.fromisoformat(stamp.replace('Z', '+00:00')).astimezone(_MOSCOW).strftime('%d.%m.%Y %H:%M')}


def save_snapshot(start, end, kurators, rows):
    first, last = _dates(start, end)
    rows, body = _source_rows(rows)
    if not isinstance(kurators, (list, tuple)) or len(kurators) > 10_000:
        raise ValidationError('Источник вернул некорректный список ведомств.')
    # Curator metadata is used only for the count and is bounded separately.
    try:
        curator_body = _json(kurators)
    except (TypeError, ValueError) as exc:
        raise ValidationError('Источник вернул некорректный список ведомств.') from exc
    if len(curator_body.encode('utf-8')) > 1024 * 1024:
        raise ValidationError('Источник вернул слишком большой список ведомств.')
    meta = f'МинЖКХ МО и объединения ({len(kurators)}) · {first:%d.%m.%Y}–{last:%d.%m.%Y}'
    stamp = _stamp()
    with _db(True) as conn:
        keys = [_resolve(conn, row, stamp) for row in rows]
        conn.execute('''INSERT INTO collective_snapshots(start,end,rows_json,keys_json,kurators_json,meta,fetched_at)
            VALUES(?,?,?,?,?,?,?) ON CONFLICT(start,end) DO UPDATE SET rows_json=excluded.rows_json,keys_json=excluded.keys_json,
            kurators_json=excluded.kurators_json,meta=excluded.meta,fetched_at=excluded.fetched_at,version=collective_snapshots.version+1''',
            (start, end, body, _json(keys), curator_body, meta, stamp))
        snapshot = conn.execute('SELECT * FROM collective_snapshots WHERE start=? AND end=?', (start, end)).fetchone()
        return _payload(conn, snapshot)


def get_snapshot(start, end):
    _dates(start, end)
    with _db() as conn:
        row = conn.execute('SELECT * FROM collective_snapshots WHERE start=? AND end=?', (start, end)).fetchone()
        return _payload(conn, row) if row else None


def latest_snapshot():
    with _db() as conn:
        row = conn.execute('SELECT * FROM collective_snapshots ORDER BY fetched_at DESC LIMIT 1').fetchone()
        return _payload(conn, row) if row else None


def _legacy_stamp(value):
    if value in (None, ''):
        return None
    value = _text(value, 40, 'Дата старой записи', strip=True, multiline=False)
    try:
        try:
            result = datetime.fromisoformat(value.replace('Z', '+00:00'))
        except ValueError:
            result = datetime.strptime(value, '%d.%m.%Y %H:%M')
        if result.tzinfo is None:
            result = result.replace(tzinfo=_MOSCOW)
        return result.astimezone(timezone.utc).isoformat(timespec='seconds').replace('+00:00', 'Z')
    except ValueError as exc:
        raise ValidationError('В старых данных есть некорректная дата сохранения.') from exc


def validate_legacy(raw):
    """Validate the extracted legacy JSON as data, without evaluating any code."""
    if not isinstance(raw, dict) or set(raw) - {'people', 'archived', 'since', 'items', 'themes'}:
        raise ValidationError('Нужен файл распределения с полями people, archived, since и items.')
    try:
        if len(_json(raw).encode('utf-8')) > MAX_IMPORT_BYTES:
            raise ValidationError('Файл распределения слишком большой.')
    except (TypeError, ValueError, RecursionError) as exc:
        raise ValidationError('Некорректный файл распределения.') from exc
    since = raw.get('since')
    if not isinstance(since, str):
        raise ValidationError('В файле не указана дата начала распределения.')
    try:
        since_date = date.fromisoformat(since)
        if since_date.isoformat() != since or since_date > today():
            raise ValueError()
    except ValueError as exc:
        raise ValidationError('Некорректная дата начала распределения.') from exc
    people, archived, seen = [], [], set()
    for field, target in (('people', people), ('archived', archived)):
        values = raw.get(field, [])
        if not isinstance(values, list):
            raise ValidationError('Ответственные должны быть перечислены списком.')
        for item in values:
            name = _name(item)
            if name.casefold() not in seen:
                seen.add(name.casefold())
                target.append(name)
    if len(seen) > MAX_PEOPLE:
        raise ValidationError(f'В справочнике может быть не более {MAX_PEOPLE} ответственных вместе с архивом.')
    items = raw.get('items', {})
    themes = raw.get('themes', {})
    if not isinstance(items, dict) or len(items) > 50_000 or not isinstance(themes, dict):
        raise ValidationError('Некорректный список обращений в файле распределения.')
    result, canonical_keys = {}, set()
    for key in items.keys() | themes.keys():
        item = items.get(key, {})
        if not isinstance(key, str) or len(key) > 601 or key.count('|') != 1 or not key.strip('|').strip():
            raise ValidationError('В файле есть обращение без корректного номера.')
        primary, secondary = (part.strip() for part in key.split('|', 1))
        if not (primary or secondary) or len(primary) > 300 or len(secondary) > 300 or not isinstance(item, dict):
            raise ValidationError('Некорректная запись обращения в файле распределения.')
        canonical = (_norm(primary), _norm(secondary))
        if canonical in canonical_keys:
            raise ValidationError('В файле повторяется номер обращения. Уберите противоречащие записи.')
        canonical_keys.add(canonical)
        if set(item) - {'resp', 'resp_at', 'resp_by', 'report', 'report_at', 'report_by', 'history', 'theme'}:
            raise ValidationError('В файле есть неизвестные поля обращения.')
        resp = _name(item['resp']) if item.get('resp') else ''
        if resp and resp.casefold() not in seen:
            # Old applications allowed deleting names. Preserve those assignments as archived.
            seen.add(resp.casefold())
            archived.append(resp)
        theme = item.get('theme', themes.get(key))
        if theme is not None and (not isinstance(theme, str) or theme not in THEMES):
            raise ValidationError('В старых данных указана неизвестная тема.')
        report = _text(item.get('report', ''), MAX_REPORT, 'Доклад', strip=True)
        history = item.get('history', [])
        if not isinstance(history, list) or len(history) > MAX_HISTORY:
            raise ValidationError('Некорректная история докладов.')
        clean_history = []
        for old in history:
            if not isinstance(old, dict) or set(old) - {'report', 'at', 'by'}:
                raise ValidationError('Некорректная запись истории доклада.')
            clean_history.append({'report': _text(old.get('report', ''), MAX_REPORT, 'Старый доклад', strip=True),
                                  'report_at': _legacy_stamp(old.get('at')), 'report_by': _actor(old.get('by') or '')})
        result[(primary, secondary)] = {'resp': resp, 'report': report, 'theme': theme,
            'assigned_at': _legacy_stamp(item.get('resp_at')), 'assigned_by': _actor(item.get('resp_by') or ''),
            'report_at': _legacy_stamp(item.get('report_at')), 'report_by': _actor(item.get('report_by') or ''), 'history': clean_history}
    if len(seen) > MAX_PEOPLE:
        raise ValidationError(f'В справочнике может быть не более {MAX_PEOPLE} ответственных вместе с архивом.')
    return {'people': people, 'archived': archived, 'since': since, 'items': result}


def import_legacy(raw, actor=''):
    """Merge a legacy file transactionally. Never replace existing manual fields.

Repeated imports are harmless: per-field edited timestamps protect intentional clears,
old directory records retain their current active/archive state, cutoff only moves back.
"""
    data, actor = validate_legacy(raw), _actor(actor)
    counts = {'assignees_added': 0, 'entries_added': 0, 'entries_updated': 0, 'entries_preserved': 0}
    with _db(True) as conn:
        stamp = _stamp()
        persons = {row['name_key']: row['id'] for row in conn.execute('SELECT * FROM collective_assignees')}
        additions = sum(name.casefold() not in persons for name in data['people'] + data['archived'])
        if len(persons) + additions > MAX_PEOPLE:
            raise ValidationError(f'После импорта в справочнике будет более {MAX_PEOPLE} ответственных.')
        for archived, names in ((0, data['people']), (1, data['archived'])):
            for name in names:
                if name.casefold() not in persons:
                    cur = conn.execute('INSERT INTO collective_assignees(name,name_key,archived,created_at,updated_at,updated_by) VALUES(?,?,?,?,?,?)',
                                       (name, name.casefold(), archived, stamp, stamp, actor))
                    persons[name.casefold()] = cur.lastrowid
                    counts['assignees_added'] += 1
        old_since = conn.execute('SELECT since FROM collective_settings WHERE id=1').fetchone()[0]
        if data['since'] < old_since:
            conn.execute('UPDATE collective_settings SET since=? WHERE id=1', (data['since'],))
        for ids, item in data['items'].items():
            source = [''] * len(HEADERS)
            source[5], source[6] = ids
            source[14] = 'Импортировано без снимка источника'
            aliases = _aliases(source)
            existed = any(conn.execute('SELECT 1 FROM collective_identity_aliases WHERE alias=?', (alias,)).fetchone() for alias in aliases)
            key = _resolve(conn, source, stamp, legacy=True)
            row = conn.execute('SELECT * FROM collective_workflows WHERE complaint_key=?', (key,)).fetchone()
            old = _get_workflow(conn, key)
            changes = {}
            if row['assigned_at'] is None and row['assignee_id'] is None and item['resp']:
                changes.update(assignee_id=persons[item['resp'].casefold()], assigned_at=item['assigned_at'] or stamp,
                               assigned_by=item['assigned_by'] or actor)
            if row['report_at'] is None and not row['report'] and (item['report'] or item['history'] or item['report_at']):
                changes.update(report=item['report'], report_at=item['report_at'] or stamp, report_by=item['report_by'] or actor)
            if row['theme_at'] is None and row['theme'] is None and item['theme'] is not None:
                changes.update(theme=item['theme'], theme_at=stamp, theme_by=actor)
            if changes:
                # A fetched source row can already exist without any human work. In that
                # case preserve the legacy report history exactly as on a first import.
                virgin = row['version'] == 0 and all(row[field] is None for field in ('assigned_at', 'report_at', 'theme_at'))
                next_version = row['version']
                if not virgin:
                    _history(conn, key, old, actor, stamp)
                    next_version += 1
                if 'report' in changes:
                    for previous in item['history'][-MAX_HISTORY:]:
                        historical = dict(old, version=next_version, **previous)
                        _history(conn, key, historical, previous['report_by'] or actor, previous['report_at'] or stamp)
                        next_version += 1
                changes.update(version=max(row['version'] + 1, next_version), updated_at=stamp, updated_by=actor)
                conn.execute('UPDATE collective_workflows SET ' + ','.join(field + '=?' for field in changes) + ' WHERE complaint_key=?',
                             (*changes.values(), key))
                counts['entries_updated' if existed else 'entries_added'] += 1
            else:
                counts['entries_preserved'] += 1
        return {**counts, 'since': min(old_since, data['since'])}
