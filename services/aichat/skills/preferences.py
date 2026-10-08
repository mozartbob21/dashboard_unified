"""User skill switches, stored only for the current authenticated account."""
import re

from fastapi import HTTPException
from services.auth.home_preferences import account_key
from utils.db import get_db_connection
from . import registry


def _schema(conn):
    conn.execute('''CREATE TABLE IF NOT EXISTS skill_preferences (
        account_key TEXT NOT NULL,
        skill_id TEXT NOT NULL,
        enabled INTEGER NOT NULL CHECK (enabled IN (0, 1)),
        PRIMARY KEY (account_key, skill_id)
    )''')


def _supported_tool_names():
    # Lazy import avoids a registry/preferences/tool initialization cycle.
    from .tools import SPECS
    return frozenset(SPECS)


def _catalog(key, snapshot):
    with get_db_connection() as conn:
        _schema(conn)
        rows = conn.execute('SELECT skill_id, enabled FROM skill_preferences WHERE account_key=?',
                            (key,)).fetchall()
    saved = {row['skill_id']: bool(row['enabled']) for row in rows}
    issues = list(snapshot['issues'])
    try:
        supported = _supported_tool_names()
    except ImportError:
        supported = frozenset()
        issues.append('Серверный реестр инструментов недоступен. Администратору нужно проверить установку.')
    items = []
    for descriptor in snapshot['items']:
        available = [tool for tool in descriptor['allowed_tools'] if tool in supported]
        unsupported = [tool for tool in descriptor['allowed_tools'] if tool not in supported]
        if unsupported:
            # Diagnostics never echo shell-like arguments or host paths from a pack.
            names = [tool if re.fullmatch(r'[A-Za-z0-9_.:-]{1,64}', tool) else 'нестандартный инструмент'
                     for tool in unsupported[:8]]
            suffix = ' и другие' if len(unsupported) > 8 else ''
            issues.append(f"Навык «{descriptor['id']}»: инструменты {', '.join(names)}{suffix} не поддерживаются сервером.")
        items.append({**descriptor, 'allowed_tools': available, 'unsupported_tools': unsupported,
                      'enabled': saved.get(descriptor['id'], True)})
    return {'items': items, 'issues': issues, 'scope': 'account'}


def catalog(user):
    key = account_key(user)
    return _catalog(key, registry.catalog_snapshot())


def enabled_skills(user):
    return [{key: value for key, value in item.items() if key not in {'enabled', 'unsupported_tools'}}
            for item in catalog(user)['items'] if item['enabled']]


def set_enabled(user, skill_id, enabled):
    key = account_key(user)
    if not isinstance(enabled, bool):
        raise HTTPException(400, 'Поле enabled должно быть логическим значением.')
    snapshot = registry.catalog_snapshot()
    if not isinstance(skill_id, str) or not any(item['id'] == skill_id for item in snapshot['items']):
        raise HTTPException(404, 'Навык не установлен или недоступен.')
    with get_db_connection() as conn:
        _schema(conn)
        conn.execute('''INSERT INTO skill_preferences(account_key, skill_id, enabled) VALUES (?, ?, ?)
            ON CONFLICT(account_key, skill_id) DO UPDATE SET enabled=excluded.enabled''',
            (key, skill_id, int(enabled)))
    return _catalog(key, snapshot)
