"""Account-scoped module favorites. Access is recalculated on every request."""
import json

from fastapi import HTTPException
from core.roles import effective_modules
from utils.db import get_db_connection

FAVORITABLE = frozenset({
    'edo', 'overdue', 'watercontrol', 'utnkr', 'cameras', 'appeals',
    'mingkh', 'summarizer', 'cds', 'mgkh_rm', 'telegram', 'zips', 'ecur', 'edds',
})


def account_key(user):
    if not user:
        raise HTTPException(401, 'Требуется авторизация')
    if user.get('kc_sub'):
        return 'keycloak:' + str(user['kc_sub'])
    if user.get('id') is not None:
        return 'local:' + str(user['id'])
    raise HTTPException(401, 'Не удалось определить учётную запись')


def preferences(user, favorites=None):
    key = account_key(user)
    modules = set(effective_modules(user))
    enabled = len(modules) > 3
    eligible = sorted(modules & FAVORITABLE) if enabled else []
    if favorites is not None:
        if not enabled:
            raise HTTPException(403, 'Избранное доступно, когда аккаунту разрешено больше трёх блоков')
        if any(module not in eligible for module in favorites):
            raise HTTPException(403, 'Этот блок нельзя добавить в избранное')
    with get_db_connection() as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS home_favorites (
            account_key TEXT PRIMARY KEY, modules TEXT NOT NULL DEFAULT '[]')''')
        if favorites is not None:
            conn.execute('''INSERT INTO home_favorites(account_key, modules) VALUES (?, ?)
                ON CONFLICT(account_key) DO UPDATE SET modules=excluded.modules''',
                (key, json.dumps(list(dict.fromkeys(favorites)))))
        row = conn.execute('SELECT modules FROM home_favorites WHERE account_key=?', (key,)).fetchone()
    try:
        saved = json.loads(row['modules']) if row else []
    except (ValueError, TypeError):
        saved = []
    if not isinstance(saved, list):
        saved = []
    return {'can_favorite': enabled, 'eligible': eligible,
            'favorites': list(dict.fromkeys(m for m in saved if isinstance(m, str) and m in eligible))}
