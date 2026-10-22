"""Current local-account rights shared by MAX entry points and the polling worker."""
from contextlib import closing
import json
from pathlib import Path
import sqlite3


def read_account(path, user_id):
    """Read existing account storage only; never migrate or seed it from the bot."""
    uri = Path(path).resolve().as_uri() + '?mode=ro'
    with closing(sqlite3.connect(uri, uri=True, timeout=3)) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT u.id,u.username,u.role,u.modules,u.is_active,u.password_hash, "
            "EXISTS(SELECT 1 FROM account_control c WHERE c.id=1 AND "
            "(c.manager_user_id=u.id OR EXISTS(SELECT 1 FROM account_managers m WHERE m.user_id=u.id))) "
            "AS can_manage_users FROM users u WHERE u.id=?", (user_id,),
        ).fetchone()
    if row is None:
        return None
    account = dict(row)
    try:
        modules = json.loads(account['modules'] or '[]')
    except (TypeError, ValueError):
        modules = []
    account['modules'] = modules if isinstance(modules, list) else []
    return account


def denial(account):
    """A displayed role or a client-side permission flag never grants access."""
    if not account or not account['is_active']:
        return 'ACCESS_REVOKED', 'Доступ отозван. Обратитесь к администратору Нейроны.'
    if 'zips' not in account['modules']:
        return 'MODULE_FORBIDDEN', 'Нет доступа к Остаткам ЗиП РСО.'
    if not account['can_manage_users']:
        return 'MANAGER_REQUIRED', 'Подключение MAX пока доступно только пользователям с правом «Управление пользователями».'
    return None
