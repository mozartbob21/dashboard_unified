"""Read encrypted source credentials only when a check is explicitly requested."""
from contextlib import closing
import json
import sqlite3

from .presentation_source import RedmineSource, SourceError


def source_factory():
    from services.auth.integrations import cipher
    from utils.db import DB_FILE
    try:
        # The standalone bot must not initialize/migrate the platform database.
        # Credentials are written only by the existing administrator endpoint.
        uri = DB_FILE.resolve().as_uri() + '?mode=ro'
        with closing(sqlite3.connect(uri, uri=True, timeout=3)) as conn:
            row = conn.execute("SELECT encrypted_value FROM integration_credentials WHERE service='presentations'").fetchone()
        value = json.loads(cipher().decrypt(row[0].encode())) if row else None
    except Exception:
        raise SourceError('SOURCE_NOT_CONFIGURED', 'Проверьте сохранённый доступ к источнику презентаций в админ-панели.') from None
    if not isinstance(value, dict) or not all(isinstance(value.get(k), str) and value[k] for k in ('username', 'password')):
        raise SourceError('SOURCE_NOT_CONFIGURED', 'Администратору нужно сохранить логин и пароль источника презентаций.')
    return RedmineSource(value['username'], value['password'])
