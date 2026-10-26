"""API credentials for source integrations; independent of neural model keys."""
import json
import os
from fastapi import HTTPException
from core.ai_errors import AIServiceError
from services.auth import integrations
from services.auth.ai_credentials import _read_service
from utils import db

SOURCES = {'omsu_rating': ('Рейтинг ОМСУ', 'OMSU_RATING_API_KEY')}


def _service(source):
    if source not in SOURCES:
        raise HTTPException(404, 'Неизвестный источник')
    return 'source-api:' + source


def get_source_api_key(source):
    try:
        return _read_service(_service(source)) or os.getenv(SOURCES[source][1], '').strip()
    except AIServiceError:
        raise HTTPException(503, 'Не удалось прочитать API-ключ источника. Проверьте ключ шифрования или сохраните API-ключ заново.') from None


def status(source):
    service = _service(source)
    try:
        value, error = _read_service(service), False
    except AIServiceError:
        value, error = None, True
    env = bool(os.getenv(SOURCES[source][1], '').strip())
    return {'source': source, 'label': SOURCES[source][0], 'override': bool(value) or error,
            'configured': bool(value) or env, 'error': error, 'environment_configured': env}


def save(source, token):
    service = _service(source)
    if not isinstance(token, str) or len(token) > 8192:
        raise HTTPException(400, 'Укажите API-ключ длиной не более 8192 символов')
    token = token.strip()
    if not token:
        return status(source)
    if any(ord(char) < 33 or ord(char) > 126 for char in token):
        raise HTTPException(400, 'API-ключ должен содержать только печатные латинские символы без пробелов')
    encrypted = integrations.cipher(create=True).encrypt(json.dumps({'token': token}).encode()).decode()
    with db.get_db_connection() as conn:
        conn.execute('INSERT INTO integration_credentials(service,encrypted_value) VALUES(?,?) '
                     'ON CONFLICT(service) DO UPDATE SET encrypted_value=excluded.encrypted_value,updated_at=CURRENT_TIMESTAMP', (service,encrypted))
        conn.execute('INSERT INTO account_notifications(message) VALUES(?)', (f'Обновлён API-ключ: {SOURCES[source][0]}.',))
    return status(source)


def remove(source):
    service = _service(source)
    with db.get_db_connection() as conn:
        conn.execute('DELETE FROM integration_credentials WHERE service=?', (service,))
        conn.execute('INSERT INTO account_notifications(message) VALUES(?)', (f'Удалён отдельный API-ключ: {SOURCES[source][0]}.',))
    return status(source)
