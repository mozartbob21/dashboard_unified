"""Per-module tokens in the existing encrypted credential store.

Runtime reads never migrate/create the accounts DB. Tokens are not exposed by
status APIs; an invalid encrypted override fails closed instead of silently
switching to a different account. Endpoint and model remain server configuration.
"""
import json
import os
import sqlite3
from fastapi import HTTPException
from core.ai_errors import AIServiceError
from services.auth import integrations
from utils import db

AI_MODULES = {
    'aichat': 'ИИ-чат и голосовой помощник',
    'summarizer': 'Сумматор',
    'appeals': 'Эмпатичные ответы',
    'tools': 'Полезные инструменты · макросы и презентации',
    'water-dashboard': 'Сводный дашборд · выводы ИИ',
}


def _service(module_id):
    if module_id not in AI_MODULES:
        raise HTTPException(404, 'Неизвестный блок ИИ')
    return 'ai:' + module_id


def _stored(module_id):
    return _read_service(_service(module_id))


def _read_service(service):
    if not db.DB_FILE.exists():
        return None
    try:
        conn = sqlite3.connect(db.DB_FILE.resolve().as_uri() + '?mode=ro', uri=True, timeout=10)
        try:
            if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='integration_credentials'").fetchone():
                return None
            row = conn.execute('SELECT encrypted_value FROM integration_credentials WHERE service=?', (service,)).fetchone()
        finally:
            conn.close()
        if row is None:
            return None
        value = json.loads(integrations.cipher().decrypt(row[0].encode()))
        token = value.get('token')
        if not isinstance(token, str) or not token:
            raise ValueError()
        return token
    except Exception:
        raise AIServiceError('AI_CREDENTIALS_ERROR') from None


def environment_key(module_id):
    if module_id == 'summarizer':
        return os.getenv('SUMMARIZER_QWEN_API_KEY', '').strip() or os.getenv('QWEN_API_KEY', '').strip()
    return os.getenv('QWEN_API_KEY', '').strip()


def resolve_key(module_id, fallback=None):
    if module_id is None:
        return (fallback or os.getenv('QWEN_API_KEY', '')).strip()
    return _stored(module_id) or (fallback or environment_key(module_id)).strip()


def status(module_id):
    _service(module_id)
    try:
        override = _stored(module_id)
        error = False
    except AIServiceError:
        override, error = None, True
    env = bool(environment_key(module_id))
    return {'module_id': module_id, 'label': AI_MODULES[module_id],
            'override': bool(override) or error, 'configured': bool(override) or env,
            'environment_configured': env, 'error': error,
            'source': 'error' if error else 'saved' if override else 'environment' if env else 'none'}


def statuses():
    return {'items': [status(module_id) for module_id in AI_MODULES]}


def save(module_id, token):
    service = _service(module_id)
    if not isinstance(token, str) or len(token) > 8192:
        raise HTTPException(400, 'Укажите токен длиной не более 8192 символов')
    token = token.strip()
    if not token:
        return status(module_id)  # Blank means keep, never clear accidentally.
    if any(ord(char) < 33 or ord(char) > 126 for char in token):
        raise HTTPException(400, 'Токен должен содержать только печатные латинские символы без пробелов')
    encrypted = integrations.cipher(create=True).encrypt(json.dumps({'token': token}).encode()).decode()
    with db.get_db_connection() as conn:
        conn.execute('INSERT INTO integration_credentials(service,encrypted_value) VALUES(?,?) '
                     'ON CONFLICT(service) DO UPDATE SET encrypted_value=excluded.encrypted_value,updated_at=CURRENT_TIMESTAMP',
                     (service, encrypted))
        conn.execute('INSERT INTO account_notifications(message) VALUES(?)',
                     (f'Обновлён токен ИИ: {AI_MODULES[module_id]}.',))
    return status(module_id)


def remove(module_id):
    service = _service(module_id)
    with db.get_db_connection() as conn:
        conn.execute('DELETE FROM integration_credentials WHERE service=?', (service,))
        conn.execute('INSERT INTO account_notifications(message) VALUES(?)',
                     (f'Удалён отдельный токен ИИ: {AI_MODULES[module_id]}. Используются настройки сервера.',))
    return status(module_id)
