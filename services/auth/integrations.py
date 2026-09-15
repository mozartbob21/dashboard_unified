"""Encrypted integration credentials; the local encryption key never goes to Git."""
import json
import os
from pathlib import Path
from cryptography.fernet import Fernet
from fastapi import HTTPException
from utils.db import get_db_connection

KEY_FILE = Path(__file__).resolve().parents[2] / 'data/auth/integrations.key'


def cipher(create=False):
    if create:
        KEY_FILE.parent.mkdir(parents=True, exist_ok=True)
        try:
            fd=os.open(KEY_FILE, os.O_WRONLY|os.O_CREAT|os.O_EXCL, 0o600)
        except FileExistsError:
            pass
        else:
            with os.fdopen(fd,'wb') as f: f.write(Fernet.generate_key())
    return Fernet(KEY_FILE.read_bytes())


def credentials(service='edds'):
    with get_db_connection() as conn:
        row=conn.execute('SELECT encrypted_value FROM integration_credentials WHERE service=?',(service,)).fetchone()
    if not row: return None
    try: return json.loads(cipher().decrypt(row['encrypted_value'].encode()))
    except Exception:
        raise HTTPException(503,'Не удалось прочитать настройки доступа. Проверьте ключ integrations.key или заново сохраните логин и пароль.')


def credential_status(service='edds'):
    value=credentials(service)
    return {'username':value['username'] if value else '', 'configured':bool(value)}


def save_credentials(username, password, service='edds'):
    username=username.strip()
    if not username: raise HTTPException(400,'Укажите логин Добродела')
    if not password:
        old=credentials(service)
        if not old or old['username']!=username:
            raise HTTPException(400,'При первом сохранении или смене логина нужен пароль')
        password=old['password']
    encrypted=cipher(create=True).encrypt(json.dumps({'username':username,'password':password}).encode()).decode()
    with get_db_connection() as conn:
        conn.execute("INSERT INTO integration_credentials(service,encrypted_value) VALUES(?,?) ON CONFLICT(service) DO UPDATE SET encrypted_value=excluded.encrypted_value,updated_at=CURRENT_TIMESTAMP",(service,encrypted))
        label='Добродел' if service=='edds' else 'АРМ ЕДДС'
        conn.execute("INSERT INTO account_notifications(message) VALUES (?)",(f'Обновлены настройки доступа: {label}.',))
