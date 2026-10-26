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
    result = {'username':value['username'] if value else '', 'configured':bool(value)}
    if service == 'zip_cloud':
        from services.zip_curator.source import DEFAULT_CLOUD_URL
        result.update(source_url=(value or {}).get('source_url') or DEFAULT_CLOUD_URL,
                      folder_path=(value or {}).get('folder_path') or '')
    return result


def save_credentials(username, password, service='edds', *, source_url=None, folder_path=None):
    username=username.strip()
    if service not in {'edds','edds_arm','mingkh','zip_cloud','presentations'}: raise HTTPException(400,'Неизвестная интеграция')
    if not username: raise HTTPException(400,'Укажите логин системы')
    old = None
    if not password:
        old = credentials(service)
    elif service == 'zip_cloud':
        try:
            old = credentials(service)
        except HTTPException:
            pass  # Explicit new credentials can recover a lost encryption key.
    if not password:
        if not old or old['username']!=username:
            raise HTTPException(400,'При первом сохранении или смене логина нужен пароль')
        password=old['password']
    value = {'username': username, 'password': password}
    if service == 'zip_cloud':
        from services.zip_curator.source import DEFAULT_CLOUD_URL
        from services.zip_curator.mosreg_cloud import normalize_source_url, normalize_folder_path, CloudSyncError
        try:
            value['source_url'] = normalize_source_url(source_url if source_url is not None else (old or {}).get('source_url') or DEFAULT_CLOUD_URL)
            value['folder_path'] = normalize_folder_path(folder_path if folder_path is not None else (old or {}).get('folder_path') or '')
        except CloudSyncError as exc:
            raise HTTPException(400, str(exc)) from None
    elif source_url is not None or folder_path is not None:
        raise HTTPException(400, 'Папка источника настраивается только для Облака МО.')
    encrypted=cipher(create=True).encrypt(json.dumps(value).encode()).decode()
    with get_db_connection() as conn:
        conn.execute("INSERT INTO integration_credentials(service,encrypted_value) VALUES(?,?) ON CONFLICT(service) DO UPDATE SET encrypted_value=excluded.encrypted_value,updated_at=CURRENT_TIMESTAMP",(service,encrypted))
        label={'edds':'Добродел','edds_arm':'АРМ ЕДДС','mingkh':'МИНЖКХ','zip_cloud':'Облако МО — ЗиП','presentations':'Еженедельные презентации'}[service]
        conn.execute("INSERT INTO account_notifications(message) VALUES (?)",(f'Обновлены настройки доступа: {label}.',))
