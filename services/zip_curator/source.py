"""Complementary stock sources; credentials are never included in public status."""
import os
from urllib.parse import urlsplit

from fastapi import HTTPException
from services.auth.integrations import credentials
from services.zip_curator.yandex_disk import DiskSyncError

DEFAULT_YANDEX_URL = 'https://disk.yandex.ru/d/WeAv-MDA6uEzpw'
DEFAULT_CLOUD_URL = 'https://cloud.mosreg.ru/apps/files/?dir=/&fileid=10968596'
SOURCE_LABELS = {'yandex': 'Яндекс Диск', 'mosreg': 'Облако МО'}


def _is_cloud_url(value):
    try:
        return urlsplit(value).hostname == 'cloud.mosreg.ru'
    except ValueError:
        return False


def _cloud_source():
    account = credentials('zip_cloud')
    if not account:
        return None
    if not isinstance(account, dict) or not all(isinstance(account.get(k), str) and account[k] for k in ('username', 'password')):
        raise DiskSyncError('Проверьте сохранённые настройки «Облако МО — ЗиП» в админ-панели.')
    return {'provider': 'mosreg', 'label': SOURCE_LABELS['mosreg'],
            'url': account.get('source_url') or DEFAULT_CLOUD_URL,
            'account': {'username': account['username'], 'password': account['password']},
            'folder_path': account.get('folder_path') or None}


def resolve_sources(public_url=None):
    """Cloud complements Yandex, including when cloud settings cannot be decrypted."""
    url = str(public_url or os.getenv('ZIP_FOLDER_A_URL') or DEFAULT_YANDEX_URL).strip()
    cloud_requested = _is_cloud_url(url)
    # A legacy cloud URL must never be passed to the public Yandex API, and must
    # not suppress the original source. Cloud access uses the saved credentials.
    sources = [{'provider': 'yandex', 'label': SOURCE_LABELS['yandex'],
                'url': DEFAULT_YANDEX_URL if cloud_requested else url}]
    try:
        cloud = _cloud_source()
    except (HTTPException, DiskSyncError):
        cloud = {'provider': 'mosreg', 'label': SOURCE_LABELS['mosreg'],
                 'error': 'Не удалось прочитать настройки Облака МО. Администратору нужно проверить сохранённый доступ.'}
    if cloud:
        sources.append(cloud)
    elif cloud_requested:
        sources.append({'provider': 'mosreg', 'label': SOURCE_LABELS['mosreg'],
                        'error': 'Сохраните доступ «Облако МО — ЗиП» в разделе «Логины и пароли». Данные Яндекс Диска обновляются отдельно.'})
    return sources


def resolve_source(public_url=None):
    """Compatibility helper for the original, always available Yandex source."""
    return resolve_sources(public_url)[0]


def public_status():
    selected = resolve_sources()
    available = {item['provider']: item for item in selected}
    sources = [{'provider': key, 'label': label,
                'configured': key in available and not available[key].get('error'),
                **({'error': available[key]['error']} if available.get(key, {}).get('error') else {})}
               for key, label in SOURCE_LABELS.items()]
    return {'provider': 'combined' if len(selected) > 1 else 'yandex', 'configured': True,
            'label': ' + '.join(item['label'] for item in selected), 'sources': sources,
            **({'warning': ' '.join(item['error'] for item in selected if item.get('error'))}
               if any(item.get('error') for item in selected) else {})}


def check_cloud_access():
    """Probe only the exact saved cloud folder; do not import or touch Yandex."""
    import requests
    from services.zip_curator.mosreg_cloud import CloudClient
    selected = _cloud_source()
    if selected is None:
        raise DiskSyncError('Сначала сохраните доступ «Облако МО — ЗиП».')
    with requests.Session() as session:
        client = CloudClient(session, selected['url'], username=selected['account']['username'],
                             password=selected['account']['password'], folder_path=selected['folder_path'])
        files = client.list_files()
    return {'ok': True, 'files': len(files),
            'message': f'Доступ к указанной папке Облака МО подтверждён. Excel-файлов: {len(files)}. При обновлении они дополнят данные Яндекс Диска.'}
