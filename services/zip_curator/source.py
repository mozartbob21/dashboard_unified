"""Server-selected stock source. Authenticated cloud credentials never reach UI."""
import os

from services.auth.integrations import credentials
from services.zip_curator.yandex_disk import DiskSyncError

DEFAULT_YANDEX_URL = 'https://disk.yandex.ru/d/WeAv-MDA6uEzpw'
DEFAULT_CLOUD_URL = 'https://cloud.mosreg.ru/apps/files/?dir=/&fileid=10968596'


def resolve_source(public_url=None):
    account = credentials('zip_cloud')
    if account:
        return {'provider': 'mosreg', 'url': account.get('source_url') or DEFAULT_CLOUD_URL,
                'account': {'username': account['username'], 'password': account['password']},
                'folder_path': account.get('folder_path') or None}
    url = str(public_url or os.getenv('ZIP_FOLDER_A_URL') or DEFAULT_YANDEX_URL).strip()
    if 'cloud.mosreg.ru' in url.lower():
        raise DiskSyncError('Сохраните доступ «Облако МО — ЗиП» в разделе «Пользователи → Логины и пароли».')
    return {'provider': 'yandex', 'url': url}


def public_status():
    selected = resolve_source()
    return {'provider': selected['provider'], 'configured': True,
            'label': 'Облако МО' if selected['provider'] == 'mosreg' else 'Яндекс Диск'}


def check_cloud_access():
    """Validate the exact saved folder; no stock imports or publication."""
    import requests
    from services.zip_curator.mosreg_cloud import CloudClient
    selected = resolve_source()
    if selected['provider'] != 'mosreg':
        raise DiskSyncError('Сначала сохраните доступ «Облако МО — ЗиП».')
    with requests.Session() as session:
        client = CloudClient(session, selected['url'], username=selected['account']['username'],
                             password=selected['account']['password'], folder_path=selected['folder_path'])
        files = client.list_files()
    return {'ok': True, 'files': len(files),
            'message': f'Доступ к указанной папке Облака МО подтверждён. Excel-файлов: {len(files)}. Данные ещё не загружались.'}
