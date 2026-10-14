"""Source dates and provider selection using synthetic workbooks only."""
import copy
import hashlib
import json
import sys
from datetime import datetime
from types import SimpleNamespace

import pytest
from services.zip_curator import core
from test.test_zip_curator_folder_sync import (
    local_store, workbook, remote_file, listing, cache, use_session, read_manifest, API, PUBLIC_URL, FakeResponse,
)

CREATED = '2025-02-03T08:15:00Z'
MODIFIED = '2026-10-13T10:00:00Z'
CLOUD_URL = 'https://cloud.mosreg.ru/index.php/apps/files/?dir=/FolderA'
ACCOUNT = {'username': 'synthetic-cloud-user', 'password': 'synthetic-cloud-password'}


def cloud(monkeypatch, items, contents):
    class Cloud:
        downloads = []
        calls = []
        def __init__(self, session, source_url, timeout=45, *, username, password, folder_path=None):
            self.calls.append({'url': source_url, 'username': username, 'password': password, 'folder_path': folder_path})
        def list_files(self):
            return copy.deepcopy(items)
        def download(self, item, read_rows):
            self.downloads.append(item['path'])
            data = contents[item['path']]
            if isinstance(data, Exception):
                raise data
            core.validate_workbook(data, item, read_rows)
            return data
    monkeypatch.setitem(sys.modules, 'services.zip_curator.mosreg_cloud', SimpleNamespace(CloudClient=Cloud))
    return Cloud


def cloud_item(data, **extra):
    return {'name': 'registry.xlsx', 'type': 'file', 'path': '/FolderA/registry.xlsx',
            'size': len(data), 'etag': 'opaque-etag-not-md5', 'created': CREATED, 'modified': MODIFIED, **extra}


def sync_cloud(**extra):
    return core.sync_and_scan_folder_a(CLOUD_URL, provider='mosreg', account=ACCOUNT, folder_path='/FolderA', **extra)


def test_cloud_dates_are_separate_and_survive_approval_publication_restore(monkeypatch, local_store):
    data = workbook(local_store)
    item = cloud_item(data)
    fake = cloud(monkeypatch, [item], {item['path']: data})
    result = sync_cloud()
    pending = result['state']['pending'][0]
    assert pending['source_created_at'] == '2025-02-03T08:15:00+00:00'
    assert pending['source_modified_at'] == '2026-10-13T10:00:00+00:00'
    assert pending['loaded_at'] not in {pending['source_created_at'], pending['source_modified_at']}
    assert pending['source_etag'] == item['etag']
    assert 'source_md5' not in pending
    manifest = read_manifest()
    assert manifest['files'][0]['created'] == pending['source_created_at']
    assert manifest['files'][0]['modified'] == pending['source_modified_at']
    assert manifest['files'][0]['md5'] == ''
    assert fake.calls[0]['folder_path'] == '/FolderA'
    serialized = json.dumps({'result': result, 'manifest': manifest})
    assert ACCOUNT['username'] not in serialized and ACCOUNT['password'] not in serialized
    assert core.approve(0)
    core.publish_clean()
    key = core.norm(pending['rso'])
    approved = copy.deepcopy(core.load_state()['clean'][key])
    core.STATE_FILE.unlink()
    assert core.load_state()['clean'][key] == approved


def test_yandex_creation_is_kept_separately_from_modification(monkeypatch, local_store):
    data = workbook(local_store)
    item = remote_file('registry.xlsx', data)
    item['created'] = CREATED
    use_session(monkeypatch, {(API, 0): [listing([item])], item['file']: [FakeResponse(content=data)]})
    result = core.sync_and_scan_folder_a(PUBLIC_URL)
    row = result['state']['pending'][0]
    assert row['source_provider'] == 'yandex'
    assert row['source_created_at'] == '2025-02-03T08:15:00+00:00'
    assert row['source_modified_at'] == '2026-10-13T10:00:00+00:00'


def test_unchanged_cloud_file_backfills_dates_without_overwriting_curator_edits(monkeypatch, local_store):
    data = workbook(local_store)
    item = cloud_item(data, created=None)
    items = [item]
    fake = cloud(monkeypatch, items, {item['path']: data})
    first = sync_cloud()['state']
    first['pending'][0]['items'][0]['qty'] = 77
    first['pending'][0]['items'][0]['via'] = 'override'
    loaded_at = first['pending'][0]['loaded_at']
    core.save_state(first)
    item['created'] = CREATED
    result = sync_cloud()
    pending = result['state']['pending'][0]
    assert result['reused'] == 1 and result['downloaded'] == 0 and result['added'] == 0
    assert len(fake.downloads) == 1
    assert pending['items'][0]['qty'] == 77 and pending['items'][0]['via'] == 'override'
    assert pending['loaded_at'] == loaded_at
    assert pending['source_created_at'] == '2025-02-03T08:15:00+00:00'


def test_equal_cloud_etag_with_changed_size_requires_download(monkeypatch, local_store):
    old = workbook(local_store, qty=3)
    new = workbook(local_store, qty=123456789012345)
    assert len(old) != len(new)
    item = cloud_item(old)
    contents = {item['path']: old}
    fake = cloud(monkeypatch, [item], contents)
    sync_cloud()
    item['size'] = len(new)
    contents[item['path']] = new
    result = sync_cloud()
    assert result['downloaded'] == 1 and result['reused'] == 0
    assert len(fake.downloads) == 2
    assert result['state']['pending'][0]['items'][0]['qty'] == 123456789012345


def test_switching_cloud_account_never_reuses_or_overwrites_old_original(monkeypatch, local_store):
    old = workbook(local_store, qty=1)
    new = workbook(local_store, qty=9)
    item = cloud_item(old)
    contents = {item['path']: old}
    fake = cloud(monkeypatch, [item], contents)
    first = sync_cloud()['state']['pending'][0]
    assert core.approve(0)
    approved = copy.deepcopy(core.load_state()['clean'])
    original = core.INPUT_DIR / first['local_file']
    item['size'] = len(new)
    contents[item['path']] = new
    result = core.sync_and_scan_folder_a(CLOUD_URL, provider='mosreg',
        account={**ACCOUNT, 'username': 'another-account'}, folder_path='/FolderA')
    assert result['downloaded'] == 1 and result['reused'] == 0
    assert len(fake.downloads) == 2
    assert result['state']['clean'] == approved
    assert original.read_bytes() == old
    assert read_manifest()['files'][0]['local_name'] != original.name


def test_failed_cloud_update_retains_previous_dates_and_curator_edits(monkeypatch, local_store):
    data = workbook(local_store)
    item = cloud_item(data)
    contents = {item['path']: data}
    cloud(monkeypatch, [item], contents)
    first = sync_cloud()['state']
    first['pending'][0]['items'][0]['qty'] = 77
    core.save_state(first)
    before = copy.deepcopy(core.load_state()['pending'][0])
    item.update(etag='new-version', created='2026-10-14T12:00:00Z', modified='2026-10-14T12:00:00Z')
    contents[item['path']] = core.DiskSyncError('Источник временно недоступен.')
    result = sync_cloud()
    assert result['partial'] and result['errors'][0]['retained']
    assert result['state']['pending'][0] == before
    assert read_manifest()['files'][0]['created'] == before['source_created_at']


def test_legacy_manifest_backfill_updates_metadata_only_and_survives_restore(monkeypatch, local_store):
    data = workbook(local_store)
    item = remote_file('legacy.xlsx', data)
    cache([(item, data)])
    manifest = read_manifest()
    manifest['files'][0]['created'] = CREATED
    core.FOLDER_A_MANIFEST_FILE.write_text(json.dumps(manifest))
    record = core.parse_reestr(core.read_xlsx_rows(data), 'legacy.xlsx')
    core._enrich(record)
    record.update(local_file='folder_a/legacy.xlsx', fname='legacy.xlsx',
                  uploaded=int(datetime.fromisoformat(MODIFIED.replace('Z', '+00:00')).timestamp()*1000),
                  loaded_at='2026-10-14T11:00:00+03:00')
    record['items'][0]['qty'] = 77
    state = {'pending': [], 'clean': {core.norm(record['rso']): record}, 'excluded_rso': {}}
    core.save_state(state)
    rows = core.export_rows_from_state(state)
    core.PUBLISHED_FILE.write_text(json.dumps({'rows': rows, 'published_at': 'original-publication',
        'rso_metadata': {core.norm(record['rso']): {k: record[k] for k in ('local_file','fname','uploaded','loaded_at')}}}))
    loaded = core.load_state()['clean'][core.norm(record['rso'])]
    assert loaded['source_created_at'] == '2025-02-03T08:15:00+00:00'
    assert loaded['items'][0]['qty'] == 77
    published = json.loads(core.PUBLISHED_FILE.read_text())
    assert published['rows'] == rows and published['published_at'] == 'original-publication'
    core.STATE_FILE.unlink()
    restored = core.load_state()['clean'][core.norm(record['rso'])]
    assert restored['source_created_at'] == loaded['source_created_at']
    assert restored['source_modified_at'] == loaded['source_modified_at']
    assert restored['loaded_at'] == record['loaded_at']


@pytest.mark.parametrize('override', [
    {'local_file': 'unrelated/legacy.xlsx'}, {'source_identity': 'different-source'},
    {'source_path': '/different-file.xlsx'}, {'source_md5': 'different-version'}, {'uploaded': 1234567890000},
])
def test_backfill_rejects_mismatched_source_or_version(monkeypatch, local_store, override):
    data = workbook(local_store)
    item = remote_file('legacy.xlsx', data)
    cache([(item, data)])
    record = {'rso':'Legacy', 'items':[], 'local_file':'folder_a/legacy.xlsx', **override}
    core.save_state({'pending':[record], 'clean':{}, 'excluded_rso':{}})
    loaded = core.load_state()['pending'][0]
    assert 'source_created_at' not in loaded and 'source_modified_at' not in loaded


def test_unknown_legacy_dates_do_not_use_local_mtime_or_neurona_load_time(monkeypatch, local_store):
    record = {'rso':'Legacy', 'items':[], 'fname':'legacy.xlsx', 'local_file':'folder_a/legacy.xlsx',
              'uploaded':1700000000000, 'loaded_at':'2026-10-14T11:00:00+03:00'}
    core.FOLDER_A_DIR.mkdir(parents=True)
    (core.FOLDER_A_DIR/'legacy.xlsx').write_bytes(workbook(local_store))
    core.save_state({'pending':[record], 'clean':{}, 'excluded_rso':{}})
    loaded = core.load_state()['pending'][0]
    assert 'source_created_at' not in loaded and 'source_modified_at' not in loaded


@pytest.mark.parametrize('value', [None, 1700000000, '', '2026-10-01', 'not-a-date', '2026-02-30T12:00:00Z'])
def test_only_valid_source_datetimes_are_accepted(value):
    assert core._source_timestamp(value) is None


def test_creation_only_is_not_mislabeled_as_modification(monkeypatch, local_store):
    data = workbook(local_store)
    item = cloud_item(data, modified=None)
    cloud(monkeypatch, [item], {item['path']: data})
    record = sync_cloud()['state']['pending'][0]
    assert record['source_created_at'] == '2025-02-03T08:15:00+00:00'
    assert 'source_modified_at' not in record
    assert read_manifest()['files'][0]['modified'] is None


def test_cloud_requires_server_credentials_before_network(monkeypatch, local_store):
    def unexpected():
        raise AssertionError('No request may be made without configured credentials')
    monkeypatch.setattr(core.requests, 'Session', unexpected)
    with pytest.raises(core.DiskSyncError, match='логин и пароль'):
        core.sync_and_scan_folder_a(CLOUD_URL, provider='mosreg')
