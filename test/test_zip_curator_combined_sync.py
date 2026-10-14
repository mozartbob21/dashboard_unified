"""Independent source refresh, content deduplication and curator-state protection."""
import copy
import hashlib
import io
import json
import sys
import zipfile
from types import SimpleNamespace

import pytest

from services.zip_curator import core
from test.test_zip_curator_folder_sync import local_store, workbook, remote_file, cache, PUBLIC_URL

CLOUD_URL = 'https://cloud.mosreg.ru/apps/files/?dir=/&fileid=10968596'
ACCOUNT = {'username': 'test-secret-user', 'password': 'test-secret-password'}
SOURCES = [
    {'provider': 'yandex', 'url': PUBLIC_URL},
    {'provider': 'mosreg', 'url': CLOUD_URL, 'account': ACCOUNT, 'folder_path': '/FolderA'},
]


@pytest.fixture
def sources(monkeypatch):
    fixtures = {provider: {'files': [], 'bytes': {}, 'downloads': []} for provider in ('yandex', 'mosreg')}

    class Client:
        provider = 'yandex'
        def __init__(self, *args, **kwargs):
            self.fixture = fixtures[self.provider]
        def list_files(self):
            files = self.fixture['files']
            if isinstance(files, Exception):
                raise files
            return copy.deepcopy(files)
        def download(self, item, read_rows):
            self.fixture['downloads'].append(item['path'])
            data = self.fixture['bytes'][item['path']]
            if isinstance(data, Exception):
                raise data
            core.validate_workbook(data, item, read_rows)
            return data
    class Cloud(Client):
        provider = 'mosreg'
    monkeypatch.setattr(core, 'DiskClient', Client)
    monkeypatch.setitem(sys.modules, 'services.zip_curator.mosreg_cloud', SimpleNamespace(CloudClient=Cloud))
    return fixtures


def add_source(fixtures, provider, name, data, *, modified='2026-10-14T12:00:00Z', created=None):
    item = remote_file(name, data, modified=modified)
    if created:
        item['created'] = created
    if provider == 'mosreg':
        item['etag'] = item.pop('md5')
    fixtures[provider]['files'].append(item)
    fixtures[provider]['bytes'][item['path']] = data
    return item


def state_by_rso(result):
    return {row['rso']: row for row in result['state']['pending']}


def test_disjoint_sources_form_union_and_reuse_each_manifest(monkeypatch, local_store, sources):
    add_source(sources, 'yandex', 'registry.xlsx', workbook(local_store, 'РСО Яндекс', 3))
    add_source(sources, 'mosreg', 'registry.xlsx', workbook(local_store, 'РСО Облако', 5))
    first = core.sync_and_scan_sources(SOURCES)
    assert first['ok'] and not first['partial']
    assert first['files'] == first['downloaded'] == first['added'] == 2
    rows = state_by_rso(first)
    assert set(rows) == {'РСО Яндекс', 'РСО Облако'}
    assert len({row['local_file'] for row in rows.values()}) == 2
    manifest = json.loads(core.FOLDER_A_MANIFEST_FILE.read_text())
    assert len(manifest['sources']) == 2
    second = core.sync_and_scan_sources(SOURCES)
    assert second['reused'] == 2 and second['downloaded'] == second['added'] == 0
    assert second['state'] == first['state']


def test_identical_bytes_different_names_parse_once_and_do_not_duplicate(monkeypatch, local_store, sources):
    data = workbook(local_store)
    add_source(sources, 'yandex', 'one.xlsx', data)
    add_source(sources, 'mosreg', 'renamed.xlsx', data)
    original = core.parse_reestr
    calls = []
    def parse(rows, fname):
        calls.append(fname)
        return original(rows, fname)
    monkeypatch.setattr(core, 'parse_reestr', parse)
    result = core.sync_and_scan_sources(SOURCES)
    assert result['duplicates'] == 1 and result['added'] == 1
    assert len(calls) == 1 and len(result['state']['pending']) == 1
    assert result['state']['pending'][0]['items'][0]['qty'] == 3


def test_same_inventory_with_different_workbook_metadata_deduplicates(local_store, sources):
    data = workbook(local_store)
    output = io.BytesIO(data)
    with zipfile.ZipFile(output, 'a') as archive:
        archive.writestr('docProps/custom.xml', '<custom>second source metadata</custom>')
    distinct = output.getvalue()
    assert hashlib.sha256(data).digest() != hashlib.sha256(distinct).digest()
    add_source(sources, 'yandex', 'one.xlsx', data)
    add_source(sources, 'mosreg', 'two.xlsx', distinct)
    result = core.sync_and_scan_sources(SOURCES)
    assert result['added'] == 1 and result['duplicates'] == 1


def test_same_items_for_different_explicit_organizations_are_not_merged(local_store, sources):
    add_source(sources, 'yandex', 'one.xlsx', workbook(local_store, 'РСО Первое'))
    add_source(sources, 'mosreg', 'two.xlsx', workbook(local_store, 'РСО Второе'))
    result = core.sync_and_scan_sources(SOURCES)
    assert result['added'] == 2 and result['duplicates'] == 0


def test_second_provider_cannot_reimport_unchanged_file_over_curator_edits(local_store, sources):
    data = workbook(local_store)
    add_source(sources, 'yandex', 'one.xlsx', data)
    first = core.sync_and_scan_sources(SOURCES)
    edited = first['state']
    edited['pending'][0]['items'][0].update(qty=77, via='override', cat='Ручная категория')
    core.save_state(edited)
    add_source(sources, 'mosreg', 'renamed.xlsx', data, modified='2026-10-15T12:00:00Z')
    second = core.sync_and_scan_sources(SOURCES)
    assert second['added'] == 0
    assert second['state']['pending'][0]['items'] == edited['pending'][0]['items']
    assert second['state']['pending'][0]['loaded_at'] == edited['pending'][0]['loaded_at']
    assert second['state']['pending'][0]['source_provider'] == 'mosreg'
    assert second['state']['pending'][0]['source_modified_at'] == '2026-10-15T12:00:00+00:00'


@pytest.mark.parametrize('failed_provider', ['yandex', 'mosreg'])
def test_one_source_failure_does_not_block_other_or_change_previous_data(local_store, sources, failed_provider):
    healthy = 'mosreg' if failed_provider == 'yandex' else 'yandex'
    add_source(sources, failed_provider, 'old.xlsx', workbook(local_store, 'Старое РСО'))
    first = core.sync_and_scan_sources(SOURCES)
    edited = first['state']
    edited['pending'][0]['items'][0]['qty'] = 77
    core.save_state(edited)
    sources[failed_provider]['files'] = core.DiskSyncError('Источник недоступен')
    add_source(sources, healthy, 'good.xlsx', workbook(local_store, 'Новое РСО', 8))
    result = core.sync_and_scan_sources(SOURCES)
    assert result['ok'] and result['partial'] and result['downloaded'] == 1
    assert result['failed'] == 0
    assert state_by_rso(result)['Старое РСО'] == edited['pending'][0]
    assert state_by_rso(result)['Новое РСО']['items'][0]['qty'] == 8
    assert {s['provider'] for s in result['sources'] if s['ok']} == {healthy}


def test_both_failed_sources_leave_state_and_manifests_unchanged(local_store, sources):
    add_source(sources, 'yandex', 'old.xlsx', workbook(local_store))
    core.sync_and_scan_sources(SOURCES)
    state = core.STATE_FILE.read_bytes()
    manifest = core.FOLDER_A_MANIFEST_FILE.read_bytes()
    for source in sources.values():
        source['files'] = core.DiskSyncError('Источник недоступен')
    result = core.sync_and_scan_sources(SOURCES)
    assert not result['ok'] and result['partial']
    assert result['added'] == 0
    assert core.STATE_FILE.read_bytes() == state
    assert core.FOLDER_A_MANIFEST_FILE.read_bytes() == manifest


def test_bad_cloud_configuration_retains_healthy_yandex_and_sanitizes_errors(local_store, sources):
    add_source(sources, 'yandex', 'good.xlsx', workbook(local_store))
    bad = {**SOURCES[1], 'error': 'Bad path for test-secret-user test-secret-password https://cloud.mosreg.ru/?token=TOP_SECRET'}
    result = core.sync_and_scan_sources([SOURCES[0], bad])
    assert result['ok'] and result['partial'] and result['added'] == 1
    serialized = json.dumps(result)
    assert ACCOUNT['username'] not in serialized and ACCOUNT['password'] not in serialized
    assert 'TOP_SECRET' not in serialized


def test_pruning_one_source_never_deletes_other_provider_cache(local_store, sources):
    add_source(sources, 'yandex', 'registry.xlsx', workbook(local_store, 'Первое РСО'))
    add_source(sources, 'mosreg', 'registry.xlsx', workbook(local_store, 'Второе РСО'))
    first = core.sync_and_scan_sources(SOURCES)
    cloud_record = state_by_rso(first)['Второе РСО']
    cloud_path = core.INPUT_DIR / cloud_record['local_file']
    before = cloud_path.read_bytes()
    sources['yandex']['files'] = []
    second = core.sync_and_scan_sources(SOURCES)
    assert second['removed'] == 1 and second['reused'] == 1
    assert cloud_path.read_bytes() == before
    assert state_by_rso(second)['Второе РСО'] == cloud_record


def test_legacy_manifest_is_migrated_and_reused_with_cloud(local_store, sources):
    data = workbook(local_store)
    item = remote_file('legacy.xlsx', data)
    cache([(item, data)])
    sources['yandex']['files'].append(item)
    sources['yandex']['bytes'][item['path']] = data
    add_source(sources, 'mosreg', 'cloud.xlsx', workbook(local_store, 'Облако РСО'))
    result = core.sync_and_scan_sources(SOURCES)
    assert result['reused'] == 1 and result['downloaded'] == 1
    assert sources['yandex']['downloads'] == []
    assert len(json.loads(core.FOLDER_A_MANIFEST_FILE.read_text())['sources']) == 2


def test_newest_source_modified_date_wins_not_download_order(local_store, sources):
    add_source(sources, 'yandex', 'newer.xlsx', workbook(local_store, qty=17), modified='2026-10-15T12:00:00Z')
    add_source(sources, 'mosreg', 'older.xlsx', workbook(local_store, qty=2), modified='2026-10-01T12:00:00Z')
    first = core.sync_and_scan_sources(SOURCES)
    assert first['added'] == 1
    assert first['state']['pending'][0]['items'][0]['qty'] == 17
    assert first['state']['pending'][0]['source_provider'] == 'yandex'
    # The healthy old source must never roll back newer data or curator edits.
    first['state']['pending'][0]['items'][0]['qty'] = 77
    core.save_state(first['state'])
    sources['yandex']['files'] = core.DiskSyncError('Таймаут')
    second = core.sync_and_scan_sources(SOURCES)
    assert second['added'] == 0 and second['partial']
    assert second['state']['pending'] == first['state']['pending']


def test_creation_date_used_only_when_modified_missing(local_store, sources):
    add_source(sources, 'yandex', 'a.xlsx', workbook(local_store, qty=17), modified=None, created='2026-10-15T12:00:00Z')
    add_source(sources, 'mosreg', 'b.xlsx', workbook(local_store, qty=2), modified='2026-10-01T12:00:00Z')
    result = core.sync_and_scan_sources(SOURCES)
    assert result['state']['pending'][0]['items'][0]['qty'] == 17


def test_unknown_remote_dates_do_not_override_timestamped_pending(local_store, sources):
    add_source(sources, 'yandex', 'a.xlsx', workbook(local_store, qty=17), modified='2026-10-15T12:00:00Z')
    add_source(sources, 'mosreg', 'b.xlsx', workbook(local_store, qty=2), modified=None)
    first = core.sync_and_scan_sources(SOURCES)
    sources['yandex']['files'] = core.DiskSyncError('Таймаут')
    second = core.sync_and_scan_sources(SOURCES)
    assert second['state']['pending'] == first['state']['pending']


def test_approved_and_excluded_workbooks_are_not_requeued_by_duplicate_filename(local_store, sources):
    file = local_store / 'nameless.xlsx'
    core.write_xlsx([['Наименование', 'Ед.', 'Количество'], ['Труба', 'шт', 3]], file)
    data = file.read_bytes()
    add_source(sources, 'yandex', 'nameless.xlsx', data)
    first = core.sync_and_scan_sources(SOURCES)
    assert first['state']['pending'][0]['rso'] == 'nameless'
    assert core.approve(0)
    add_source(sources, 'mosreg', 'other-name.xlsx', data, modified='2026-10-16T12:00:00Z')
    assert core.sync_and_scan_sources(SOURCES)['state']['pending'] == []
    core.delete_rso('nameless')
    assert core.sync_and_scan_sources(SOURCES)['state']['pending'] == []


def test_failure_of_new_version_does_not_rescan_old_cache(local_store, sources):
    old = workbook(local_store)
    item = add_source(sources, 'yandex', 'a.xlsx', old)
    first = core.sync_and_scan_sources(SOURCES)
    first['state']['pending'][0]['items'][0]['qty'] = 77
    core.save_state(first['state'])
    new = workbook(local_store, qty=12)
    item.update(md5=hashlib.md5(new).hexdigest(), size=len(new))
    sources['yandex']['bytes'][item['path']] = core.DiskSyncError('Файл недоступен')
    add_source(sources, 'mosreg', 'b.xlsx', workbook(local_store, 'Новое РСО', 5))
    result = core.sync_and_scan_sources(SOURCES)
    assert result['ok'] and result['partial'] and result['failed'] == 1
    assert state_by_rso(result)['Тестовое РСО'] == first['state']['pending'][0]


@pytest.mark.parametrize('legacy_provider', ['yandex', 'mosreg'])
def test_legacy_curator_edits_survive_newer_copy_in_second_source(local_store, sources, legacy_provider):
    other = 'mosreg' if legacy_provider == 'yandex' else 'yandex'
    data = workbook(local_store)
    add_source(sources, legacy_provider, 'old.xlsx', data, modified='2026-10-01T12:00:00Z')
    single = [s for s in SOURCES if s['provider'] == legacy_provider]
    state = core.sync_and_scan_sources(single)['state']
    pending = state['pending'][0]
    for field in (*core.FINGERPRINT_FIELDS, 'rso_explicit'):
        pending.pop(field, None)
    pending['items'][0].update(qty=77, via='override', cat='Ручная категория')
    core.save_state(state)
    # Real pre-upgrade manifests had no immutable content fingerprint.
    manifest = json.loads(core.FOLDER_A_MANIFEST_FILE.read_text())
    manifest.pop('sources', None)
    manifest.pop('version', None)
    for entry in manifest['files']:
        entry.pop('sha256', None)
    core.FOLDER_A_MANIFEST_FILE.write_text(json.dumps(manifest))
    add_source(sources, other, 'renamed.xlsx', data, modified='2026-10-15T12:00:00Z')
    result = core.sync_and_scan_sources(SOURCES)
    restored = result['state']['pending'][0]
    assert result['added'] == 0 and len(result['state']['pending']) == 1
    assert restored['items'] == pending['items']
    assert restored['loaded_at'] == pending['loaded_at']
    assert restored['source_provider'] == other
    assert restored['source_modified_at'] == '2026-10-15T12:00:00+00:00'


def test_equivalent_newer_inventory_advances_freshness_and_prevents_later_rollback(local_store, sources):
    data = workbook(local_store)
    item = add_source(sources, 'yandex', 'same.xlsx', data, modified='2026-10-01T12:00:00Z')
    first = core.sync_and_scan_sources(SOURCES)['state']
    first['pending'][0]['items'][0].update(qty=77, via='override')
    core.save_state(first)
    original_loaded_at = first['pending'][0]['loaded_at']
    path = local_store / 'newdate.xlsx'
    core.write_xlsx([
        ['Тестовое РСО'], ['15.10.2026'], ['Наименование', 'Ед.', 'Количество'], ['Труба', 'шт', 3],
    ], path)
    data = path.read_bytes()
    item.update(md5=hashlib.md5(data).hexdigest(), size=len(data), modified='2026-10-15T12:00:00Z')
    sources['yandex']['bytes'][item['path']] = data
    updated = core.sync_and_scan_sources(SOURCES)['state']['pending'][0]
    assert updated['items'] == first['pending'][0]['items']
    assert updated['date'] == '15.10.2026'
    assert updated['source_modified_at'] == '2026-10-15T12:00:00+00:00'
    assert updated['loaded_at'] == original_loaded_at
    add_source(sources, 'mosreg', 'intermediate.xlsx', workbook(local_store, qty=5), modified='2026-10-10T12:00:00Z')
    sources['yandex']['files'] = core.DiskSyncError('Недоступен')
    later = core.sync_and_scan_sources(SOURCES)
    assert later['added'] == 0 and later['state']['pending'][0] == updated


def test_migration_does_not_hash_unverified_changed_cache_as_old_source(local_store, sources):
    data = workbook(local_store)
    add_source(sources, 'yandex', 'old.xlsx', data)
    state = core.sync_and_scan_sources(SOURCES)['state']
    row = state['pending'][0]
    for field in core.FINGERPRINT_FIELDS:
        row.pop(field, None)
    row['items'][0]['qty'] = 77
    core.save_state(state)
    # A different valid workbook on disk must not become the approved old baseline.
    (core.INPUT_DIR / row['local_file']).write_bytes(workbook(local_store, qty=9))
    loaded = core.load_state()['pending'][0]
    assert not loaded.get('source_record_hash')
    assert loaded['items'][0]['qty'] == 77
