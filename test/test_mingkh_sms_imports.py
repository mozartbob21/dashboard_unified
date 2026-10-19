"""Synthetic archives only: never access the supplied user archive or portal."""
from datetime import date
import importlib.util
import io
import json
from pathlib import Path
import stat
import sys
import zipfile

import pytest
from openpyxl import Workbook

from services.mingkh.sms import context, rez_data, sys_data


@pytest.fixture
def importer(tmp_path, monkeypatch):
    source = Path(__file__).resolve().parents[1] / 'services' / 'mingkh' / 'sms' / 'imports.py'
    spec = importlib.util.spec_from_file_location('services.mingkh.sms.imports', source)
    module = importlib.util.module_from_spec(spec)
    monkeypatch.setitem(sys.modules, spec.name, module)
    spec.loader.exec_module(module)
    monkeypatch.setattr(context, 'DATA', tmp_path)
    monkeypatch.setattr(context, 'today', lambda: date(2026, 10, 19))
    monkeypatch.setattr(sys_data, 'today', lambda: date(2026, 10, 19))
    monkeypatch.setattr(rez_data, 'FOUND', str(tmp_path / 'rez_found.json'))
    return module


def archive(files):
    target = io.BytesIO()
    with zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED) as output:
        for name, value in files.items():
            output.writestr(name, value if isinstance(value, bytes) else json.dumps(value, ensure_ascii=False))
    return target.getvalue()


def registry(street='Город А, улица Новая'):
    workbook = Workbook()
    workbook.active.title = 'Sheet1'
    workbook.active.append(['Улица', 'ОМСУ', 'Статус'])
    workbook.active.append([street, 'Округ А', 'Жалобы продолжаются'])
    target = io.BytesIO()
    workbook.save(target)
    return target.getvalue()


def found(first='2026-01-01', last='2026-02-01'):
    return {'street-key': {'addr': 'Улица Новая, 1', 'omsu': 'Округ А', 'score': 5,
                           'first_from': first, 'first_to': first, 'last_from': last, 'last_to': last}}


def snapshot(tmp_path):
    return {str(path.relative_to(tmp_path)): path.read_bytes() for path in tmp_path.rglob('*')
            if path.is_file() and not path.name.startswith('.') and path.suffix != '.lock'}


def test_archive_import_and_repeat_without_duplicates(importer, tmp_path):
    payload = archive({'смс/data/omsu.json': [['1', 'Округ А']],
                       'смс/data/rez_found.json': found(),
                       'смс/data/sys_registry/2026-10-14.xlsx': registry(),
                       'смс/.env': b'SECRET=must-not-read',
                       'смс/server.py': b'raise Exception("must-not-execute")'})
    first = importer.import_archive(payload)
    assert first['stats']['changed_files'] == 3
    assert first['stats']['registries_imported'] == 1
    assert first['stats']['ignored_files'] == 2
    before = snapshot(tmp_path)
    second = importer.import_archive(payload)
    assert second['stats']['changed_files'] == 0
    assert second['stats']['registries_skipped'] == 1
    assert snapshot(tmp_path) == before
    assert not (tmp_path / '.env').exists() and not (tmp_path / 'server.py').exists()


def test_merge_keeps_earliest_latest_and_current_municipality_name(importer, tmp_path):
    context.atomic_json(tmp_path / 'omsu.json', [['1', 'Текущее название']])
    context.atomic_json(tmp_path / 'rez_found.json', found('2026-02-01', '2026-09-01'))
    result = importer.import_archive(archive({'data/omsu.json': [['1', 'Старое название'], ['2', 'Округ Б']],
                                               'data/rez_found.json': found('2026-01-01', '2026-05-01')}))
    assert dict(context.json_read(tmp_path / 'omsu.json')) == {'1': 'Текущее название', '2': 'Округ Б'}
    stored = context.json_read(tmp_path / 'rez_found.json')['street-key']
    assert stored['first_from'] == '2026-01-01'
    assert stored['last_to'] == '2026-09-01'
    assert result['stats']['municipalities_added'] == 1


def test_registry_conflict_keeps_every_original_file(importer, tmp_path):
    importer.import_archive(archive({'data/sys_registry/2026-10-14.xlsx': registry()}))
    before = snapshot(tmp_path)
    with pytest.raises(ValueError, match='другой реестр'):
        importer.import_archive(archive({'data/omsu.json': [['1', 'Округ А']],
                                        'data/rez_found.json': found(),
                                        'data/sys_registry/2026-10-14.xlsx': registry('Другая улица')}))
    assert snapshot(tmp_path) == before


@pytest.mark.parametrize('bad', [b'{broken', b'{"key":NaN}', b'{"key":{},"key":{}}'])
def test_corrupt_json_rejects_all_before_any_write(importer, tmp_path, bad):
    with pytest.raises(ValueError):
        importer.import_archive(archive({'data/omsu.json': [['1', 'Округ А']], 'data/rez_found.json': bad}))
    assert snapshot(tmp_path) == {}


def test_corrupt_xlsx_rejects_all_before_any_write(importer, tmp_path):
    with pytest.raises(ValueError):
        importer.import_archive(archive({'data/omsu.json': [['1', 'Округ А']], 'data/sys_registry/2026-10-14.xlsx': b'not excel'}))
    assert snapshot(tmp_path) == {}


@pytest.mark.parametrize('name', ['../data/omsu.json', '/data/omsu.json', 'C:/data/omsu.json',
                                 'root/../data/omsu.json', 'data\\omsu.json', 'root//data/omsu.json'])
def test_unsafe_paths_rejected(importer, tmp_path, name):
    with pytest.raises(ValueError):
        importer.import_archive(archive({name: [['1', 'Округ А']]}))
    assert snapshot(tmp_path) == {}


def test_symlink_archive_rejected(importer):
    target = io.BytesIO()
    with zipfile.ZipFile(target, 'w') as output:
        info = zipfile.ZipInfo('data/omsu.json')
        info.create_system = 3
        info.external_attr = (stat.S_IFLNK | 0o777) << 16
        output.writestr(info, '/etc/anything')
    with pytest.raises(ValueError, match='ссылки'):
        importer.import_archive(target.getvalue())


def test_zip_limits(importer, monkeypatch):
    payload = archive({'data/omsu.json': [['1', 'Округ А']], 'ignored.txt': b'x'})
    monkeypatch.setattr(importer, 'MAX_ARCHIVE_BYTES', len(payload) - 1)
    with pytest.raises(ValueError, match='20 МБ'):
        importer.import_archive(payload)
    monkeypatch.setattr(importer, 'MAX_ARCHIVE_BYTES', len(payload) + 1)
    monkeypatch.setattr(importer, 'MAX_MEMBERS', 1)
    with pytest.raises(ValueError, match='слишком много'):
        importer.import_archive(payload)


def test_excessive_compression_rejected(importer):
    with pytest.raises(ValueError, match='сжатый'):
        importer.import_archive(archive({'data/omsu.json': [['1', 'Округ А']], 'ignored.bin': b'0' * (2 * 1024 * 1024)}))


def test_different_roots_same_target_are_rejected(importer):
    with pytest.raises(ValueError, match='вариантов'):
        importer.import_archive(archive({'one/data/omsu.json': [], 'two/data/omsu.json': []}))


def test_only_allowlisted_files_are_read(importer, monkeypatch):
    payload = archive({'data/omsu.json': [['1', 'Округ А']], '.env': b'not-read',
                       'data/rez_rows/2026-10.json': b'not-read', 'sys_registry/2026-01-01.xlsx': b'not-read'})
    original = zipfile.ZipFile.open
    opened = []
    def monitored(self, name, *args, **kwargs):
        opened.append(name.filename if isinstance(name, zipfile.ZipInfo) else name)
        return original(self, name, *args, **kwargs)
    monkeypatch.setattr(zipfile.ZipFile, 'open', monitored)
    importer.import_archive(payload)
    assert opened == ['data/omsu.json']


def test_replacement_error_rolls_back_prior_files(importer, tmp_path, monkeypatch):
    context.atomic_json(tmp_path / 'omsu.json', [['1', 'Текущее название']])
    before = snapshot(tmp_path)
    original = importer.os.replace
    count = 0
    def fail_second(source, destination):
        nonlocal count
        count += 1
        if count == 2:
            raise OSError('Synthetic disk failure')
        return original(source, destination)
    monkeypatch.setattr(importer.os, 'replace', fail_second)
    with pytest.raises(OSError):
        importer.import_archive(archive({'data/omsu.json': [['2', 'Округ Б']], 'data/rez_found.json': found()}))
    assert snapshot(tmp_path) == before
    assert not list(tmp_path.glob('.sms-import-*'))


def test_active_collector_blocks_import_without_changes(importer, tmp_path):
    with context.process_lock(tmp_path / '.collector.lock', timeout=0):
        with pytest.raises(ValueError, match='обновляются'):
            importer.import_archive(archive({'data/omsu.json': [['1', 'Округ А']]}))
    assert snapshot(tmp_path) == {}


def test_storage_symlink_rejected(importer, tmp_path):
    other = tmp_path / 'other.json'
    other.write_text('[]')
    (tmp_path / 'omsu.json').symlink_to(other)
    with pytest.raises(ValueError, match='ссылку'):
        importer.import_archive(archive({'data/omsu.json': [['1', 'Округ А']]}))
    assert other.read_text() == '[]'


def test_encrypted_zip_flag_is_rejected(importer):
    payload = bytearray(archive({'data/omsu.json': [['1', 'Округ А']]}))
    local = payload.index(b'PK\x03\x04')
    central = payload.index(b'PK\x01\x02')
    payload[local + 6] |= 1
    payload[central + 8] |= 1
    with pytest.raises(ValueError, match='паролем'):
        importer.import_archive(bytes(payload))


def test_failed_rollback_preserves_backups_for_recovery(importer, tmp_path, monkeypatch):
    context.atomic_json(tmp_path / 'omsu.json', [['1', 'Текущее название']])
    original = importer.os.replace
    count = 0
    def fail_after_first(source, destination):
        nonlocal count
        count += 1
        if count >= 2:
            raise OSError('Synthetic unavailable disk')
        return original(source, destination)
    monkeypatch.setattr(importer.os, 'replace', fail_after_first)
    with pytest.raises(RuntimeError, match='Резервные копии сохранены'):
        importer.import_archive(archive({'data/omsu.json': [['2', 'Округ Б']], 'data/rez_found.json': found()}))
    backups = list(tmp_path.glob('.sms-import-*'))
    assert len(backups) == 1
    assert (backups[0] / '0.old').exists() and (backups[0] / 'manifest.json').exists()


def test_legacy_found_without_last_dates_is_idempotent(importer, tmp_path):
    value = found()
    value['street-key'].pop('last_from')
    value['street-key'].pop('last_to')
    payload = archive({'data/rez_found.json': value})
    first = importer.import_archive(payload)
    assert first['stats']['changed_files'] == 1
    before = snapshot(tmp_path)
    second = importer.import_archive(payload)
    assert second['stats']['changed_files'] == 0
    assert snapshot(tmp_path) == before
