"""Allowlisted, transactional import of the standalone SMS workspace archive.

Only reference data is read. Code, credentials, uploaded filenames and cached
portal responses are never deployed or executed from an archive.
"""
from contextlib import ExitStack
import io
import json
import os
from pathlib import Path
import re
import stat
import shutil
import tempfile
import zipfile

from . import context, rez_data, sys_data

MAX_ARCHIVE_BYTES = 20 * 1024 * 1024
MAX_MEMBERS = 4000
MAX_EXPANDED_BYTES = 256 * 1024 * 1024
MAX_MEMBER_BYTES = 64 * 1024 * 1024
MAX_COMPRESSION_RATIO = 200
MAX_OMSU_BYTES = 2 * 1024 * 1024
MAX_FOUND_BYTES = 64 * 1024 * 1024


def _target(info):
    name = info.filename
    if (not name or len(name) > 1024 or '\\' in name or ':' in name
            or any(ord(char) < 32 for char in name)):
        raise ValueError('В архиве обнаружено недопустимое имя файла.')
    parts = name.rstrip('/').split('/')
    if name.startswith('/') or any(part in ('', '.', '..') for part in parts) or len(parts) > 40:
        raise ValueError('Архив содержит небезопасный путь к файлу.')
    mode = info.external_attr >> 16
    if stat.S_ISLNK(mode) or (stat.S_IFMT(mode) not in (0, stat.S_IFREG, stat.S_IFDIR)):
        raise ValueError('Архив не должен содержать ссылки или специальные файлы.')
    if info.flag_bits & 1:
        raise ValueError('Архивы с паролем не поддерживаются.')
    if info.is_dir():
        return None
    if tuple(parts[-2:]) in (('data', 'omsu.json'), ('data', 'rez_found.json')):
        return parts[-1]
    if len(parts) >= 3 and parts[-3:-1] == ['data', 'sys_registry']:
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}\.xlsx', parts[-1]):
            raise ValueError('Реестр системных адресов должен иметь имя ГГГГ-ММ-ДД.xlsx.')
        sys_data._report_date(parts[-1][:-5])
        return 'sys_registry/' + parts[-1]
    return None


def _json_bytes(value, limit):
    payload = json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(',', ':')).encode('utf-8')
    if len(payload) > min(limit, context.MAX_STORE_BYTES):
        raise ValueError('После объединения получилось слишком много данных.')
    return payload


def _read_json(payload, limit):
    if len(payload) > min(limit, context.MAX_STORE_BYTES):
        raise ValueError('Файл JSON в архиве слишком большой.')
    def pairs(values):
        result = {}
        for key, value in values:
            if key in result:
                raise ValueError('В JSON повторяются ключи записей.')
            result[key] = value
        return result
    try:
        return json.loads(payload, object_pairs_hook=pairs,
                          parse_constant=lambda _: (_ for _ in ()).throw(ValueError()))
    except (ValueError, UnicodeDecodeError, RecursionError):
        raise ValueError('В архиве повреждён файл JSON. Импорт отменён.') from None


def _omsu(value):
    if not isinstance(value, list) or len(value) > 1000:
        raise ValueError('В архиве некорректный справочник ОМСУ.')
    result = {}
    for row in value:
        if not isinstance(row, (list, tuple)) or len(row) != 2:
            raise ValueError('В справочнике ОМСУ нужны пары «идентификатор, название».')
        ident, title = str(row[0]), row[1]
        if (not re.fullmatch(r'-1|[0-9]{1,10}', ident) or not isinstance(title, str)
                or not title.strip() or len(title) > 300 or any(ord(char) < 32 for char in title)):
            raise ValueError('В справочнике ОМСУ обнаружена некорректная запись.')
        if ident == '-1':
            continue
        if ident in result:
            raise ValueError('В справочнике ОМСУ повторяется идентификатор.')
        result[ident] = title.strip()
    return result


def _prepare(payload):
    if not isinstance(payload, bytes) or not payload or len(payload) > MAX_ARCHIVE_BYTES:
        raise ValueError('Загрузите ZIP-архив размером не более 20 МБ.')
    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            entries = archive.infolist()
            if len(entries) > MAX_MEMBERS or sum(info.file_size for info in entries) > MAX_EXPANDED_BYTES:
                raise ValueError('Архив содержит слишком много файлов или данных.')
            seen_names, selected, targets = set(), [], set()
            for info in entries:
                target = _target(info)
                name = info.filename.casefold()
                if name in seen_names:
                    raise ValueError('В архиве повторяются имена файлов.')
                seen_names.add(name)
                if (info.file_size > MAX_MEMBER_BYTES
                        or info.file_size > max(1024 * 1024, info.compress_size * MAX_COMPRESSION_RATIO)):
                    raise ValueError('Архив содержит слишком большой или чрезмерно сжатый файл.')
                if target:
                    if target in targets:
                        raise ValueError('В архиве несколько вариантов одного справочника или реестра.')
                    targets.add(target)
                    selected.append((info, target))
            if not selected:
                raise ValueError('В архиве нет поддерживаемых данных: data/omsu.json, data/rez_found.json или data/sys_registry/ГГГГ-ММ-ДД.xlsx.')
            prepared = {}
            for info, target in selected:
                limit = (MAX_OMSU_BYTES if target == 'omsu.json' else
                         MAX_FOUND_BYTES if target == 'rez_found.json' else sys_data.MAX_REGISTRY_BYTES)
                if info.file_size > limit:
                    raise ValueError('Файл справочника или реестра превышает допустимый размер.')
                with archive.open(info) as stream:
                    content = stream.read(limit + 1)
                if len(content) > limit:
                    raise ValueError('Файл справочника или реестра превышает допустимый размер.')
                if target == 'omsu.json':
                    prepared[target] = _omsu(_read_json(content, limit))
                elif target == 'rez_found.json':
                    prepared[target] = rez_data.validate_found_data(_read_json(content, limit))
                    for record in prepared[target].values():
                        record.setdefault('last_from', record['first_from'])
                        record.setdefault('last_to', record['first_to'])
                else:
                    prepared[target] = sys_data.prepare_registry(content)
            return prepared, sum(not info.is_dir() for info in entries) - len(selected)
    except (zipfile.BadZipFile, RuntimeError, NotImplementedError, OSError, EOFError):
        raise ValueError('Не удалось прочитать ZIP-архив. Проверьте его целостность и формат.') from None


def _safe_destination(root, relative):
    path = root / relative
    current = path
    while current != root:
        if current.is_symlink():
            raise ValueError('Хранилище содержит ссылку вместо файла. Импорт отменён.')
        current = current.parent
    if root.is_symlink():
        raise ValueError('Хранилище содержит ссылку вместо папки. Импорт отменён.')
    return path


def _write_staged(path, data):
    with path.open('wb') as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def _commit(root, changes):
    """Rollback completed replacements if any later replacement fails."""
    if not changes:
        return
    stage = Path(tempfile.mkdtemp(prefix='.sms-import-', dir=root))
    keep_backups = False
    try:
        planned = []
        for index, (relative, content) in enumerate(changes.items()):
            destination = _safe_destination(root, relative)
            destination.parent.mkdir(parents=True, exist_ok=True)
            staged = stage / f'{index}.new'
            backup = stage / f'{index}.old'
            _write_staged(staged, content)
            if destination.exists():
                if not destination.is_file() or destination.stat().st_size > context.MAX_STORE_BYTES:
                    raise ValueError('Существующий файл хранилища не может быть заменён.')
                _write_staged(backup, destination.read_bytes())
            planned.append((destination, staged, backup))
        # Keep the manifest with backups if recovery itself encounters a disk error.
        _write_staged(stage / 'manifest.json', _json_bytes(
            [{'target': str(target.relative_to(root)), 'backup': backup.name,
              'existed': backup.exists()} for target, _staged, backup in planned], MAX_OMSU_BYTES))
        completed = []
        try:
            for destination, staged, backup in planned:
                os.replace(staged, destination)
                completed.append((destination, backup))
        except BaseException:
            recovery_errors = []
            for destination, backup in reversed(completed):
                try:
                    if backup.exists():
                        os.replace(backup, destination)
                    else:
                        destination.unlink(missing_ok=True)
                except OSError as exc:
                    recovery_errors.append(exc)
            if recovery_errors:
                keep_backups = True
                raise RuntimeError('Импорт прерван ошибкой диска. Резервные копии сохранены; администратору нужно восстановить хранилище.') from None
            raise
    finally:
        if not keep_backups:
            shutil.rmtree(stage, ignore_errors=True)


def import_archive(payload):
    """Import allowlisted data atomically; the caller authenticates an admin."""
    prepared, ignored = _prepare(payload)
    root = Path(context.DATA)
    if root.is_symlink():
        raise ValueError('Хранилище содержит ссылку вместо папки. Импорт отменён.')
    root.mkdir(parents=True, exist_ok=True)
    stats = {'municipalities_added': 0, 'municipalities_preserved': 0,
             'resonant_addresses': 0, 'registries_imported': 0, 'registries_skipped': 0,
             'ignored_files': ignored, 'changed_files': 0}
    with context.process_lock(root / '.collector.lock', timeout=0), ExitStack() as stack:
        if 'omsu.json' in prepared:
            stack.enter_context(context.process_lock(root / '.omsu.lock', timeout=0))
        if 'rez_found.json' in prepared:
            stack.enter_context(rez_data._found_lock)
            stack.enter_context(context.process_lock(root / 'rez_found.lock', timeout=0))
        changes = {}
        for relative, incoming in prepared.items():
            destination = _safe_destination(root, relative)
            if relative == 'omsu.json':
                current = _omsu(context.json_read(destination, max_bytes=MAX_OMSU_BYTES)) if destination.exists() else {}
                stats['municipalities_preserved'] = len(current.keys() & incoming.keys())
                combined = {**incoming, **current}  # Current platform names take precedence.
                if len(combined) > 1000:
                    raise ValueError('После объединения справочник ОМСУ слишком большой.')
                stats['municipalities_added'] = len(combined) - len(current)
                if stats['municipalities_added'] or not destination.exists():
                    changes[relative] = _json_bytes(sorted(combined.items(), key=lambda row: row[1]), MAX_OMSU_BYTES)
            elif relative == 'rez_found.json':
                current = rez_data.validate_found_data(context.json_read(destination, max_bytes=MAX_FOUND_BYTES)) if destination.exists() else {}
                combined = rez_data.merge_found_data(current, incoming)
                stats['resonant_addresses'] = len(combined)
                if combined != current or not destination.exists():
                    changes[relative] = _json_bytes(combined, MAX_FOUND_BYTES)
            else:
                if destination.exists():
                    if destination.stat().st_size > sys_data.MAX_REGISTRY_BYTES:
                        raise ValueError('Существующий реестр слишком большой. Импорт отменён.')
                    if sys_data.registry_fingerprint(destination.read_bytes()) != sys_data.registry_fingerprint(incoming):
                        raise ValueError('На дату %s уже сохранён другой реестр. Архив не импортирован.' % destination.stem)
                    stats['registries_skipped'] += 1
                else:
                    changes[relative] = incoming
                    stats['registries_imported'] += 1
        stats['changed_files'] = len(changes)
        _commit(root, changes)
    message = ('Данные архива импортированы. Текущие справочники и история сохранены.'
               if stats['changed_files'] else 'Все данные этого архива уже сохранены. Дубликаты не добавлены.')
    return {'ok': True, 'message': message, 'stats': stats}
