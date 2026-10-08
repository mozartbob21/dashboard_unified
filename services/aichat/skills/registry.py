"""Bounded Agent Skills discovery and lazy, contained text resource reads.

Only administrator-installed directories are considered. Discovery reads YAML
frontmatter, never the instruction body or references. No process is executed.
"""
from __future__ import annotations

import os
import re
import stat
from itertools import islice
from pathlib import Path, PurePosixPath

import yaml

SKILLS_DIR = Path(__file__).resolve().parents[3] / 'skills' / 'aichat'
MAX_PACKS = 64
MAX_FRONTMATTER_BYTES = 32 * 1024
MAX_INSTRUCTION_BYTES = 64 * 1024
MAX_RESOURCE_BYTES = 64 * 1024
MAX_METADATA_ENTRIES = 32
MAX_TOOL_NAMES = 32
NAME = re.compile(r'[a-z0-9]+(?:-[a-z0-9]+)*\Z')
TEXT_SUFFIXES = frozenset({'.md', '.txt', '.json'})


class SkillError(ValueError):
    """A safe error suitable for the chat UI; never includes host file paths."""


class _UniqueSafeLoader(yaml.SafeLoader):
    pass


def _unique_mapping(loader, node, deep=False):
    result = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        if not isinstance(key, str) or key in result:
            raise SkillError('Ключи YAML должны быть уникальными строками.')
        result[key] = loader.construct_object(value_node, deep=deep)
    return result


_UniqueSafeLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _unique_mapping)


def _valid_id(value):
    return isinstance(value, str) and 1 <= len(value) <= 64 and NAME.fullmatch(value) is not None


def _plain_text(value, maximum, *, required=False):
    if not isinstance(value, str) or len(value) > maximum:
        raise SkillError('Некорректный тип или длина поля YAML.')
    if any(ord(c) < 32 and c not in '\n\t' for c in value):
        raise SkillError('В YAML есть недопустимые управляющие символы.')
    if required and not value.strip():
        raise SkillError('Обязательное поле YAML пустое.')
    return value.strip()


def _root():
    try:
        root = SKILLS_DIR.resolve(strict=True)
        if not root.is_dir():
            raise SkillError('Каталог навыков недоступен.')
        return root
    except (OSError, RuntimeError):
        raise SkillError('Каталог навыков недоступен.') from None


def _pack(skill_id):
    if not _valid_id(skill_id):
        raise SkillError('Недопустимый идентификатор навыка.')
    root = _root()
    path = root / skill_id
    try:
        if path.is_symlink() or not path.is_dir() or path.resolve(strict=True).parent != root:
            raise SkillError('Навык недоступен или повреждён.')
        return path
    except (OSError, RuntimeError):
        raise SkillError('Навык недоступен или повреждён.') from None


def _open_text_file(path, pack):
    """Reject non-regular files, escaping paths and symlinks before opening."""
    try:
        relative = path.relative_to(pack)
        current = pack
        for part in relative.parts:
            current = current / part
            if current.is_symlink():
                raise SkillError('Символические ссылки в наборе навыка недоступны.')
        resolved = path.resolve(strict=True)
        resolved.relative_to(pack.resolve(strict=True))
        if not stat.S_ISREG(path.stat().st_mode):
            raise SkillError('Ресурс навыка должен быть обычным текстовым файлом.')
        flags = os.O_RDONLY | getattr(os, 'O_BINARY', 0) | getattr(os, 'O_NOFOLLOW', 0)
        # A final-component symlink swap is also rejected on hosts with O_NOFOLLOW.
        fd = os.open(path, flags)
        try:
            if not stat.S_ISREG(os.fstat(fd).st_mode):
                raise SkillError('Ресурс навыка должен быть обычным текстовым файлом.')
            return os.fdopen(fd, 'rb')
        except BaseException:
            os.close(fd)
            raise
    except SkillError:
        raise
    except (OSError, RuntimeError, ValueError):
        raise SkillError('Ресурс навыка недоступен или находится вне набора.') from None


def _frontmatter(stream):
    first = stream.readline(MAX_FRONTMATTER_BYTES + 1)
    consumed = len(first)
    if first.removeprefix(b'\xef\xbb\xbf').rstrip(b'\r\n') != b'---':
        raise SkillError('SKILL.md должен начинаться с YAML frontmatter.')
    lines = []
    while consumed <= MAX_FRONTMATTER_BYTES:
        line = stream.readline(MAX_FRONTMATTER_BYTES - consumed + 1)
        consumed += len(line)
        if not line or consumed > MAX_FRONTMATTER_BYTES:
            break
        if line.rstrip(b'\r\n') == b'---':
            try:
                text = b''.join(lines).decode('utf-8')
                # No aliases/anchors: their expansion can exceed the bounded input.
                if any(isinstance(token, (yaml.tokens.AliasToken, yaml.tokens.AnchorToken))
                       for token in yaml.scan(text)):
                    raise SkillError('YAML aliases и anchors в навыках не поддерживаются.')
                value = yaml.load(text, Loader=_UniqueSafeLoader)
            except SkillError:
                raise
            except (yaml.YAMLError, UnicodeError, RecursionError):
                raise SkillError('Некорректный YAML в SKILL.md.') from None
            if not isinstance(value, dict):
                raise SkillError('Frontmatter должен быть YAML-объектом.')
            return value
        lines.append(line)
    raise SkillError('Frontmatter не закрыт или превышает 32 КиБ.')


def _descriptor(pack, data):
    name = data.get('name')
    if not _valid_id(name) or name != pack.name:
        raise SkillError('Поле name должно совпадать с именем каталога навыка.')
    description = _plain_text(data.get('description'), 1024, required=True)
    for key, maximum in (('license', MAX_FRONTMATTER_BYTES), ('compatibility', 500)):
        if key in data:
            _plain_text(data[key], maximum, required=True)
    metadata = data.get('metadata', {})
    if not isinstance(metadata, dict) or len(metadata) > MAX_METADATA_ENTRIES:
        raise SkillError('Metadata должна быть небольшим словарём строк.')
    clean_metadata = {}
    for key, value in metadata.items():
        clean_key = _plain_text(key, 128, required=True)
        if clean_key in clean_metadata:
            raise SkillError('Ключи metadata должны быть уникальными.')
        clean_metadata[clean_key] = _plain_text(value, MAX_FRONTMATTER_BYTES)
    title = _plain_text(clean_metadata.get('neurona-title', 'Дополнительный навык'), 120, required=True)
    if not re.search('[А-Яа-яЁё]', title):
        title = 'Дополнительный навык'
    allowed = data.get('allowed-tools', '')
    allowed = _plain_text(allowed, 4096)
    tools = list(dict.fromkeys(allowed.split()))
    if len(tools) > MAX_TOOL_NAMES or any(len(tool) > 128 for tool in tools):
        raise SkillError('Слишком много имён инструментов в allowed-tools.')
    return {'id': name, 'name': name, 'title': title, 'description': description,
            'allowed_tools': tools, 'metadata': clean_metadata}


def _metadata(pack):
    with _open_text_file(pack / 'SKILL.md', pack) as stream:
        return _descriptor(pack, _frontmatter(stream))


def catalog_snapshot():
    """One bounded, uncached scan with safe diagnostics for invalid installed packs."""
    result = {'items': [], 'issues': []}
    try:
        root = _root()
        with os.scandir(root) as iterator:
            entries = list(islice(iterator, MAX_PACKS + 1))
    except (OSError, SkillError):
        result['issues'].append('Каталог навыков недоступен. Администратору нужно проверить установку.')
        return result
    if len(entries) > MAX_PACKS:
        result['issues'].append('Каталог содержит больше 64 записей. Часть наборов не загружена.')
    for entry in sorted(entries[:MAX_PACKS], key=lambda e: e.name):
        try:
            if entry.is_symlink():
                result['issues'].append('Пропущена символическая ссылка в каталоге навыков.')
                continue
            if not entry.is_dir(follow_symlinks=False):
                continue
            if not _valid_id(entry.name):
                result['issues'].append('Пропущен набор с недопустимым именем каталога.')
                continue
            result['items'].append(_metadata(_pack(entry.name)))
        except (OSError, SkillError, RecursionError):
            label = f' «{entry.name}»' if _valid_id(entry.name) else ''
            result['issues'].append(f'Набор{label} пропущен: проверьте SKILL.md и текстовые ресурсы.')
    return result


def discover():
    """Return metadata only. Newly installed folders are visible on the next call."""
    return catalog_snapshot()['items']


def load_instructions(skill_id):
    """Read a selected skill's body only when the chat activates that skill."""
    pack = _pack(skill_id)
    with _open_text_file(pack / 'SKILL.md', pack) as stream:
        _descriptor(pack, _frontmatter(stream))
        content = stream.read(MAX_INSTRUCTION_BYTES + 1)
    if len(content) > MAX_INSTRUCTION_BYTES:
        raise SkillError('Инструкции навыка превышают 64 КиБ.')
    try:
        text = content.decode('utf-8').strip()
        if '\0' in text:
            raise UnicodeError()
        return text
    except UnicodeError:
        raise SkillError('Инструкции навыка должны быть текстом UTF-8.') from None


def read_resource(skill_id, path):
    """Read only explicitly requested references/*.md|txt|json within this pack."""
    if not isinstance(path, str) or len(path) > 512 or '\\' in path or ':' in path:
        raise SkillError('Недопустимый путь ресурса навыка.')
    parts = path.split('/')
    if len(parts) < 2 or parts[0] != 'references' or any(p in ('', '.', '..') for p in parts):
        raise SkillError('Доступны только текстовые файлы внутри references/.')
    relative = PurePosixPath(path)
    if relative.is_absolute() or relative.suffix.lower() not in TEXT_SUFFIXES:
        raise SkillError('Разрешены только файлы .md, .txt и .json в references/.')
    pack = _pack(skill_id)
    _metadata(pack)  # Removed or newly invalid packs cannot keep serving old resources.
    with _open_text_file(pack.joinpath(*parts), pack) as stream:
        content = stream.read(MAX_RESOURCE_BYTES + 1)
    if len(content) > MAX_RESOURCE_BYTES:
        raise SkillError('Ресурс навыка превышает 64 КиБ.')
    try:
        text = content.decode('utf-8')
        if '\0' in text:
            raise UnicodeError()
        return text
    except UnicodeError:
        raise SkillError('Ресурс навыка должен быть текстом UTF-8.') from None
