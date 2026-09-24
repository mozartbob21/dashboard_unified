"""Organization-wide presentation assets, with non-destructive legacy migration."""
import hashlib
import io
import json
import threading
import uuid
from pathlib import Path

from PIL import Image, ImageOps
from pptx import Presentation
from . import workspace
from .documents import validate_office
from .workspace import ToolError, safe_name, MAX_FILE

LOCK = threading.RLock()


def _atomic(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name('.' + uuid.uuid4().hex + '.tmp')
    try:
        tmp.write_bytes(data)
        tmp.replace(path)
    finally:
        tmp.unlink(missing_ok=True)


def validate_template(data):
    validate_office(data)
    try:
        Presentation(io.BytesIO(data))
    except Exception:
        raise ToolError('Не удалось прочитать шаблон PowerPoint. Нужен корректный файл PPTX.') from None


def emblem_png(data):
    try:
        with Image.open(io.BytesIO(data)) as image:
            if image.format not in {'PNG', 'JPEG'} or image.width * image.height > 16000000:
                raise ValueError()
            output = io.BytesIO()
            ImageOps.exif_transpose(image).convert('RGBA').save(output, format='PNG')
            return output.getvalue()
    except (OSError, ValueError, Image.DecompressionBombError):
        raise ToolError('Нужен PNG или JPEG до 16 мегапикселей.') from None


def root():
    directory = workspace.ROOT.parent / 'branding'
    with LOCK:
        directory.mkdir(parents=True, exist_ok=True)
        marker = directory / 'migration.json'
        if marker.exists():
            return directory
        # Previous versions had both an organization folder and per-account folders.
        parents = [workspace.ROOT.parent]
        if workspace.ROOT.exists():
            parents += sorted(p for p in workspace.ROOT.iterdir() if p.is_dir() and not p.is_symlink())
        warnings, emblems = [], []
        for parent in parents:
            folder = parent / 'templates'
            if folder.is_dir() and not folder.is_symlink():
                for path in sorted(folder.iterdir()):
                    if not path.is_file() or path.is_symlink() or path.suffix.lower() != '.pptx':
                        continue
                    try:
                        safe_name(path.name, {'.pptx'})
                        if path.stat().st_size > MAX_FILE:
                            raise ToolError('Слишком большой шаблон.')
                        data = path.read_bytes()
                        validate_template(data)
                        target = directory / 'templates' / path.name
                        if target.exists() and target.read_bytes() != data:
                            digest = hashlib.sha256(data).hexdigest()[:10]
                            target = target.with_name(path.stem[:140] + '-' + digest + '.pptx')
                        if not target.exists():
                            _atomic(target, data)
                    except (ToolError, OSError):
                        warnings.append('Не перенесён шаблон «' + path.name + '». Исходный файл сохранён.')
            for filename in ('emblem.png', 'emblem.jpg', 'emblem.jpeg'):
                path = parent / filename
                if path.is_file() and not path.is_symlink():
                    try:
                        if path.stat().st_size > MAX_FILE:
                            raise ToolError('Слишком большой герб.')
                        data = emblem_png(path.read_bytes())
                        digest = hashlib.sha256(data).hexdigest()[:16]
                        _atomic(directory / 'previous-emblems' / (digest + '.png'), data)
                        emblems.append((parent == workspace.ROOT.parent, path.stat().st_mtime_ns, data))
                    except (ToolError, OSError):
                        warnings.append('Прежний герб не удалось перенести. Исходный файл сохранён.')
        if emblems and not (directory / 'emblem.png').exists():
            # Prefer the former organization asset, otherwise the most recent upload.
            chosen = max(emblems, key=lambda item: item[:2])
            _atomic(directory / 'emblem.png', chosen[2])
            if len({hashlib.sha256(item[2]).digest() for item in emblems}) > 1:
                warnings.append('Найдено несколько прежних гербов. Проверьте общий герб; все исходные варианты сохранены.')
        _atomic(marker, json.dumps({'version': 1, 'warnings': warnings}, ensure_ascii=False).encode())
        return directory


def options():
    with LOCK:
        directory = root()
        emblem = directory / 'emblem.png'
        return {'templates': sorted(p.name for p in (directory / 'templates').glob('*') if p.is_file() and p.suffix.lower() == '.pptx' and not p.is_symlink()),
                'emblem_url': '/tools/emblem?v=' + str(emblem.stat().st_mtime_ns) if emblem.is_file() else None,
                'branding_warnings': json.loads((directory / 'migration.json').read_text(encoding='utf-8'))['warnings']}


def save_template(name, data):
    name = safe_name(name, {'.pptx'})
    validate_template(data)
    with LOCK:
        _atomic(root() / 'templates' / name, data)
        return options()


def delete_template(name):
    with LOCK:
        (root() / 'templates' / safe_name(name, {'.pptx'})).unlink(missing_ok=True)
        return options()


def save_emblem(data):
    data = emblem_png(data)
    with LOCK:
        _atomic(root() / 'emblem.png', data)
        return options()


def capture(template):
    """Read a consistent asset snapshot before queuing a private conversion job."""
    with LOCK:
        directory = root()
        assets = {}
        if template:
            path = directory / 'templates' / safe_name(template, {'.pptx'})
            if not path.is_file() or path.is_symlink():
                raise ToolError('Общий шаблон не найден. Обновите список шаблонов.')
            assets['templates/' + template] = path.read_bytes()
        path = directory / 'emblem.png'
        if path.is_file() and not path.is_symlink():
            assets['emblem.png'] = path.read_bytes()
        return assets


def prepare_job(job, assets):
    directory = job / 'input' / 'branding'
    directory.mkdir(parents=True)
    for name, data in assets.items():
        path = directory / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
    return directory
