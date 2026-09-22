"""Private tool workspaces and bounded background jobs."""
import hashlib
import json
import logging
import re
import threading
import time
import shutil
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2] / 'data/tools/accounts'
POOL = ThreadPoolExecutor(max_workers=2, thread_name_prefix='tools')
SLOTS = threading.BoundedSemaphore(8)
MAX_FILE = 20 * 1024 * 1024
MAX_TOTAL = 60 * 1024 * 1024


class ToolError(ValueError):
    pass


def owner_key(user):
    identity = ('kc:' + user['kc_sub']) if user.get('kc_sub') else ('local:' + str(user['id']))
    return hashlib.sha256(identity.encode()).hexdigest()[:32]


def account_root(user):
    root = ROOT / owner_key(user)
    root.mkdir(parents=True, exist_ok=True)
    # Lazy retention: completed artifacts expire after seven days.
    for status in (root / 'jobs').glob('*/status.json'):
        try:
            if time.time() - status.stat().st_mtime > 7 * 86400:
                state = json.loads(status.read_text(encoding='utf-8')).get('state')
                if state in ('done', 'error') and not status.parent.is_symlink():
                    shutil.rmtree(status.parent, ignore_errors=True)
        except (OSError, ValueError):
            pass
    return root


def safe_name(name, suffix=None):
    name = str(name or '').strip()
    if not name or name in ('.', '..') or any(c in name for c in '/\\\x00:') or len(name) > 180:
        raise ToolError('Недопустимое имя файла.')
    if suffix and Path(name).suffix.lower() not in suffix:
        raise ToolError('Неподдерживаемый формат файла.')
    return name


async def read_upload(upload, suffixes):
    name = safe_name(getattr(upload, 'filename', ''), suffixes)
    payload = await upload.read(MAX_FILE + 1)
    if not payload or len(payload) > MAX_FILE:
        raise ToolError('Файл пуст или превышает 20 МБ.')
    return name, payload


def job_path(root, job_id):
    if not re.fullmatch(r'[a-f0-9]{32}', job_id):
        raise ToolError('Задание не найдено.')
    return root / 'jobs' / job_id


def _status(path, value):
    tmp = path / 'status.tmp'
    tmp.write_text(json.dumps(value, ensure_ascii=False), encoding='utf-8')
    tmp.replace(path / 'status.json')


def start_job(root, kind, operation):
    if not SLOTS.acquire(blocking=False):
        raise ToolError('Очередь заполнена. Повторите чуть позже.')
    job_id = uuid.uuid4().hex
    path = job_path(root, job_id)
    try:
        path.mkdir(parents=True)
        _status(path, {'state': 'queued', 'tool': kind})
    except Exception:
        SLOTS.release()
        raise

    def run():
        try:
            _status(path, {'state': 'running', 'tool': kind})
            result = operation(path) or {}
            _status(path, {'state': 'done', 'tool': kind, **result})
        except ToolError as exc:
            _status(path, {'state': 'error', 'tool': kind, 'message': str(exc)})
        except Exception as exc:
            logging.getLogger('tools').error('Tool %s failed (%s)', kind, type(exc).__name__)
            _status(path, {'state': 'error', 'tool': kind, 'message': 'Обработка не удалась. Проверьте формат и содержимое файла.'})
        finally:
            # Uploaded originals are temporary; retain only the requested result.
            shutil.rmtree(path / 'input', ignore_errors=True)
            shutil.rmtree(path / 'input_shots', ignore_errors=True)
            SLOTS.release()
    POOL.submit(run)
    return job_id
