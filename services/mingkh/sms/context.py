"""Private runtime paths and bounded, atomic storage for the SMS workspace."""
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
import json
import os
from pathlib import Path
import tempfile
import time

DATA = Path(__file__).resolve().parents[3] / 'data' / 'mingkh' / 'sms'
MOSCOW = timezone(timedelta(hours=3))
MAX_STORE_BYTES = 128 * 1024 * 1024


def today():
    return datetime.now(MOSCOW).date()


def timestamp():
    return datetime.now(MOSCOW).isoformat(timespec='seconds')


def iso_date(value):
    try:
        if type(value) is date:
            return value
        if not isinstance(value, str) or len(value) != 10:
            raise ValueError
        parsed = date.fromisoformat(value)
        if parsed.isoformat() != value:
            raise ValueError
        return parsed
    except (TypeError, ValueError):
        raise ValueError('Укажите дату в формате ГГГГ-ММ-ДД.') from None


def date_range(start, end, *, max_days=1096):
    start, end = iso_date(start), iso_date(end)
    if start > end:
        raise ValueError('Дата «с» не может быть позже даты «по».')
    if start < date(2000, 1, 1) or end > today():
        raise ValueError('Выберите даты с 2000 года по сегодняшний день.')
    if (end - start).days + 1 > max_days:
        raise ValueError('Выбран слишком большой период. Сократите его.')
    return start, end


def atomic_json(path, value):
    """Write one complete private snapshot; keep the previous version on failure."""
    path = Path(path)
    payload = json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(',', ':')).encode('utf-8')
    if len(payload) > MAX_STORE_BYTES:
        raise ValueError('Слишком большой объём данных. Сократите период.')
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix='.' + path.name + '-', dir=path.parent)
    try:
        with os.fdopen(fd, 'wb') as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp, path)
    finally:
        try:
            os.unlink(temp)
        except FileNotFoundError:
            pass


def json_read(path, max_bytes=MAX_STORE_BYTES):
    with Path(path).open('rb') as stream:
        payload = stream.read(max_bytes + 1)
    if len(payload) > max_bytes:
        raise ValueError('Сохранённые данные слишком большие. Обратитесь к администратору.')
    try:
        return json.loads(payload, parse_constant=lambda _: (_ for _ in ()).throw(ValueError()))
    except (ValueError, UnicodeDecodeError):
        raise ValueError('Сохранённые данные повреждены. Повторите сбор или восстановите архив.') from None


@contextmanager
def process_lock(path, timeout=10):
    """Cross-process lock, compatible with Windows servers and Unix."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('a+b') as handle:
        handle.seek(0, os.SEEK_END)
        if not handle.tell():
            handle.write(b'0')
            handle.flush()
        deadline = time.monotonic() + timeout
        while True:
            handle.seek(0)
            try:
                if os.name == 'nt':
                    import msvcrt
                    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise ValueError('Данные сейчас обновляются. Повторите действие чуть позже.') from None
                time.sleep(0.05)
        try:
            yield
        finally:
            handle.seek(0)
            if os.name == 'nt':
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
