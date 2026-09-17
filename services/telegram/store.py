"""Encrypted Telegram sessions, outside the publicly served data directory."""
import json
import os
import sqlite3
import tempfile
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from cryptography.fernet import Fernet


class Busy(Exception):
    pass


class Store:
    def __init__(self, root=None, key=None):
        self.root = Path(root or os.getenv('NEURONA_TELEGRAM_DIR') or
                         Path(__file__).resolve().parents[2] / '.private' / 'telegram')
        project = Path(__file__).resolve().parents[2]
        if any(self.root.resolve().is_relative_to(project / name) for name in ('data', 'static', 'generated')):
            raise RuntimeError('Telegram-сессии нельзя хранить в веб-каталогах data, static или generated.')
        self.root.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.path = self.root / 'sessions.sqlite'
        key_path = self.root / 'session.key'
        key = key or os.getenv('TELEGRAM_SESSION_KEY')
        if not key:
            if not key_path.exists():
                if self.path.exists():
                    raise RuntimeError('Ключ Telegram-сессий отсутствует. Восстановите session.key из резервной копии.')
                with tempfile.NamedTemporaryFile(dir=self.root, delete=False) as stream:
                    temp_key = Path(stream.name)
                    stream.write(Fernet.generate_key())
                try:
                    os.link(temp_key, key_path)
                except FileExistsError:
                    pass
                finally:
                    temp_key.unlink()
            key = key_path.read_bytes()
        self.cipher = Fernet(key.encode() if isinstance(key, str) else key)
        with self.connect() as db:
            db.executescript('''
                CREATE TABLE IF NOT EXISTS accounts(owner TEXT PRIMARY KEY, payload BLOB NOT NULL);
                CREATE TABLE IF NOT EXISTS leases(owner TEXT PRIMARY KEY, token TEXT NOT NULL, expires REAL NOT NULL);
            ''')
        os.chmod(self.path, 0o600)

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.path, timeout=5)
        try:
            with db:
                yield db
        finally:
            db.close()

    def get(self, owner):
        with self.connect() as db:
            row = db.execute('SELECT payload FROM accounts WHERE owner=?', (owner,)).fetchone()
        return json.loads(self.cipher.decrypt(row[0])) if row else {}

    def put(self, owner, payload):
        encrypted = self.cipher.encrypt(json.dumps(payload, ensure_ascii=False).encode())
        with self.connect() as db:
            db.execute('INSERT INTO accounts VALUES (?,?) ON CONFLICT(owner) DO UPDATE SET payload=excluded.payload',
                       (owner, encrypted))

    def delete(self, owner):
        with self.connect() as db:
            db.execute('DELETE FROM accounts WHERE owner=?', (owner,))

    @contextmanager
    def claim(self, owner):
        token = uuid.uuid4().hex
        with self.connect() as db:
            db.execute('DELETE FROM leases WHERE expires<?', (time.time(),))
            try:
                db.execute('INSERT INTO leases VALUES (?,?,?)', (owner, token, time.time() + 120))
            except sqlite3.IntegrityError:
                raise Busy() from None
        try:
            yield
        finally:
            with self.connect() as db:
                db.execute('DELETE FROM leases WHERE owner=? AND token=?', (owner, token))
