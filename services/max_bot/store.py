"""Two-step MAX account links and a small durable polling coordination store.

The platform account DB is read-only. No passwords, contacts, messages, searches,
bot token, or clear linking codes are stored here. MAX IDs are only routing identities.
"""
from __future__ import annotations

from contextlib import closing, contextmanager
import hashlib
import hmac
import json
import math
import re
import secrets
import sqlite3
import time
from pathlib import Path

from .access import read_account, denial


ALPHABET = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"
MIGRATION = Path(__file__).resolve().parents[2] / "migrations/20261022_max_bot.sql"
PERSISTENT_BINDINGS_MIGRATION = Path(__file__).resolve().parents[2] / "migrations/20261023_max_persistent_bindings.sql"
MAX_BINDING_TIMESTAMP = 253402300800  # First second of year 10000, outside supported dates.
# A far-future sentinel keeps old workers from deleting links during a rolling
# restart. New code treats this exact value as permanent, not as a deadline.
PERMANENT_BINDING_EXPIRY = 253402300799.0


class StoreError(Exception):
    def __init__(self, code, message, status_code=400):
        self.code, self.message, self.status_code = code, message, status_code
        super().__init__(message)


def _now(value=None):
    current = time.time() if value is None else float(value)
    if not math.isfinite(current):
        raise ValueError("Invalid clock")
    return current


def _sha(value):
    return hashlib.sha256(str(value).encode("utf-8")).hexdigest()


def _identifier(value, *, signed=False):
    if isinstance(value, bool) or not isinstance(value, (int, str)) or not re.fullmatch(r"-?[0-9]{1,19}", str(value)):
        raise StoreError("INVALID_ID", "Некорректный идентификатор.")
    number = int(value)
    if number == 0 or number > 2**63 - 1 or number < (-2**63 if signed else 1):
        raise StoreError("INVALID_ID", "Некорректный идентификатор.")
    return str(number)


def _code(value):
    if not isinstance(value, str) or len(value) > 32:
        raise StoreError("LINK_INVALID", "Код недействителен или истёк. Получите новый код в Нейроне.")
    normalized = re.sub(r"[ -]", "", value.upper()).translate(str.maketrans({"O": "0", "I": "1", "L": "1"}))
    if len(normalized) != 12 or any(c not in ALPHABET for c in normalized):
        raise StoreError("LINK_INVALID", "Код недействителен или истёк. Получите новый код в Нейроне.")
    return normalized


def _masked(value):
    return "•••• " + str(value)[-4:]


class Store:
    def __init__(self, settings):
        self.settings = settings
        self.db_path, self.accounts_path = Path(settings.db_path), Path(settings.accounts_path)
        self.link_ttl = int(getattr(settings, "link_ttl", 300))
        if not 60 <= self.link_ttl <= 600:
            raise ValueError("Invalid linking lifetime")
        if self.db_path.resolve() == self.accounts_path.resolve():
            raise ValueError("MAX bot storage must be separate from platform accounts")
        self.db_path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
        try:
            self.db_path.touch(mode=0o600, exist_ok=False)
        except FileExistsError:
            pass
        self.db_path.chmod(0o600)
        with self._connection() as conn:
            conn.executescript(MIGRATION.read_text(encoding="utf-8"))
            conn.executescript(PERSISTENT_BINDINGS_MIGRATION.read_text(encoding="utf-8"))
            conn.execute("INSERT OR IGNORE INTO max_bot_meta(key,value) VALUES ('rate_secret',?)", (secrets.token_hex(32),))
            self._rate_secret = bytes.fromhex(conn.execute("SELECT value FROM max_bot_meta WHERE key='rate_secret'").fetchone()[0])

    @contextmanager
    def _connection(self):
        try:
            with closing(sqlite3.connect(self.db_path, timeout=5)) as conn:
                conn.row_factory = sqlite3.Row
                try:
                    yield conn
                    conn.commit()
                except Exception:
                    conn.rollback()
                    raise
        except sqlite3.Error:
            raise StoreError("SYSTEM_UNAVAILABLE", "Система временно недоступна. Повторите позже.", 503) from None

    def _account(self, user_id):
        identifier = int(_identifier(user_id))
        try:
            user = read_account(self.accounts_path, identifier)
        except (OSError, sqlite3.Error):
            raise StoreError("SYSTEM_UNAVAILABLE", "Система временно недоступна. Повторите позже.", 503) from None
        rejected = denial(user)
        if rejected:
            raise StoreError(*rejected, status_code=403)
        user["is_active"] = bool(user["is_active"])
        return user

    @staticmethod
    def binding_is_active(expires_at, current):
        """Recognize the permanent sentinel; malformed expiry values fail closed."""
        return (type(expires_at) in (int, float) and math.isfinite(expires_at)
                and (expires_at == PERMANENT_BINDING_EXPIRY
                     or max(0, current) < expires_at < MAX_BINDING_TIMESTAMP))

    @staticmethod
    def _cleanup(conn, current):
        conn.execute("DELETE FROM max_bot_codes WHERE expires_at<=?", (current,))
        conn.execute("DELETE FROM max_bot_pending WHERE expires_at<=?", (current,))
        conn.execute("DELETE FROM max_bot_bindings WHERE expires_at>0 AND expires_at!=? AND expires_at<=?",
                     (PERMANENT_BINDING_EXPIRY, current))

    def create_link(self, user_id, now=None):
        current = _now(now)
        user = self._account(user_id)
        self.rate_limit("link-create", str(user["id"]), 5, 300, now=current)
        code = "".join(secrets.choice(ALPHABET) for _ in range(12))
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            self._cleanup(conn, current)
            conn.execute("DELETE FROM max_bot_codes WHERE user_id=?", (user["id"],))
            conn.execute("DELETE FROM max_bot_pending WHERE user_id=?", (user["id"],))
            conn.execute("INSERT INTO max_bot_codes VALUES (?,?,?,?,?,0)",
                         (_sha(code), user["id"], _sha(user["password_hash"]), current, current + self.link_ttl))
        return {"code": code, "expires_at": current + self.link_ttl}

    def redeem_link(self, code, max_user_id, chat_id, now=None):
        current = _now(now)
        max_id, chat = _identifier(max_user_id), _identifier(chat_id, signed=True)
        # Limit even syntactically invalid codes to prevent online guessing.
        self.rate_limit("link-redeem", max_id, 5, 300, now=current)
        self.rate_limit("link-chat", chat, 5, 300, now=current)
        code_hash = _sha(_code(code))
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT * FROM max_bot_codes WHERE code_hash=?", (code_hash,)).fetchone()
            if not row or row["consumed"] or row["expires_at"] <= current:
                raise StoreError("LINK_INVALID", "Код недействителен или истёк. Получите новый код в Нейроне.")
            user = self._account(row["user_id"])
            if not hmac.compare_digest(row["password_fingerprint"], _sha(user["password_hash"])):
                raise StoreError("LINK_INVALID", "Код недействителен или истёк. Получите новый код в Нейроне.")
            request_id = secrets.token_hex(16)
            conn.execute("UPDATE max_bot_codes SET consumed=1 WHERE code_hash=?", (code_hash,))
            conn.execute("INSERT INTO max_bot_pending VALUES (?,?,?,?,?,?,?)", (
                request_id, user["id"], max_id, chat, row["password_fingerprint"], current, current + self.link_ttl))
        return {"request_id": request_id, "status": "pending", "expires_at": current + self.link_ttl}

    def binding_status(self, user_id, now=None):
        current = _now(now)
        user = self._account(user_id)
        with self._connection() as conn:
            self._cleanup(conn, current)
            binding = conn.execute("SELECT * FROM max_bot_bindings WHERE user_id=?", (user["id"],)).fetchone()
            pending = conn.execute("SELECT * FROM max_bot_pending WHERE user_id=? ORDER BY created_at DESC", (user["id"],)).fetchall()
        fingerprint = _sha(user["password_hash"])
        active = bool(binding and self.binding_is_active(binding["expires_at"], current)
                      and hmac.compare_digest(binding["password_fingerprint"], fingerprint))
        result = {"active": active, "pending": [
            {"request_id": row["request_id"], "max_user_mask": _masked(row["max_user_id"]),
             "created_at": row["created_at"], "expires_at": row["expires_at"]}
            for row in pending if hmac.compare_digest(row["password_fingerprint"], fingerprint)
        ]}
        if active:
            permanent = binding["expires_at"] == PERMANENT_BINDING_EXPIRY
            result["binding"] = {"max_user_mask": _masked(binding["max_user_id"]),
                                 "expires_at": None if permanent else binding["expires_at"],
                                 "permanent": permanent}
        return result

    def confirm_link(self, user_id, request_id, now=None):
        current = _now(now)
        user = self._account(user_id)
        if not isinstance(request_id, str) or not re.fullmatch(r"[0-9a-f]{32}", request_id):
            raise StoreError("LINK_INVALID", "Запрос привязки не найден или истёк.")
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT * FROM max_bot_pending WHERE request_id=? AND user_id=?", (request_id, user["id"])).fetchone()
            if not row or row["expires_at"] <= current or not hmac.compare_digest(row["password_fingerprint"], _sha(user["password_hash"])):
                raise StoreError("LINK_INVALID", "Запрос привязки не найден или истёк.")
            # Expired bindings never block a new explicit confirmation.
            self._cleanup(conn, current)
            conflict = conn.execute("SELECT user_id FROM max_bot_bindings WHERE max_user_id=? AND user_id!=?", (row["max_user_id"], user["id"])).fetchone()
            if conflict:
                raise StoreError("LINK_CONFLICT", "Этот аккаунт MAX уже связан с другой учётной записью. Сначала отключите прежнюю привязку.", 409)
            latest = self._account(user["id"])
            if not hmac.compare_digest(row["password_fingerprint"], _sha(latest["password_hash"])):
                raise StoreError("LINK_INVALID", "Учётные данные изменились. Получите новый код.")
            conn.execute("INSERT INTO max_bot_bindings VALUES (?,?,?,?,?,?) ON CONFLICT(user_id) DO UPDATE SET "
                         "max_user_id=excluded.max_user_id,chat_id=excluded.chat_id,password_fingerprint=excluded.password_fingerprint,"
                         "created_at=excluded.created_at,expires_at=excluded.expires_at", (
                             user["id"], row["max_user_id"], row["chat_id"], row["password_fingerprint"], current,
                             PERMANENT_BINDING_EXPIRY))
            conn.execute("DELETE FROM max_bot_codes WHERE user_id=?", (user["id"],))
            conn.execute("DELETE FROM max_bot_pending WHERE user_id=?", (user["id"],))
        return {"active": True, "binding": {"max_user_mask": _masked(row["max_user_id"]),
                                           "expires_at": None, "permanent": True}}

    def unlink(self, user_id):
        # Local authenticated account identity is supplied by the existing router.
        # Unlink stays possible after grant revocation or password reset.
        identifier = int(_identifier(user_id))
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            for table in ("max_bot_codes", "max_bot_pending", "max_bot_bindings"):
                conn.execute(f"DELETE FROM {table} WHERE user_id=?", (identifier,))

    def unlink_max(self, max_user_id, chat_id):
        """Handle /stop using verified private-chat routing, even after local revocation."""
        max_id, chat = _identifier(max_user_id), _identifier(chat_id, signed=True)
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT user_id FROM max_bot_bindings WHERE max_user_id=? AND chat_id=?", (max_id, chat)).fetchone()
            if row:
                for table in ("max_bot_codes", "max_bot_pending", "max_bot_bindings"):
                    conn.execute(f"DELETE FROM {table} WHERE user_id=?", (row["user_id"],))
            conn.execute("DELETE FROM max_bot_pending WHERE max_user_id=? AND chat_id=?", (max_id, chat))

    def ensure_bot_id(self, bot_id):
        """Bind this security/cursor DB to GET /me identity, never to token text."""
        identity = _identifier(bot_id)
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT value FROM max_bot_meta WHERE key='bot_user_id'").fetchone()
            if row and row["value"] != identity:
                raise StoreError("BOT_ID_CHANGED", "Эта база привязок относится к другому боту MAX. Проверьте настройки; автоматическая замена запрещена.", 409)
            conn.execute("INSERT OR IGNORE INTO max_bot_meta(key,value) VALUES ('bot_user_id',?)", (identity,))

    def authorize(self, max_user_id, chat_id, now=None):
        current = _now(now)
        max_id, chat = _identifier(max_user_id), _identifier(chat_id, signed=True)
        with self._connection() as conn:
            row = conn.execute("SELECT * FROM max_bot_bindings WHERE max_user_id=? AND chat_id=?", (max_id, chat)).fetchone()
        if not row or not self.binding_is_active(row["expires_at"], current):
            raise StoreError("LINK_REQUIRED", "Для подключения аккаунта MAX нажмите «Подключить MAX» на главной странице Нейроны.", 401)
        user = self._account(row["user_id"])
        if not hmac.compare_digest(row["password_fingerprint"], _sha(user["password_hash"])):
            raise StoreError("LINK_REQUIRED", "Учётные данные изменились. Привяжите MAX заново в Нейроне.", 401)
        return {key: user[key] for key in ("id", "username", "role", "modules", "is_active")}

    def rate_limit(self, bucket, key, limit, window, now=None):
        if not re.fullmatch(r"[a-z0-9_-]{1,40}", bucket) or not 1 <= limit <= 10000 or not 1 <= window <= 86400:
            raise ValueError("Invalid rate limit")
        current = _now(now)
        digest = hmac.new(self._rate_secret, (bucket + ":" + str(key)).encode(), hashlib.sha256).hexdigest()
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute("DELETE FROM max_bot_rate_events WHERE occurred_at<=?", (current - 86400,))
            count = conn.execute("SELECT count(*) FROM max_bot_rate_events WHERE bucket=? AND key_hash=? AND occurred_at>?", (bucket, digest, current - window)).fetchone()[0]
            if count >= limit:
                raise StoreError("RATE_LIMITED", "Слишком много попыток. Повторите через несколько минут.", 429)
            conn.execute("INSERT INTO max_bot_rate_events(bucket,key_hash,occurred_at) VALUES (?,?,?)", (bucket, digest, current))

    def audit(self, event, outcome, now=None):
        if event not in {"link_create", "link_redeem", "link_confirm", "unlink", "search", "contacts", "poll", "send"}:
            raise ValueError("Invalid audit event")
        if outcome not in {"ok", "denied", "invalid", "rate_limited", "error", "uncertain", "unavailable"}:
            raise ValueError("Invalid audit outcome")
        current = _now(now)
        retention = max(1, min(90, int(getattr(self.settings, "audit_days", 30)))) * 86400
        try:
            with self._connection() as conn:
                conn.execute("BEGIN IMMEDIATE")
                conn.execute("DELETE FROM max_bot_audit WHERE occurred_at<=?", (current - retention,))
                conn.execute("INSERT INTO max_bot_audit(occurred_at,event,outcome) VALUES (?,?,?)", (current, event, outcome))
                conn.execute("DELETE FROM max_bot_audit WHERE id <= (SELECT id FROM max_bot_audit ORDER BY id DESC LIMIT 1 OFFSET 10000)")
            return True
        except StoreError:
            return False

    def lease(self, owner, ttl=60, now=None):
        if not isinstance(owner, str) or not re.fullmatch(r"[A-Za-z0-9_-]{1,128}", owner) or not 5 <= ttl <= 300:
            raise ValueError("Invalid lease")
        current = _now(now)
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT owner,expires_at FROM max_bot_lease WHERE id=1").fetchone()
            if row and row["expires_at"] > current and row["owner"] != owner:
                return False
            conn.execute("INSERT INTO max_bot_lease VALUES (1,?,?) ON CONFLICT(id) DO UPDATE SET owner=excluded.owner,expires_at=excluded.expires_at", (owner, current + ttl))
        return True

    def release_lease(self, owner):
        with self._connection() as conn:
            conn.execute("DELETE FROM max_bot_lease WHERE id=1 AND owner=?", (owner,))

    def get_marker(self):
        with self._connection() as conn:
            row = conn.execute("SELECT value FROM max_bot_meta WHERE key='poll_marker'").fetchone()
        return int(row[0]) if row else None

    def set_marker(self, marker, owner=None, now=None):
        if isinstance(marker, bool) or not isinstance(marker, (int, str)) or not re.fullmatch(r"[0-9]{1,20}", str(marker)):
            raise ValueError("Invalid poll marker")
        with self._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            if owner is not None:
                lease = conn.execute("SELECT 1 FROM max_bot_lease WHERE id=1 AND owner=? AND expires_at>?", (owner, _now(now))).fetchone()
                if not lease:
                    raise StoreError("LEASE_LOST", "Другой процесс уже обрабатывает события MAX.", 409)
            conn.execute("INSERT INTO max_bot_meta(key,value) VALUES ('poll_marker',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (str(int(marker)),))

    def claim_event(self, key, now=None):
        if not isinstance(key, str) or not 1 <= len(key) <= 1024:
            raise ValueError("Invalid event key")
        with self._connection() as conn:
            cursor = conn.execute("INSERT OR IGNORE INTO max_bot_events VALUES (?,'claimed',?,NULL)", (_sha(key), _now(now)))
            return cursor.rowcount == 1

    def finish_event(self, key, status, now=None):
        if status not in {"done", "ignored", "failed", "sent", "outbound_uncertain"}:
            raise ValueError("Invalid event status")
        with self._connection() as conn:
            conn.execute("UPDATE max_bot_events SET status=?,finished_at=? WHERE event_hash=? AND status='claimed'", (status, _now(now), _sha(key)))
