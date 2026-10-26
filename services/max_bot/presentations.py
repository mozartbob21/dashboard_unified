"""Opt-in weekly-document delivery using the existing private MAX binding.

Discovery and delivery are separate durable operations. A first complete listing
is a baseline, never a backlog mailing. An interrupted/ambiguous send is held for
manual investigation rather than replayed: MAX does not provide an idempotency
key. Importing or constructing the service performs no network operation.
"""
from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
import re
import secrets
import tempfile
import threading
import time

from .store import PERMANENT_BINDING_EXPIRY, StoreError, _identifier

MIGRATION = Path(__file__).resolve().parents[2] / "migrations/20261023_max_presentations.sql"
MAX_FILE_BYTES = 50 * 1024 * 1024
LEASE_SECONDS = 120
MAX_DELIVERIES_PER_TICK = 10
SOURCE_BLOCKING_ERRORS = {"SOURCE_AUTH_FAILED", "SOURCE_NOT_CONFIGURED", "SOURCE_LOGIN_CHANGED",
                          "SOURCE_CONNECTION_FAILED", "SOURCE_HTTP_FAILED"}


def _error_code(error, fallback):
    code = getattr(error, "code", "")
    return code if isinstance(code, str) and re.fullmatch(r"[A-Z][A-Z0-9_]{0,63}", code) else fallback


def _uid(user):
    return int(_identifier(user["id"] if isinstance(user, dict) else user))


class PresentationService:
    def __init__(self, store, *, source_factory=None, client=None, interval=3600, clock=time.time, waiter=None):
        if not 60 <= interval <= 86400:
            raise ValueError("Invalid presentation interval")
        self.store, self.source_factory, self.client = store, source_factory, client
        self.interval, self.clock = float(interval), clock
        self._tick_lock = threading.Lock()
        self._stop = threading.Event()
        self._wait = waiter or self._stop.wait
        self._thread = None
        with self.store._connection() as conn:
            conn.executescript(MIGRATION.read_text(encoding="utf-8"))

    def _require_account(self, user):
        current = self.store._account(_uid(user))
        if "presentations" not in current["modules"]:
            raise StoreError("MODULE_FORBIDDEN", "Нет доступа к еженедельным презентациям.", 403)
        return current

    def _binding(self, user_id):
        with self.store._connection() as conn:
            row = conn.execute("SELECT * FROM max_bot_bindings WHERE user_id=?", (user_id,)).fetchone()
        return dict(row) if row else None

    def set_subscription(self, user, max_id, chat_id, enabled):
        if type(enabled) is not bool:
            raise StoreError("INVALID_SUBSCRIPTION", "Некорректное состояние подписки.")
        user_id = _uid(user)
        max_id, chat_id = _identifier(max_id), _identifier(chat_id, signed=True)
        authorized = self.store.authorize(max_id, chat_id, now=self.clock())
        if authorized["id"] != user_id:
            raise StoreError("LINK_REQUIRED", "Привязка MAX не подтверждена.", 403)
        self._require_account(user_id)
        binding = self._binding(user_id)
        if not binding or binding["max_user_id"] != max_id or binding["chat_id"] != chat_id:
            raise StoreError("LINK_REQUIRED", "Привязка MAX не подтверждена.", 403)
        now = self.clock()
        with self.store._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            old = conn.execute("SELECT * FROM max_presentations_subscriptions WHERE user_id=?", (user_id,)).fetchone()
            same_binding = old and old["max_user_id"] == max_id and old["chat_id"] == chat_id and old["binding_created_at"] == binding["created_at"]
            # Repeated button presses never advance the subscription or enqueue old files.
            if old and same_binding and bool(old["enabled"]) == enabled:
                pass
            else:
                generation = (old["generation"] if old else 0) + 1
                conn.execute("INSERT INTO max_presentations_subscriptions VALUES (?,?,?,?,?,?,?,?,NULL) "
                    "ON CONFLICT(user_id) DO UPDATE SET max_user_id=excluded.max_user_id,chat_id=excluded.chat_id,"
                    "binding_created_at=excluded.binding_created_at,generation=excluded.generation,enabled=excluded.enabled,"
                    "activated_at=excluded.activated_at,updated_at=excluded.updated_at,last_error=NULL",
                    (user_id, max_id, chat_id, binding["created_at"], generation, int(enabled), now, now))
                conn.execute("UPDATE max_presentations_deliveries SET state='cancelled',updated_at=? "
                    "WHERE user_id=? AND state IN ('pending','retry','uploading')", (now, user_id))
        return self.status(user_id)

    def status(self, user):
        user_id = self._require_account(user)["id"]
        with self.store._connection() as conn:
            state = dict(conn.execute("SELECT * FROM max_presentations_state WHERE id=1").fetchone())
            sub = conn.execute("SELECT * FROM max_presentations_subscriptions WHERE user_id=?", (user_id,)).fetchone()
            counts = dict(conn.execute("SELECT state,count(*) FROM max_presentations_deliveries WHERE user_id=? "
                "GROUP BY state", (user_id,)).fetchall())
            issues = [dict(row) for row in conn.execute(
                "SELECT d.id,f.name,d.state,d.last_error AS code,d.generation FROM max_presentations_deliveries d "
                "JOIN max_presentations_files f ON f.attachment_id=d.attachment_id WHERE d.user_id=? "
                "AND d.state IN ('failed','uncertain') ORDER BY d.updated_at DESC,d.id DESC LIMIT 5", (user_id,))]
        active = bool(sub and sub["enabled"])
        if active:
            try:
                self._validate_subscription(dict(sub))
            except StoreError as exc:
                if exc.status_code >= 500:
                    raise
                active = False
        error = (sub["last_error"] if sub else None) or state["last_error"] or (issues[0]["code"] if issues else None)
        for issue in issues:
            issue["name"] = issue["name"][:200]
            issue["retryable"] = bool(active and sub and issue.pop("generation") == sub["generation"])
        return {"enabled": active, "subscribed": active, "issues": issues, "baseline_ready": state["baseline_at"] is not None,
                "last_checked": state["last_checked"], "last_new": state["last_new"],
                "last_error": error, "error_code": error,
                "pending": sum(counts.get(key, 0) for key in ("pending", "retry", "uploading", "sending")),
                "uncertain": counts.get("uncertain", 0), "failed": counts.get("failed", 0)}

    def retry_delivery(self, user, max_id, chat_id, delivery_id):
        """Explicit personal retry; caller must explain possible duplicate delivery.

        A successful/foreign/old-generation delivery can never be restarted.
        Repeated retries while already queued/in progress are idempotent.
        """
        user_id = self._require_account(user)["id"]
        max_id, chat_id = _identifier(max_id), _identifier(chat_id, signed=True)
        delivery_id = int(_identifier(delivery_id))
        authorized = self.store.authorize(max_id, chat_id, now=self.clock())
        if authorized["id"] != user_id:
            raise StoreError("LINK_REQUIRED", "Привязка MAX не подтверждена.", 403)
        with self.store._connection() as conn:
            sub = conn.execute("SELECT * FROM max_presentations_subscriptions WHERE user_id=?", (user_id,)).fetchone()
        if not sub or not sub["enabled"] or sub["max_user_id"] != max_id or sub["chat_id"] != chat_id:
            raise StoreError("SUBSCRIPTION_STOPPED", "Сначала включите личную подписку.", 409)
        self._validate_subscription(dict(sub))
        with self.store._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT d.state FROM max_presentations_deliveries d "
                "JOIN max_presentations_subscriptions s ON s.user_id=d.user_id AND s.generation=d.generation "
                "JOIN max_bot_bindings b ON b.user_id=s.user_id AND b.created_at=s.binding_created_at "
                "AND b.max_user_id=s.max_user_id AND b.chat_id=s.chat_id "
                "WHERE d.id=? AND d.user_id=? AND s.generation=? AND s.enabled=1 "
                "AND (b.expires_at=? OR b.expires_at>?)",
                (delivery_id, user_id, sub["generation"], PERMANENT_BINDING_EXPIRY, self.clock())).fetchone()
            if not row or row["state"] not in {"failed", "uncertain", "pending", "retry", "uploading", "sending"}:
                raise StoreError("DELIVERY_NOT_RETRYABLE", "Эту отправку нельзя повторить. Обновите состояние подписки.", 409)
            if row["state"] in {"failed", "uncertain"}:
                conn.execute("UPDATE max_presentations_deliveries SET state='retry',attempts=0,next_attempt=0,"
                    "last_error=NULL,updated_at=? WHERE id=? AND user_id=?", (self.clock(), delivery_id, user_id))
        return self.status(user_id)

    def _defer_if_stopped(self, delivery_id, owner):
        if not self._stop.is_set():
            return False
        self._record(delivery_id, "retry", owner=owner, code="MONITOR_STOPPED", retry_after=10)
        return True

    def _validate_subscription(self, sub):
        binding = self._binding(sub["user_id"])
        if not binding or binding["created_at"] != sub["binding_created_at"]:
            raise StoreError("LINK_REQUIRED", "Подключите MAX и включите подписку заново.", 403)
        user = self.store.authorize(sub["max_user_id"], sub["chat_id"], now=self.clock())
        if user["id"] != sub["user_id"] or "presentations" not in user["modules"]:
            raise StoreError("MODULE_FORBIDDEN", "Нет доступа к еженедельным презентациям.", 403)
        with self.store._connection() as conn:
            current = conn.execute("SELECT enabled,generation FROM max_presentations_subscriptions WHERE user_id=?", (sub["user_id"],)).fetchone()
        if not current or not current["enabled"] or current["generation"] != sub["generation"]:
            raise StoreError("SUBSCRIPTION_STOPPED", "Подписка отключена.", 403)
        return user

    def _disable(self, sub, code, owner):
        with self.store._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            self._owned_in_connection(conn, owner)
            conn.execute("UPDATE max_presentations_subscriptions SET enabled=0,last_error=?,updated_at=? "
                "WHERE user_id=? AND generation=?", (code, self.clock(), sub["user_id"], sub["generation"]))
            conn.execute("UPDATE max_presentations_deliveries SET state='cancelled',updated_at=? WHERE user_id=? "
                "AND generation=? AND state IN ('pending','retry','uploading')", (self.clock(), sub["user_id"], sub["generation"]))

    def _claim(self, owner):
        now = self.clock()
        with self.store._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            row = conn.execute("SELECT * FROM max_presentations_lease WHERE id=1").fetchone()
            if row and row["expires_at"] > now:
                return False
            conn.execute("INSERT INTO max_presentations_lease VALUES(1,?,?) ON CONFLICT(id) DO UPDATE SET "
                "owner=excluded.owner,expires_at=excluded.expires_at", (owner, now + LEASE_SECONDS))
            # Upload has no user-visible effect; an interrupted message may have arrived.
            conn.execute("UPDATE max_presentations_deliveries SET state='retry',updated_at=? "
                "WHERE state='uploading'", (now,))
            conn.execute("UPDATE max_presentations_deliveries SET state='uncertain',last_error='DELIVERY_INTERRUPTED',"
                "updated_at=? WHERE state='sending'", (now,))
        return True

    def _owned_in_connection(self, conn, owner):
        row = conn.execute("SELECT owner,expires_at FROM max_presentations_lease WHERE id=1").fetchone()
        if not row or row["owner"] != owner or row["expires_at"] <= self.clock():
            raise StoreError("MONITOR_LEASE_LOST", "Проверка продолжится при следующем запуске.", 503)

    def _owned(self, owner):
        with self.store._connection() as conn:
            self._owned_in_connection(conn, owner)

    def _heartbeat(self, owner, done):
        while not done.wait(20):
            try:
                with self.store._connection() as conn:
                    conn.execute("UPDATE max_presentations_lease SET expires_at=? WHERE id=1 AND owner=? AND expires_at>?",
                        (self.clock() + LEASE_SECONDS, owner, self.clock()))
            except StoreError:
                return

    @staticmethod
    def _metadata(item, now):
        values = [str(getattr(item, key, "") or "") for key in ("id", "name", "description", "created_at", "path")]
        if not values[0] or not values[1] or any(len(value) > limit for value, limit in zip(values, (128, 1024, 2000, 128, 4096))):
            raise ValueError("Invalid attachment metadata")
        return (*values, now)

    def _scan(self, owner):
        if self.source_factory is None:
            raise StoreError("SOURCE_NOT_CONFIGURED", "Источник презентаций ещё не настроен.", 503)
        with self.source_factory() as source:
            listing = source.list_attachments()
        now = self.clock()
        items = [self._metadata(item, now) for item in listing]
        if len(items) > 20000 or len({item[0] for item in items}) != len(items):
            raise ValueError("Invalid attachment listing")
        self._owned(owner)
        with self.store._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            self._owned_in_connection(conn, owner)
            state = conn.execute("SELECT baseline_at FROM max_presentations_state WHERE id=1").fetchone()
            baseline = state["baseline_at"] is None
            new_count = 0
            for item in items:
                inserted = conn.execute("INSERT OR IGNORE INTO max_presentations_files VALUES(?,?,?,?,?,?)", item).rowcount
                if not inserted or baseline:
                    continue
                new_count += 1
                conn.execute("INSERT OR IGNORE INTO max_presentations_deliveries "
                    "(attachment_id,user_id,generation,state,created_at,updated_at) "
                    "SELECT ?,user_id,generation,'pending',?,? FROM max_presentations_subscriptions WHERE enabled=1",
                    (item[0], now, now))
            conn.execute("UPDATE max_presentations_state SET baseline_at=COALESCE(baseline_at,?),last_checked=?,"
                "last_new=CASE WHEN ? > 0 THEN ? ELSE last_new END,next_scan=?,last_error=NULL,failures=0 WHERE id=1",
                (now, now, new_count, now, now + self.interval))
        return new_count

    def _record(self, delivery_id, state, *, owner, code=None, retry_after=0):
        with self.store._connection() as conn:
            return bool(conn.execute("UPDATE max_presentations_deliveries SET state=?,last_error=?,next_attempt=?,updated_at=? "
                "WHERE id=? AND state!='cancelled' AND EXISTS(SELECT 1 FROM max_presentations_lease "
                "WHERE id=1 AND owner=? AND expires_at>?)",
                (state, code, self.clock() + retry_after, self.clock(), delivery_id, owner, self.clock())).rowcount)

    def _begin_send(self, row, owner):
        with self.store._connection() as conn:
            conn.execute("BEGIN IMMEDIATE")
            return bool(conn.execute(
                "UPDATE max_presentations_deliveries SET state='sending',updated_at=? WHERE id=? AND state='uploading' "
                "AND EXISTS(SELECT 1 FROM max_presentations_subscriptions s WHERE s.user_id=? "
                "AND s.enabled=1 AND s.generation=?) AND EXISTS(SELECT 1 FROM max_presentations_lease "
                "WHERE id=1 AND owner=? AND expires_at>?)",
                (self.clock(), row["id"], row["user_id"], row["generation"], owner, self.clock())).rowcount)

    def _pause_source(self, code, owner):
        if code not in SOURCE_BLOCKING_ERRORS:
            return
        with self.store._connection() as conn:
            conn.execute("UPDATE max_presentations_state SET last_error=?,next_scan=?,failures=failures+1 WHERE id=1 "
                "AND EXISTS(SELECT 1 FROM max_presentations_lease WHERE id=1 AND owner=? AND expires_at>?)",
                (code, self.clock() + 300, owner, self.clock()))

    def _deliver(self, row, owner):
        sub = {key: row[key] for key in ("user_id", "max_user_id", "chat_id", "binding_created_at", "generation")}
        sending = False
        try:
            self._owned(owner)
            self._validate_subscription(sub)
            with self.store._connection() as conn:
                claimed = conn.execute("UPDATE max_presentations_deliveries SET state='uploading',attempts=attempts+1,"
                    "updated_at=? WHERE id=? AND state IN ('pending','retry') AND EXISTS(SELECT 1 FROM "
                    "max_presentations_lease WHERE id=1 AND owner=? AND expires_at>?)", (self.clock(), row["id"], owner, self.clock())).rowcount
            if not claimed:
                return False
            item = SimpleNamespace(id=row["attachment_id"], name=row["name"], description=row["description"],
                                   created_at=row["created_at_source"], path=row["source_path"])
            if self.client is None or self.source_factory is None:
                raise StoreError("SOURCE_NOT_CONFIGURED", "Источник презентаций ещё не настроен.", 503)
            with tempfile.TemporaryDirectory(prefix="max-presentation-", dir=self.store.db_path.parent) as folder:
                directory = Path(folder)
                directory.chmod(0o700)
                with self.source_factory() as source:
                    path = Path(source.download(item, directory))
                if path.is_symlink() or not path.resolve().is_relative_to(directory.resolve()) or not path.is_file():
                    raise ValueError("Invalid downloaded file")
                path.chmod(0o600)
                if path.stat().st_size > MAX_FILE_BYTES:
                    raise StoreError("PRESENTATION_TOO_LARGE", "Файл превышает допустимый размер отправки.", 413)
                self._owned(owner)
                self._validate_subscription(sub)
                if self._defer_if_stopped(row["id"], owner):
                    return False
                token = self.client.upload_file(path)
                caption = "Новая еженедельная презентация\n" + row["name"][:500]
                if row["created_at_source"]:
                    caption += "\nДата в источнике: " + row["created_at_source"][:128]
                if row["description"]:
                    caption += "\n" + row["description"][:1000]
                # MAX may need time to process an uploaded file. Keep the same
                # token through these definite rejections instead of re-uploading.
                for processing_attempt in range(6):
                    self._owned(owner)
                    self._validate_subscription(sub)
                    if self._defer_if_stopped(row["id"], owner):
                        return False
                    # Persist intent BEFORE contacting MAX; crash recovery must not replay it.
                    if not self._begin_send(row, owner):
                        return False
                    if self._defer_if_stopped(row["id"], owner):
                        return False
                    sending = True
                    try:
                        result = self.client.send_file(int(row["max_user_id"]), token, caption)
                    except Exception as exc:
                        if (getattr(exc, "code", None) != "MAX_ATTACHMENT_NOT_READY" or not getattr(exc, "retryable", False)
                                or getattr(exc, "uncertain", False) or processing_attempt == 5):
                            raise
                        sending = False
                        self._record(row["id"], "uploading", owner=owner, code="MAX_ATTACHMENT_NOT_READY")
                        if self._wait((2, 4, 8, 16, 30)[processing_attempt]):
                            self._record(row["id"], "retry", owner=owner, code="MONITOR_STOPPED", retry_after=10)
                            return False
                        continue
                    if result is not True:
                        raise RuntimeError("Delivery not confirmed")
                    return self._record(row["id"], "sent", owner=owner)
        except StoreError as exc:
            if exc.code == "MONITOR_LEASE_LOST":
                return False
            self._pause_source(exc.code, owner)
            if exc.status_code < 500 and exc.code != "PRESENTATION_TOO_LARGE":
                self._disable(sub, exc.code, owner)
                self._record(row["id"], "cancelled", owner=owner, code=exc.code)
            else:
                self._record(row["id"], "uncertain" if sending else "failed" if exc.code == "PRESENTATION_TOO_LARGE" else "retry",
                    owner=owner, code=exc.code, retry_after=300)
        except Exception as exc:
            code = _error_code(exc, "PRESENTATION_DELIVERY_FAILED")
            self._pause_source(code, owner)
            # Explicit API rejection is safe to retry; unknown errors during send are ambiguous.
            uncertain = bool(getattr(exc, "uncertain", sending if not hasattr(exc, "retryable") else False))
            retryable = bool(getattr(exc, "retryable", not sending)) and code != "SOURCE_TOO_LARGE"
            state = "uncertain" if sending and uncertain else "retry" if retryable and row["attempts"] < 7 else "failed"
            self._record(row["id"], state, owner=owner, code=code, retry_after=min(300, 5 * 2 ** min(row["attempts"], 6)))
        return False

    def tick(self):
        result = {"scanned": False, "sent": 0, "busy": False}
        if not self._tick_lock.acquire(blocking=False):
            return {**result, "busy": True}
        owner, done, heartbeat = secrets.token_hex(16), threading.Event(), None
        claimed = False
        try:
            claimed = self._claim(owner)
            if not claimed:
                return {**result, "busy": True}
            heartbeat = threading.Thread(target=self._heartbeat, args=(owner, done), daemon=True)
            heartbeat.start()
            with self.store._connection() as conn:
                subs = [dict(row) for row in conn.execute("SELECT * FROM max_presentations_subscriptions WHERE enabled=1")]
                state = dict(conn.execute("SELECT * FROM max_presentations_state WHERE id=1").fetchone())
            for sub in subs:
                try:
                    self._validate_subscription(sub)
                except StoreError as exc:
                    if exc.status_code >= 500:
                        raise
                    self._disable(sub, exc.code, owner)
            if state["next_scan"] <= self.clock():
                try:
                    result["new"] = self._scan(owner)
                    result["scanned"] = True
                except Exception as exc:
                    code = _error_code(exc, "PRESENTATION_SOURCE_FAILED")
                    with self.store._connection() as conn:
                        conn.execute("UPDATE max_presentations_state SET last_error=?,failures=failures+1,next_scan=? WHERE id=1 "
                            "AND EXISTS(SELECT 1 FROM max_presentations_lease WHERE id=1 AND owner=? AND expires_at>?)",
                            (code, self.clock() + min(300, 30 * 2 ** min(state["failures"], 4)), owner, self.clock()))
                    result["error_code"] = code
                    return result
            elif state["last_error"]:
                # A broken login/source must not trigger another login per queued file.
                return {**result, "error_code": state["last_error"]}
            with self.store._connection() as conn:
                deliveries = [dict(row) for row in conn.execute(
                    "SELECT d.*,s.max_user_id,s.chat_id,s.binding_created_at,f.name,f.description,"
                    "f.created_at AS created_at_source,f.source_path FROM max_presentations_deliveries d "
                    "JOIN max_presentations_subscriptions s ON s.user_id=d.user_id AND s.generation=d.generation "
                    "JOIN max_presentations_files f ON f.attachment_id=d.attachment_id "
                    "WHERE s.enabled=1 AND d.state IN ('pending','retry') AND d.next_attempt<=? ORDER BY d.id LIMIT ?",
                    (self.clock(), MAX_DELIVERIES_PER_TICK))]
            for row in deliveries:
                if self._stop.is_set():
                    break
                result["sent"] += int(self._deliver(row, owner))
                with self.store._connection() as conn:
                    source_error = conn.execute("SELECT last_error FROM max_presentations_state WHERE id=1").fetchone()[0]
                if source_error:
                    result["error_code"] = source_error
                    break
            return result
        except StoreError as exc:
            return {**result, "error_code": exc.code}
        finally:
            done.set()
            if heartbeat:
                heartbeat.join(timeout=1)
            try:
                if claimed:
                    with self.store._connection() as conn:
                        conn.execute("DELETE FROM max_presentations_lease WHERE id=1 AND owner=?", (owner,))
            finally:
                self._tick_lock.release()

    def start(self):
        """Use beside the existing poller; never start another /updates consumer."""
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        def run():
            while not self._stop.is_set():
                try:
                    self.tick()
                except Exception:
                    # Do not log source errors: upstream bodies may contain credentials.
                    pass
                self._stop.wait(10)
        self._thread = threading.Thread(target=run, name="max-presentations", daemon=True)
        self._thread.start()

    def close(self):
        self._stop.set()
        if self._thread and self._thread is not threading.current_thread():
            self._thread.join(timeout=2)
