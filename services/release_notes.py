"""Account-scoped release receipts with current, server-side module filtering."""
import json
import re
from datetime import datetime, timezone
from pathlib import Path

from fastapi import HTTPException

from core.roles import MODULE_NAMES, effective_modules
from services.release_catalog import RELEASES
from utils.db import get_db_connection

MIGRATION_ID = "20261015_release_notes"
MIGRATION_FILE = Path(__file__).resolve().parents[1] / "migrations" / (MIGRATION_ID + ".sql")
_ID = re.compile(r"[a-z0-9][a-z0-9._-]{0,99}\Z")


def _ensure_schema(conn):
    if not conn.execute("SELECT 1 FROM app_schema_migrations WHERE name=?", (MIGRATION_ID,)).fetchone():
        # The migration owns its short transaction. Call before a request's
        # snapshot/write transaction, including when concurrent workers start.
        conn.executescript(MIGRATION_FILE.read_text(encoding="utf-8"))


def initialize():
    """Optional eager initialization; API calls also migrate an existing DB."""
    with get_db_connection() as conn:
        _ensure_schema(conn)


def _principal(conn, user):
    if not isinstance(user, dict) or not user.get("is_active", True):
        raise HTTPException(401, "Необходимо войти в систему")
    if user.get("kc_sub"):
        subject = user["kc_sub"]
        if not isinstance(subject, str) or not subject or len(subject) > 255:
            raise HTTPException(401, "Необходимо войти в систему")
        # Existing Keycloak grants are determined from authenticated claims.
        return "keycloak:" + subject, set(effective_modules(user)), False
    if type(user.get("id")) is not int or user["id"] <= 0:
        raise HTTPException(401, "Необходимо войти в систему")
    row = conn.execute("SELECT id, modules, is_active FROM users WHERE id=?", (user["id"],)).fetchone()
    if row is None or not row["is_active"]:
        raise HTTPException(401, "Необходимо войти в систему")
    try:
        modules = json.loads(row["modules"] or "[]")
    except (ValueError, TypeError):
        modules = []
    if not isinstance(modules, list):
        modules = []
    # The token/DOM may still contain grants revoked in another administrator's
    # session. Read current local grants within the same DB transaction.
    allowed = {item for item in modules if isinstance(item, str) and item in MODULE_NAMES}
    manager = conn.execute(
        "SELECT 1 FROM account_control c WHERE c.id=1 AND (c.manager_user_id=? "
        "OR EXISTS (SELECT 1 FROM account_managers m WHERE m.user_id=?))",
        (row["id"], row["id"]),
    ).fetchone() is not None
    return f"local:{row['id']}", allowed, manager


def _validate_catalog():
    """Fail closed on a broken deployment catalog, before sending any entries."""
    ids = set()
    for release in RELEASES:
        release_id = release.get("id", "")
        if not isinstance(release_id, str) or not _ID.fullmatch(release_id) or release_id in ids:
            raise ValueError("Invalid or duplicate release id")
        ids.add(release_id)
        for field, limit in (("title", 160), ("version", 40), ("published_at", 50)):
            if not isinstance(release.get(field), str) or not 0 < len(release[field]) <= limit:
                raise ValueError("Invalid release metadata")
        published = datetime.fromisoformat(release["published_at"].replace("Z", "+00:00"))
        if published.tzinfo is None:
            raise ValueError("Release date requires a timezone")
        entries = release.get("entries")
        if not isinstance(entries, (list, tuple)) or not entries or len(entries) > 30:
            raise ValueError("Invalid release entries")
        entry_ids = set()
        for entry in entries:
            entry_id = entry.get("id", "")
            if not isinstance(entry_id, str) or not _ID.fullmatch(entry_id) or entry_id in entry_ids:
                raise ValueError("Invalid or duplicate release entry id")
            entry_ids.add(entry_id)
            audience = entry.get("audience")
            if audience not in {"all", "admin", "module"}:
                raise ValueError("Unknown release audience")
            if audience == "module" and entry.get("module_id") not in MODULE_NAMES:
                raise ValueError("Unknown release module")
            for field, limit in (("title", 160), ("text", 1500)):
                if not isinstance(entry.get(field), str) or not 0 < len(entry[field]) <= limit:
                    raise ValueError("Invalid release entry text")
    return RELEASES


def _visible(release, allowed, manager):
    entries = []
    for entry in release["entries"]:
        audience = entry["audience"]
        module_id = entry.get("module_id") if audience == "module" else None
        if not (audience == "all" or (audience == "admin" and manager)
                or (audience == "module" and module_id in allowed)):
            continue
        entries.append({"id": entry["id"], "module_id": module_id,
                        "module_name": MODULE_NAMES[module_id] if module_id else
                        ("Администрирование" if audience == "admin" else "Платформа"),
                        "title": entry["title"], "text": entry["text"]})
    if not entries:
        return None
    return {key: release[key] for key in ("id", "version", "title", "published_at")} | {"entries": entries}


def _list(user, *, include_read=False):
    catalog = _validate_catalog()
    with get_db_connection() as conn:
        _ensure_schema(conn)
        conn.execute("BEGIN")
        account_key, allowed, manager = _principal(conn, user)
        read_dates = {row[0]: row[1] for row in conn.execute(
            "SELECT release_id,read_at FROM release_note_receipts WHERE account_key=?", (account_key,))}
        releases = [{**visible, "read_at": read_dates.get(release["id"])} for release in catalog
                    if include_read or release["id"] not in read_dates
                    if (visible := _visible(release, allowed, manager)) is not None]
    releases.sort(key=lambda item: (datetime.fromisoformat(item["published_at"].replace("Z", "+00:00")), item["id"]), reverse=True)
    return {"releases": releases, "pending_count": sum(item["read_at"] is None for item in releases)}


def pending(user):
    return _list(user)


def history(user):
    return _list(user, include_read=True)


def acknowledge(user, release_id):
    catalog = _validate_catalog()
    with get_db_connection() as conn:
        _ensure_schema(conn)
        conn.execute("BEGIN IMMEDIATE")
        account_key, allowed, manager = _principal(conn, user)
        release = next((item for item in catalog if item["id"] == release_id), None)
        if release is None or _visible(release, allowed, manager) is None:
            raise HTTPException(404, "Обновление не найдено")
        conn.execute("INSERT OR IGNORE INTO release_note_receipts(account_key,release_id,read_at) VALUES (?,?,?)",
                     (account_key, release_id, datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")))
        row = conn.execute("SELECT read_at FROM release_note_receipts WHERE account_key=? AND release_id=?",
                           (account_key, release_id)).fetchone()
    return {"ok": True, "release_id": release_id, "read_at": row["read_at"]}
