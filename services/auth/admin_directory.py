"""Paged local account directory. Presence contains only the latest server timestamp per account."""
import json
import time
from datetime import datetime, timezone

from fastapi import HTTPException, Request

from services.auth.accounts import is_parent_manager
from services.auth.security import ADMIN_ROLES
from utils.db import get_db_connection

PRESENCE_TTL = 90


def can_inspect(user):
    if not user or user.get("kc_sub") or type(user.get("id")) is not int:
        return False
    with get_db_connection() as conn:
        current = conn.execute("""SELECT u.role, u.id=c.manager_user_id AS parent,
            m.user_id IS NOT NULL AS delegated FROM users u CROSS JOIN account_control c
            LEFT JOIN account_managers m ON m.user_id=u.id
            WHERE c.id=1 AND u.id=? AND u.username=? AND u.is_active=1""",
            (user["id"], user["username"])).fetchone()
    return bool(current and (current["parent"] or
                (current["delegated"] and current["role"].strip().casefold() in ADMIN_ROLES)))


def require_inspector(request: Request):
    user = getattr(request.state, "user", None)
    if not can_inspect(user):
        raise HTTPException(403, "Доступ только для администратора админ-панели")
    return user


def touch(conn, user_id, timestamp):
    # No IDs, timestamps or status flags supplied by the client.
    conn.execute("""INSERT INTO account_presence(user_id,last_seen)
        SELECT id,? FROM users WHERE id=? AND is_active=1
            AND NOT EXISTS (SELECT 1 FROM account_archives WHERE user_id=users.id)
        ON CONFLICT(user_id) DO UPDATE SET last_seen=excluded.last_seen
        WHERE excluded.last_seen >= account_presence.last_seen + 5""", (int(timestamp), user_id))


def heartbeat(user):
    if not user:
        raise HTTPException(401, "Требуется авторизация")
    if not user.get("kc_sub") and type(user.get("id")) is int:
        with get_db_connection() as conn:
            touch(conn, user["id"], time.time())


def directory(user, *, q, archive, page, size, inspector, now=None, exact_login=None, presence="all"):
    timestamp = int(time.time() if now is None else now)
    active_sql = "(u.is_active=1 AND a.user_id IS NULL AND p.last_seen IS NOT NULL AND p.last_seen BETWEEN :cutoff AND :now)"
    joined = """FROM users u CROSS JOIN account_control c
        LEFT JOIN account_archives a ON a.user_id=u.id
        LEFT JOIN account_presence p ON p.user_id=u.id
        LEFT JOIN account_managers m ON m.user_id=u.id"""
    filters = ["c.id=1"]
    if archive != "all":
        filters.append("a.user_id IS NULL" if archive == "current" else "a.user_id IS NOT NULL")
    params = {"now": timestamp, "cutoff": timestamp - PRESENCE_TTL}
    if exact_login is not None:
        filters.append("u.username=:exact_login")
        params["exact_login"] = exact_login
    if presence in ("online", "offline"):
        filters.append(active_sql if presence == "online" else "NOT " + active_sql)
    terms = q.casefold().replace("ё", "е").split()
    order = "fold(u.username),u.id"
    if terms:
        fields = ["u.username", "u.email"]
        if inspector:
            fields.append("u.role")
        for index, term in enumerate(dict.fromkeys(terms)):
            key = f"term_{index}"
            filters.append("(" + " OR ".join(f"instr(fold({field}), :{key})>0" for field in fields) + ")")
            params[key] = term
        params["query"] = " ".join(terms)
        order = "CASE WHEN fold(u.username)=:query THEN 0 WHEN instr(fold(u.username), :query)=1 THEN 1 ELSE 2 END," + order
    where = " WHERE " + " AND ".join(filters)
    with get_db_connection() as conn:
        conn.create_function("fold", 1, lambda text: (text or "").casefold().replace("ё", "е"))
        conn.execute("BEGIN")  # Keep counters and this page in one snapshot.
        totals = conn.execute(f"SELECT count(*), count(CASE WHEN {active_sql} THEN 1 END) {joined} WHERE c.id=1", params).fetchone()
        matched = conn.execute("SELECT count(*) " + joined + where, params).fetchone()[0]
        pages = max(1, (matched + size - 1) // size)
        page = min(page, pages)
        params.update(size=size, offset=(page - 1) * size)
        rows = conn.execute(f"""SELECT u.id,u.username,u.email,u.role,u.modules,u.is_active,u.created_at,
            a.archived_at, u.id=c.manager_user_id AS is_manager,
            (u.id=c.manager_user_id OR m.user_id IS NOT NULL) AS can_manage_users,
            {active_sql} AS online, s.last_seen_at,
            NULLIF(MAX(COALESCE(p.last_seen,0),COALESCE(CAST(strftime('%s',s.last_seen_at) AS INTEGER),0)),0) AS presence_seen
            {joined} LEFT JOIN user_activity s ON s.user_id=u.id AND s.module_id='__platform__'
            {where} ORDER BY {order} LIMIT :size OFFSET :offset""", params).fetchall()
    users = []
    for row in rows:
        item = dict(row)
        for key in ("is_active", "is_manager", "can_manage_users"):
            item[key] = bool(item[key])
        online = bool(item.pop("online"))
        item["online"] = online
        last_seen = item.pop("presence_seen")
        item["last_online_at"] = datetime.fromtimestamp(last_seen,timezone.utc).isoformat() if last_seen and 0 < last_seen <= timestamp else None
        item["modules"] = json.loads(item["modules"])
        seen = item.pop("last_seen_at")
        item["activity"] = {"last_seen_at": seen} if seen else None
        users.append(item)
    result = {"users": users, "total": totals[0], "matched": matched, "page": page, "pages": pages,
              "current_user_id": user["id"], "current_username": user["username"], "is_parent_manager": is_parent_manager(user)}
    result.update(online=totals[1], offline=totals[0] - totals[1], presence_ttl=PRESENCE_TTL)
    return result
