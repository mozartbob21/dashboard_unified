"""Хранилище отчётов сумматора + правила утверждения из BPMN:
- минимум 2 утверждения (из ролей) -> approved
- отклонение ТОЛЬКО с комментарием -> revision
"""
import json
import time
import uuid
from pathlib import Path

DATA = Path(__file__).resolve().parents[2] / "data" / "summarizer"
FILE = DATA / "reports.json"
MIN_APPROVALS = 2


def _load():
    if FILE.exists():
        try:
            return json.loads(FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []


def _save(items):
    DATA.mkdir(parents=True, exist_ok=True)
    FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")


def create_report(text, result, user):
    items = _load()
    item = {
        "id": uuid.uuid4().hex[:8],
        "created_at": time.strftime("%d.%m.%Y %H:%M"),
        "author": user,
        "source": text[:20000],
        "result": result,
        "approvals": {},
        "rejects": [],
        "status": "pending",
        "revision_comment": "",
    }
    items.insert(0, item)
    _save(items)
    return item


def get_report(rid):
    return next((it for it in _load() if it["id"] == rid), None)


def list_reports(limit=20):
    return _load()[:limit]


def approve(rid, user):
    items = _load()
    for it in items:
        if it["id"] == rid:
            if it["status"] != "pending":
                return it
            it["approvals"][user] = True
            if len(it["approvals"]) >= MIN_APPROVALS:
                it["status"] = "approved"
            _save(items)
            return it
    return None


def reject(rid, user, comment):
    items = _load()
    for it in items:
        if it["id"] == rid:
            it["rejects"].append({"user": user, "comment": comment,
                                  "at": time.strftime("%d.%m.%Y %H:%M")})
            it["status"] = "revision"
            it["revision_comment"] = comment
            _save(items)
            return it
    return None


def to_pending(rid):
    items = _load()
    for it in items:
        if it["id"] == rid:
            it["status"] = "pending"
            it["approvals"] = {}
            _save(items)
            return it
    return None
