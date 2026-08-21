"""Хранилище диалогов AI-чата."""
import json
import uuid
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data" / "aichat"
DIALOGS_FILE = DATA_DIR / "dialogs.json"

def now_iso():
    return datetime.now().isoformat(timespec="seconds")

def _load():
    if not DIALOGS_FILE.exists():
        return []
    try:
        data = json.loads(DIALOGS_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []

def _save(items):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    DIALOGS_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")

def list_dialogs():
    items = sorted(_load(), key=lambda d: d.get("updated_at") or "", reverse=True)
    out = []
    for d in items:
        msgs = d.get("messages") or []
        last = msgs[-1]["content"][:80] if msgs else "Нет сообщений"
        out.append({"id": d["id"], "title": d.get("title") or "Новый чат",
                    "updated_at": d.get("updated_at") or "", "preview": last,
                    "count": len(msgs)})
    return out

def get_dialog(did):
    for d in _load():
        if d.get("id") == did:
            return d
    return None

def create_dialog(title="Новый чат"):
    d = {"id": str(uuid.uuid4())[:8], "title": title or "Новый чат",
         "created_at": now_iso(), "updated_at": now_iso(), "messages": []}
    items = _load()
    items.append(d)
    _save(items)
    return d

def delete_dialog(did):
    items = _load()
    new = [d for d in items if d.get("id") != did]
    if len(new) == len(items):
        return False
    _save(new)
    return True

def append_message(did, role, content, file_name=None):
    items = _load()
    for d in items:
        if d.get("id") != did:
            continue
        d.setdefault("messages", []).append(
            {"role": role, "content": content, "ts": now_iso(), "file": file_name})
        d["updated_at"] = now_iso()
        if role == "user" and d.get("title") in ("Новый чат", "") and content:
            d["title"] = content[:60]
        _save(items)
        return d
    return None
