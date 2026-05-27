import json
import uuid
from datetime import datetime
from pathlib import Path

from .analyzer import analyze_appeal


BASE_DIR = Path(__file__).resolve().parents[2]
APPEALS_DATA_DIR = BASE_DIR / "data" / "appeals"
APPEALS_FILE = APPEALS_DATA_DIR / "appeals.json"


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def load_appeals():
    if not APPEALS_FILE.exists():
        return []

    try:
        data = json.loads(APPEALS_FILE.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
        if isinstance(data, dict) and isinstance(data.get("items"), list):
            return data["items"]
    except Exception:
        return []

    return []


def save_appeals(items):
    APPEALS_DATA_DIR.mkdir(parents=True, exist_ok=True)
    APPEALS_FILE.write_text(
        json.dumps(items, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def get_status_label(status):
    mapping = {
        "awaiting_facts": "Ожидаются факты",
        "awaiting_review": "На проверке",
        "approved": "Утверждено",
        "rejected": "Отклонено",
        "sent": "Отправлено",
    }
    return mapping.get(status or "", status or "Неизвестно")


def enrich_appeal(item):
    item = dict(item or {})
    analysis = item.get("analysis_data") or {}

    item["status_label"] = get_status_label(item.get("status"))
    item["final_index"] = analysis.get("index_final", item.get("final_index"))
    item["priority_level"] = analysis.get("priority_level", item.get("priority_level", "planned"))
    item["priority_label"] = analysis.get("priority_label", item.get("priority_label", "Плановый"))
    item["criticality_label"] = analysis.get("criticality_label", "Не определена")
    item["emotion_label"] = analysis.get("emotion_label", "Не определена")

    try:
        item["analysis_data_pretty"] = json.dumps(analysis, ensure_ascii=False, indent=2)
    except Exception:
        item["analysis_data_pretty"] = "{}"

    return item


def list_appeals(status_filter=""):
    items = [enrich_appeal(item) for item in load_appeals()]
    items.sort(key=lambda x: x.get("created_at") or "", reverse=True)

    if status_filter:
        items = [item for item in items if item.get("status") == status_filter]

    return items


def get_appeal(request_id):
    for item in load_appeals():
        if item.get("request_id") == request_id:
            return enrich_appeal(item)
    return None


def add_history(item, action, payload=None):
    history = item.setdefault("history", [])
    history.append({
        "created_at": now_iso(),
        "action": action,
        "payload": payload or {},
    })


def create_appeal(subject, original_text, sender_email="manual@local"):
    request_id = str(uuid.uuid4())[:8]
    analysis_data = analyze_appeal(subject, original_text)

    item = {
        "request_id": request_id,
        "sender_email": sender_email or "manual@local",
        "subject": subject or "Обращение без темы",
        "original_text": original_text or "",
        "analysis_data": analysis_data,
        "final_index": analysis_data.get("index_final"),
        "priority_level": analysis_data.get("priority_level"),
        "priority_label": analysis_data.get("priority_label"),
        "facts_for_reply": "",
        "reply_template": "",
        "draft": "",
        "manual_reply": "",
        "status": "awaiting_facts",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "draft_created_at": "",
        "sent_at": "",
        "history": [],
    }

    add_history(item, "created", {
        "subject": item["subject"],
        "emotion": analysis_data.get("emotion_label"),
        "criticality": analysis_data.get("criticality_label"),
        "index": analysis_data.get("index_final"),
    })

    items = load_appeals()
    items.append(item)
    save_appeals(items)

    return enrich_appeal(item)


def update_appeal(request_id, **fields):
    items = load_appeals()
    updated = None

    for item in items:
        if item.get("request_id") != request_id:
            continue

        for key, value in fields.items():
            item[key] = value

        item["updated_at"] = now_iso()
        updated = item
        break

    if updated is not None:
        save_appeals(items)
        return enrich_appeal(updated)

    return None


def append_appeal_history(request_id, action, payload=None):
    items = load_appeals()

    for item in items:
        if item.get("request_id") == request_id:
            add_history(item, action, payload)
            item["updated_at"] = now_iso()
            save_appeals(items)
            return enrich_appeal(item)

    return None


def calculate_stats():
    all_items = load_appeals()

    return {
        "total": len(all_items),
        "awaiting_facts": len([x for x in all_items if x.get("status") == "awaiting_facts"]),
        "awaiting_review": len([x for x in all_items if x.get("status") == "awaiting_review"]),
        "approved": len([x for x in all_items if x.get("status") == "approved"]),
        "rejected": len([x for x in all_items if x.get("status") == "rejected"]),
        "sent": len([x for x in all_items if x.get("status") == "sent"]),
    }