import json
from pathlib import Path


def save_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_json(path, default=None):
    path = Path(path)
    if not path.exists():
        return default
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def safe_int(value, default=0):
    try:
        if value is None:
            return default
        if isinstance(value, bool):
            return default
        if isinstance(value, (int, float)):
            return int(value)
        if isinstance(value, str):
            cleaned = (
                value.strip()
                .replace(" ", "")
                .replace("\u00A0", "")
                .replace(",", ".")
            )
            if not cleaned:
                return default
            return int(float(cleaned))
    except Exception:
        pass
    return default