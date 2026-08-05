"""Хранилище системных уведомлений. Админ добавляет записи вручную в data/notifications.json."""
import json
import time
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"
FILE = DATA_DIR / "notifications.json"

DEFAULT = {
    "notifications": [
        {
            "id": 1,
            "icon": "🎉",
            "title": "Добро пожаловать в Нейрону!",
            "text": "Система уведомлений подключена. Администратор добавляет оповещения вручную в файл notifications.json.",
            "time": int(time.time()),
            "author": "Система",
        }
    ]
}


def list_all() -> list:
    if not FILE.exists():
        FILE.parent.mkdir(parents=True, exist_ok=True)
        FILE.write_text(json.dumps(DEFAULT, ensure_ascii=False, indent=2), encoding="utf-8")
        return DEFAULT["notifications"]
    try:
        data = json.loads(FILE.read_text(encoding="utf-8"))
        return list(data.get("notifications", []))
    except Exception:
        return []
        