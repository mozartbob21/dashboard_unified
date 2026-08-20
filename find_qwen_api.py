"""Ищет OpenAI-совместимый эндпоинт Qwen в сети компании."""
import httpx

# ▼▼▼ ВПИШИ сюда адрес из адресной строки страницы Qwen Studio (без пути) ▼▼▼
HOSTS = [
    "http://localhost:8000",
    "http://localhost:8080",
    # например: "http://10.0.0.15:8000",
    # например: "http://qwen.mosreg.local",
]

PATHS = ["/v1", "/api/v1", "/openai/v1", ""]

found = False
for host in HOSTS:
    for path in PATHS:
        base = host.rstrip("/") + path
        try:
            r = httpx.get(base + "/models", timeout=5, verify=False)
            if r.status_code == 200 and "data" in r.json():
                ids = [m["id"] for m in r.json()["data"]]
                print(f"✅ НАШЁЛ: {base}")
                print("   модели:", ids)
                found = True
        except Exception:
            pass
if not found:
    print("❌ не нашлось — добавь хост из адресной строки в список HOSTS")