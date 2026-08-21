"""Финальная проверка Qwen с правильной моделью."""
import os
import httpx
from dotenv import load_dotenv
load_dotenv()

BASE = os.getenv("QWEN_API_BASE", "https://aiplatform.mosreg.ru/api/user-models/v1").rstrip("/")
KEY = os.getenv("QWEN_API_KEY", "").strip()
MODEL = os.getenv("QWEN_MODEL", "qwen3.8-27b-fp8").strip()

print(f"=== Финальная проверка Qwen ===")
print(f"BASE: {BASE}")
print(f"MODEL: {MODEL}")
print(f"KEY: {KEY[:15]}...\n")

print(f"Тестовый чат с моделью {MODEL}:")
try:
    r = httpx.post(f"{BASE}/chat/completions",
                   headers={"Authorization": f"Bearer {KEY}",
                            "Content-Type": "application/json"},
                   json={"model": MODEL,
                         "messages": [{"role": "user", "content": "Скажи привет одним словом"}],
                         "temperature": 0.4, "max_tokens": 50, "stream": False},
                   timeout=60, verify=True)
    print(f"   status: {r.status_code}")
    if r.status_code == 200:
        data = r.json()
        answer = data["choices"][0]["message"]["content"]
        print(f"   ✅ ответ: {answer}")
        print("\n🎉 Qwen работает! Перезапускай сервер.")
    else:
        print(f"   ❌ {r.text[:500]}")
except Exception as e:
    print(f"   ❌ ошибка: {e}")