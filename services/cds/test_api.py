"""
Тест 2-уровневой авторизации: NTLM (IIS) + Basic (1С) для OData.
"""
import base64

import requests
from requests_ntlm import HttpNtlmAuth

from services.cds.config import (
    CDS_HTTP_USER,
    CDS_HTTP_PASSWORD,
    CDS_1C_USER,
    CDS_1C_PASSWORD,
)

BASE = "http://176.100.216.181:63871/CDS"
ODATA = f"{BASE}/odata/standard.odata/"


def basic_header(user: str, password: str) -> str:
    token = base64.b64encode(f"{user}:{password}".encode("utf-8")).decode()
    return f"Basic {token}"


print("=" * 60)
print("ТЕСТ A: Смотрим заголовки 401 на OData (кто его выдаёт)")
print("=" * 60)
r = requests.get(ODATA, auth=HttpNtlmAuth(CDS_HTTP_USER, CDS_HTTP_PASSWORD), timeout=30)
print(f"Статус: {r.status_code}")
for k, v in r.headers.items():
    print(f"  {k}: {v}")
print(f"Тело (300 симв): {r.text[:300]}")

print()
print("=" * 60)
print("ТЕСТ B: NTLM + затем Basic 1С на том же соединении")
print("=" * 60)
session = requests.Session()
session.auth = HttpNtlmAuth(CDS_HTTP_USER, CDS_HTTP_PASSWORD)

# 1. Прогреваем соединение NTLM-ом
r1 = session.get(f"{BASE}/ru/", timeout=30)
print(f"Прогрев NTLM: {r1.status_code}")

# 2. На том же соединении шлём Basic для 1С
session.auth = None  # выключаем NTLM, чтобы не перетирал заголовок
r2 = session.get(
    ODATA + "$metadata",
    headers={"Authorization": basic_header(CDS_1C_USER, CDS_1C_PASSWORD)},
    timeout=30,
)
print(f"OData с Basic 1С: {r2.status_code}")
if r2.status_code == 200:
    print("🎉🎉🎉 OData ДОСТУПЕН! Полный API без браузера!")
    print(r2.text[:1000])
else:
    print(f"Тело: {r2.text[:300]}")

print()
print("=" * 60)
print("ТЕСТ C: NTLM с credentials от 1С (вдруг они доменные)")
print("=" * 60)
r3 = requests.get(
    ODATA + "$metadata",
    auth=HttpNtlmAuth(CDS_1C_USER, CDS_1C_PASSWORD),
    timeout=30,
)
print(f"Статус: {r3.status_code}")
if r3.status_code == 200:
    print("🎉 Сработало с 1С-credentials через NTLM!")
    print(r3.text[:500])