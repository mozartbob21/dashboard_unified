# -*- coding: utf-8 -*-
import sqlite3
from pathlib import Path

DB = r"data\dashboard.db"

if not Path(DB).exists():
    print(f"[!!] Файл {DB} не найден в текущей папке")
    raise SystemExit(1)

conn = sqlite3.connect(DB)
cur = conn.cursor()

print("=" * 70)
print("ВСЕ ТАБЛИЦЫ В БАЗЕ:")
print("=" * 70)
cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
tables = [r[0] for r in cur.fetchall()]
for t in tables:
    print(f"  • {t}")

print()
print("=" * 70)
print("СХЕМА КАЖДОЙ ТАБЛИЦЫ:")
print("=" * 70)
for t in tables:
    print(f"\n--- {t} ---")
    cur.execute(f"PRAGMA table_info({t});")
    for col in cur.fetchall():
        # (cid, name, type, notnull, default, pk)
        print(f"  {col[1]:<25} {col[2]:<15} {'PK' if col[5] else ''}")

# Пробуем найти таблицу пользователей
print()
print("=" * 70)
print("ПОИСК ПОЛЬЗОВАТЕЛЕЙ:")
print("=" * 70)

candidate_tables = [t for t in tables if 'user' in t.lower()]
for t in candidate_tables:
    print(f"\n>>> Содержимое таблицы '{t}':")
    try:
        cur.execute(f"SELECT * FROM {t} LIMIT 20;")
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]
        print("   Колонки:", cols)
        for row in rows:
            print("   ", row)
    except Exception as e:
        print(f"   Ошибка: {e}")

conn.close()