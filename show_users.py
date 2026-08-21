"""Показывает зарегистрированных пользователей из dashboard.db."""
import sqlite3

c = sqlite3.connect("dashboard.db")
tables = [r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")]
print("Таблицы:", tables)

t = next((x for x in tables if "user" in x.lower()), None)
if not t:
    print("❌ таблица пользователей не найдена")
    exit(1)

print(f"\nСхема таблицы {t}:")
for r in c.execute(f"PRAGMA table_info({t})"):
    print("  ", r[1], "-", r[2])

print("\nПользователи:")
rows = list(c.execute(f"SELECT * FROM {t}"))
cols = [d[0] for d in c.description]
print(" | ".join(cols))
print("-" * 80)
for r in rows:
    # пароли/хеши не печатаем целиком
    out = []
    for col, val in zip(cols, r):
        s = str(val)
        if any(k in col.lower() for k in ("pass", "hash", "secret")) and len(s) > 8:
            s = s[:4] + "…"
        out.append(s)
    print(" | ".join(out))
print(f"\nВсего: {len(rows)}")