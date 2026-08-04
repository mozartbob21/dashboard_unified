# add_water_dashboard.py
"""Добавляет модуль water-dashboard в права пользователей.
Сам находит БД, в которой есть таблица users."""
import sqlite3
import json
import os

NEW_MODULE = "water-dashboard"
TARGET_USERS = ["admin"]          # кому выдать доступ

CANDIDATES = ["data/dashboard.db", "dashboard.db"]


def find_db_with_users():
    """Проверяет кандидатов и возвращает путь к БД, где есть таблица users."""
    for path in CANDIDATES:
        if not os.path.exists(path):
            print(f"  [{path}] файл не найден")
            continue
        conn = sqlite3.connect(path)
        tables = [r[0] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'")]
        conn.close()
        print(f"  [{path}] таблицы: {tables}")
        if "users" in tables:
            return path
    return None


def main():
    print("Поиск базы данных...")
    db_path = find_db_with_users()

    if not db_path:
        print("\nНе найдена БД с таблицей users!")
        return

    print(f"\nИспользую БД: {db_path}\n")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row

    changed = 0
    for u in conn.execute("SELECT id, username, modules FROM users").fetchall():
        if u["username"] not in TARGET_USERS:
            continue

        mods = json.loads(u["modules"] or "[]")
        if NEW_MODULE not in mods:
            mods.append(NEW_MODULE)
            conn.execute(
                "UPDATE users SET modules = ? WHERE id = ?",
                (json.dumps(mods, ensure_ascii=False), u["id"]),
            )
            changed += 1
            print(f"+ {u['username']}: добавлен модуль {NEW_MODULE}")

    conn.commit()

    print("\n--- Проверка ---")
    for r in conn.execute("SELECT username, modules FROM users"):
        has = NEW_MODULE in json.loads(r["modules"] or "[]")
        mark = "✅" if has else "  "
        print(f"{mark} {r['username']:12} -> {r['modules']}")

    conn.close()
    print(f"\nГотово. Изменено пользователей: {changed}")


if __name__ == "__main__":
    main()
    