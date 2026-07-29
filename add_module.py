# -*- coding: utf-8 -*-
import sqlite3
import json

DB = r"data\dashboard.db"
USERNAME = "admin"
NEW_MODULE = "ecur"   # <-- при необходимости измени название модуля

conn = sqlite3.connect(DB)
cur = conn.cursor()

# Текущее состояние
cur.execute("SELECT id, username, modules FROM users WHERE username=?", (USERNAME,))
row = cur.fetchone()
if not row:
    print(f"[!!] Пользователь '{USERNAME}' не найден")
    raise SystemExit(1)

uid, uname, modules_str = row
print(f"ДО:  {uname} -> {modules_str}")

modules = json.loads(modules_str)

if NEW_MODULE in modules:
    print(f"[i] Модуль '{NEW_MODULE}' уже есть у пользователя. Ничего не делаем.")
else:
    modules.append(NEW_MODULE)
    new_str = json.dumps(modules, ensure_ascii=False)
    cur.execute("UPDATE users SET modules=? WHERE id=?", (new_str, uid))
    conn.commit()
    print(f"[✓] Добавлен модуль '{NEW_MODULE}'")

# Проверка
cur.execute("SELECT username, modules FROM users WHERE username=?", (USERNAME,))
uname, modules_str = cur.fetchone()
print(f"ПОСЛЕ: {uname} -> {modules_str}")

conn.close()