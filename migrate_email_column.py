# migrate_email_column.py
import sqlite3

DB = "data/dashboard.db"

conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row

# Проверяем, есть ли колонка email
cols = [r[1] for r in conn.execute("PRAGMA table_info(users)")]
print(f"Колонки в users: {cols}")

if "email" not in cols:
    conn.execute("ALTER TABLE users ADD COLUMN email TEXT")
    print("✅ Добавлена колонка email в таблицу users")
else:
    print("✅ Колонка email уже существует")

# Создаём таблицы для регистрации, если их нет
conn.executescript("""
CREATE TABLE IF NOT EXISTS pending_registrations (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    username      TEXT NOT NULL,
    email         TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    code_hash     TEXT NOT NULL,
    attempts      INTEGER NOT NULL DEFAULT 0,
    created_at    TEXT NOT NULL,
    expires_at    TEXT NOT NULL,
    last_send_at  TEXT
);
CREATE TABLE IF NOT EXISTS user_settings (
    user_id    INTEGER PRIMARY KEY,
    data       TEXT NOT NULL DEFAULT '{}',
    updated_at TEXT NOT NULL
);
""")
print("✅ Таблицы pending_registrations и user_settings готовы")

conn.commit()
conn.close()
print("\nМиграция завершена.")