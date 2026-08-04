# set_role.py
"""Простая смена роли: python set_role.py <логин> [роль]"""
import sqlite3
import sys

DB = "data/dashboard.db"

ROLES = ["Администратор", "Руководитель", "Пользователь",
         "Контроль данных", "УТНКР", "Камеры", "Эмпатичные ответы"]


def main():
    if len(sys.argv) < 2:
        print("Использование: python set_role.py <логин> [роль]")
        print("Доступные роли:", ", ".join(ROLES))
        return

    username = sys.argv[1]
    role = sys.argv[2] if len(sys.argv) > 2 else "Руководитель"

    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    row = conn.execute(
        "SELECT id, username, role FROM users WHERE username = ?", (username,)
    ).fetchone()

    if not row:
        print(f"❌ Пользователь '{username}' не найден.")
        conn.close()
        return

    conn.execute("UPDATE users SET role = ? WHERE id = ?", (role, row["id"]))
    conn.commit()
    print(f"✅ {username}: роль '{row['role']}' → '{role}'")
    conn.close()


if __name__ == "__main__":
    main()
    