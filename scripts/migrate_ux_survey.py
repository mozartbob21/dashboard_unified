"""Explicit migration entry point; startup also applies this idempotent migration."""
from pathlib import Path
import argparse
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from utils import db


def main():
    parser = argparse.ArgumentParser(description="Создать таблицы опроса удобства сервиса")
    parser.add_argument("--database", type=Path, help="Путь к SQLite; по умолчанию БД платформы")
    args = parser.parse_args()
    if args.database:
        db.DB_FILE = args.database.resolve()
    with db.get_db_connection() as conn:
        applied = conn.execute("SELECT name FROM app_schema_migrations WHERE name=?",
                               ("20261013_ux_survey",)).fetchone()
    if not applied:
        raise RuntimeError("Миграция опроса не применена")
    print("Миграция 20261013_ux_survey применена. Существующие ответы сохранены.")


if __name__ == "__main__":
    main()
