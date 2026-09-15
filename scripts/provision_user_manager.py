"""Создать отдельную учётную запись управления (повторный запуск не меняет пароль)."""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from services.auth import registration  # Инициализация email и таблиц регистрации.
from services.auth.accounts import provision_manager

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--username", default="user_manager")
    args = parser.parse_args()
    print(json.dumps(provision_manager(args.username), ensure_ascii=False))
