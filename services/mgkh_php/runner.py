"""
Мост Python <-> PHP для модуля МГКХ Redmine (проверка дат).

Запускает services/mgkh_php/runner.php через subprocess,
читает JSON из stdout, сохраняет в data/mgkh_rm/result.json.
"""
from pathlib import Path
import json
import subprocess
import sys

BASE_DIR = Path(__file__).resolve().parents[2]  # .../unified_dashboard
PHP_RUNNER = BASE_DIR / "services" / "mgkh_php" / "runner.php"
DATA_DIR = BASE_DIR / "data" / "mgkh_rm"
RESULT_FILE = DATA_DIR / "result.json"


def stage(msg: str) -> None:
    """Печать этапа — app.py ловит строки STAGE: и показывает в UI."""
    print(f"STAGE:{msg}", flush=True)


def main() -> int:
    stage("Запуск PHP-моста")

    if not PHP_RUNNER.is_file():
        print(f"Не найден PHP-скрипт: {PHP_RUNNER}", file=sys.stderr)
        return 1

    try:
        proc = subprocess.run(
            ["php", str(PHP_RUNNER)],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=180,
        )
    except FileNotFoundError:
        print("PHP не найден. Установите php и убедитесь, что он в PATH.", file=sys.stderr)
        return 1
    except subprocess.TimeoutExpired:
        print("PHP-скрипт выполнялся слишком долго и был остановлен.", file=sys.stderr)
        return 1

    # STAGE-строки из PHP (он пишет их в stderr) — пробрасываем в UI.
    for line in (proc.stderr or "").splitlines():
        line = line.strip()
        if line.startswith("STAGE:"):
            print(line, flush=True)
        elif line:
            print(f"[PHP] {line}", file=sys.stderr)

    stdout = (proc.stdout or "").strip()

    if not stdout:
        print("PHP не вернул данных (пустой stdout).", file=sys.stderr)
        return 1

    try:
        data = json.loads(stdout)
    except json.JSONDecodeError as e:
        print(f"Не удалось разобрать JSON от PHP: {e}", file=sys.stderr)
        print(f"stdout PHP: {stdout[:500]}", file=sys.stderr)
        return 1

    if not data.get("ok"):
        error = data.get("error", "Неизвестная ошибка PHP")
        print(f"PHP вернул ошибку: {error}", file=sys.stderr)
        return 1

    stage("Сохранение результата")
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    RESULT_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    stage("Готово")
    b = data.get("buckets", {})
    total = len(b.get("close", [])) + len(b.get("extend", [])) + len(b.get("rework", []))
    print(f"Готово: сохранено {total} задач в {RESULT_FILE}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())