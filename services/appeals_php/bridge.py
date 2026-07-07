import json
import subprocess
from pathlib import Path

PHP_SCRIPT = Path(__file__).resolve().parent / "analyzer.php"


def analyze_appeal_via_php(subject: str, text: str) -> dict:
    # 1. Собираем вход в JSON-строку
    payload = json.dumps(
        {"subject": subject, "text": text},
        ensure_ascii=False,
    )

    # 2. Запускаем PHP, передаём payload в stdin
    process = subprocess.run(
        ["php", str(PHP_SCRIPT)],
        input=payload,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    # 3. Проверяем, не упал ли PHP
    if process.returncode != 0:
        raise RuntimeError(f"PHP error: {process.stderr}")

    # 4. Парсим ответ обратно в dict
    return json.loads(process.stdout)


# Быстрый тест: запусти этот файл напрямую
if __name__ == "__main__":
    result = analyze_appeal_via_php(
        subject="Нет воды",
        text="Уже неделю нет холодной воды по улице Ленина!",
    )
    print("Ответ от PHP:")
    print(json.dumps(result, ensure_ascii=False, indent=2))