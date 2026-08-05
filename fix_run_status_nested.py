"""Исправляет вложенный water_dashboard внутри mgkh_rm в run_status."""
import ast
import re
import shutil
import time
from pathlib import Path

APP = Path("app.py")
BOM = b"\xef\xbb\xbf"

raw = APP.read_bytes()
if raw.startswith(BOM):
    raw = raw[len(BOM):]
    print("🧹 BOM удалён")
code = raw.decode("utf-8")

# Шаг 1. Удаляем вложенный "water_dashboard": {...} внутри mgkh_rm
# Ищем "mgkh_rm": { ... вложенный "water_dashboard" ... }
pattern = re.compile(
    r'(\s*)"water_dashboard":\s*\{[^}]*"running"[^}]*"stage"[^}]*"message"[^}]*"last_error"[^}]*\},',
    re.DOTALL,
)

# Ищем внутри блока mgkh_rm
mgkh_start = code.find('"mgkh_rm":')
mgkh_end = code.find('"cds":', mgkh_start) if mgkh_start != -1 else -1

if mgkh_start != -1 and mgkh_end != -1:
    chunk = code[mgkh_start:mgkh_end]
    m = pattern.search(chunk)
    if m:
        code = code[:mgkh_start] + chunk[:m.start()] + chunk[m.end():] + code[mgkh_end:]
        print("🗑️ Удалён вложенный water_dashboard из mgkh_rm")
    else:
        print("ℹ️ Вложенного water_dashboard не найдено")
else:
    print("⚠️ Блок mgkh_rm не найден")

# Шаг 2. Добавляем water_dashboard как отдельный ключ run_status ПОСЛЕ mgkh_rm
if '"water_dashboard":' not in code:
    # Находим конец блока mgkh_rm: ищем его закрывающую },
    mgkh_start = code.find('"mgkh_rm":')
    brace = code.find("{", mgkh_start)
    depth = 0
    end = brace
    for i, ch in enumerate(code[brace:], brace):
        if ch == "{": depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                end = i
                break
    comma = code.find(",", end)
    insert_at = comma + 1

    new_block = '''
    "water_dashboard": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к обновлению сводного дашборда.",
        "last_error": "",
    },'''
    code = code[:insert_at] + new_block + code[insert_at:]
    print("✅ Добавлен water_dashboard как ключ run_status")
else:
    print("ℹ️ water_dashboard уже есть в run_status")

# Проверка
try:
    ast.parse(code)
    print("✅ Синтаксис OK")
except SyntaxError as e:
    print(f"❌ Синтаксическая ошибка: строка {e.lineno}: {e.msg}")
    raise SystemExit(1)

# Бэкап и запись
bak = APP.with_name(f"app.py.bak.{int(time.time())}")
shutil.copy2(APP, bak)
print(f"📦 Бэкап: {bak.name}")
APP.write_bytes(code.encode("utf-8"))
print("Готово. Перезапусти сервер.")