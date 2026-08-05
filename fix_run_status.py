"""
Надёжно добавляет ключ water_dashboard в словарь run_status.
Ищет именно run_status (по уникальному маркеру внутри словаря),
а не другие похожие блоки.
"""
import ast
import shutil
from pathlib import Path

APP = Path("app.py")
BOM = b"\xef\xbb\xbf"

raw = APP.read_bytes()
if raw.startswith(BOM):
    raw = raw[len(BOM):]
    print("🧹 Удалён BOM")

code = raw.decode("utf-8")

# Уникальный маркер run_status — только в нём есть "camera_prescriptions": {
if "water_dashboard" in code and '"water_dashboard":' in code:
    print("✅ Уже есть. Ничего не делаем.")
else:
    marker = '"mgkh_rm":'
    # Ищем вхождение, после которого идёт "camera_prescriptions"
    idx = -1
    for m in [i for i in range(len(code)) if code.startswith(marker, i)]:
        tail = code[m:m+500]
        if '"camera_prescriptions":' in tail:
            idx = m
            break

    if idx == -1:
        # Fallback: ищем маркер в блоке run_status = {
        rs_start = code.find("run_status = {")
        if rs_start != -1:
            idx = code.find('"mgkh_rm":', rs_start)

    if idx == -1:
        print("❌ Не нашёл подходящее место в run_status")
        raise SystemExit(1)

    # Ищем конец блока mgkh_rm: {...},
    brace_start = code.find("{", idx)
    depth = 0
    end = brace_start
    for i, ch in enumerate(code[brace_start:], brace_start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                end = i
                break

    # Ищем запятую после закрывающей скобки
    comma = code.find(",", end)
    insert_at = comma + 1

    new_block = '''
    "water_dashboard": {
        "running": False,
        "stage": "Ожидание запуска",
        "message": "Система готова к обновлению сводного дашборда.",
        "last_error": "",
    },'''

    new_code = code[:insert_at] + new_block + code[insert_at:]

    # Проверка синтаксиса
    try:
        ast.parse(new_code)
    except SyntaxError as e:
        print(f"❌ Синтаксическая ошибка после правки: {e}")
        raise SystemExit(1)

    # Бэкап и запись
    import time
    bak = APP.with_name(f"app.py.bak.{int(time.time())}")
    shutil.copy2(APP, bak)
    print(f"📦 Бэкап: {bak.name}")
    APP.write_bytes(new_code.encode("utf-8"))
    print("✅ Добавлен ключ water_dashboard в run_status")

print("Готово. Перезапусти сервер: python -m uvicorn app:app --reload --port 8000")