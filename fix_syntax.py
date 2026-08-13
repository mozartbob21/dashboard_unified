"""Чинит строку с обратным слэшем в роуте удаления шаблона."""
import ast
from pathlib import Path

p = Path("app.py")
code = p.read_text(encoding="utf-8")

old = r'''if not name or "/" in name or "\" in name or ".." in name:'''
new = '''if not name or "/" in name or '\\\\' in name or ".." in name:'''

# Пробуем разные варианты — где-то слэш мог записаться по-разному
for bad in [
    'if not name or "/" in name or "\\" in name or ".." in name:',
    'if not name or "/" in name or "\\\\" in name or ".." in name:',
    'if not name or "/" in name or "\\\\"in name or ".." in name:',
]:
    if bad in code:
        code = code.replace(bad, new, 1)
        break
else:
    # Ищем по строке с "in name" в районе "template/delete"
    import re
    m = re.search(r'if not name or "/" in name or ".+" in name or "\.\." in name:', code)
    if m:
        code = code[:m.start()] + new + code[m.end():]
        print("✅ найден по regex и исправлен")
    else:
        print("❌ не нашёл проблемную строку")
        # Выведем строки 2760-2770
        lines = code.splitlines()
        for i in range(2759, min(2775, len(lines))):
            print(f"{i+1}: {lines[i]}")
        exit(1)

try:
    ast.parse(code)
except SyntaxError as e:
    print(f"❌ после правки синтаксис всё ещё битый: {e}")
    exit(1)

p.write_text(code, encoding="utf-8")
print("✅ app.py исправлен и прошёл проверку синтаксиса")