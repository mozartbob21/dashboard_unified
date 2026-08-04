# remove_bom.py
"""Удаляет BOM (U+FEFF) из начала файла app.py и проверяет синтаксис."""
import ast

path = "app.py"

with open(path, "rb") as f:
    raw = f.read()

# BOM в UTF-8 = байты EF BB BF
BOM = b"\xef\xbb\xbf"
had_bom = raw.startswith(BOM)

if had_bom:
    raw = raw[len(BOM):]
    with open(path, "wb") as f:
        f.write(raw)
    print(f"✅ BOM удалён из {path} ({len(raw)} байт)")
else:
    print(f"ℹ️  BOM в {path} не найден")

# Проверяем синтаксис
try:
    ast.parse(raw.decode("utf-8"))
    print("✅ Синтаксис OK — можно запускать сервер")
except SyntaxError as e:
    print(f"❌ Синтаксическая ошибка: строка {e.lineno}: {e.msg}")