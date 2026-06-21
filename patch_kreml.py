import glob, re

files = glob.glob("templates/*.html")
report = []

for path in files:
    with open(path, encoding="utf-8") as f:
        text = f.read()
    original = text

    # --- Правка 1: добавить ветку kreml в цикл переключения ---
    # Ищем "nextMode = "dark";" и добавляем после блока ветку kreml.
    # Меняем строку '} else {\n ... nextMode = "auto-moscow";'
    # на вариант с kreml ПЕРЕД auto-moscow.
    text = re.sub(
        r'(nextMode = "dark";\s*\n\s*\})\s*else\s*\{\s*\n(\s*)nextMode = "auto-moscow";',
        r'\1 else if (currentMode === "dark") {\n\2    nextMode = "kreml";\n\2} else {\n\2nextMode = "auto-moscow";',
        text
    )

    # --- Правка 2: добавить иконку/текст для kreml ---
    # Ищем ветку light и добавляем после неё ветку kreml перед else.
    text = re.sub(
        r'(text\.textContent = "Светлая";\s*\n\s*\})\s*else\s*\{',
        r'\1 else if (themeMode === "kreml") {\n                if (icon) icon.textContent = "🏛️";\n                if (text) text.textContent = "Московская";\n            } else {',
        text
    )

    if text != original:
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        report.append((path, "✅ изменён"))
    else:
        report.append((path, "⏭️  без изменений"))

print("\n=== ОТЧЁТ ===")
for p, s in report:
    print(f"{s}  {p}")
