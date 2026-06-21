import glob, re

files = glob.glob("templates/*.html")
report = []

for path in files:
    with open(path, encoding="utf-8") as f:
        text = f.read()
    original = text

    # Пропускаем home (там иконка уже есть руками)
    if 'text.textContent = "Московская"' in text:
        report.append((path, "⏭️  иконка уже есть"))
        continue

    # Ищем ветку light с текстом "Светлая тема" и закрывающую скобку перед else,
    # вставляем ветку kreml ПЕРЕД "} else {"
    pattern = r'(if \(text\) text\.textContent = "Светлая тема";\s*\n\s*\})\s*else\s*\{'
    replacement = (
        r'\1 else if (themeMode === "kreml") {\n'
        r'                if (icon) icon.textContent = "🏛️";\n'
        r'                if (text) text.textContent = "Московская";\n'
        r'            } else {'
    )
    text = re.sub(pattern, replacement, text)

    if text != original:
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        report.append((path, "✅ иконка добавлена"))
    else:
        report.append((path, "⚠️  НЕ найден паттерн — смотреть руками"))

print("\n=== ОТЧЁТ ИКОНКИ ===")
for p, s in report:
    print(f"{s}  {p}")
