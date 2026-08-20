"""Эталонно-сложные HTML и прогон PRO-конвертера в обоих режимах."""
from pathlib import Path
from services.tools import pptx_converter as c

T = Path("tests_html")
T.mkdir(exist_ok=True)

# ── 1. Дизайнерский: только div + inline-стили, без семантики ──
(T / "pro1_styled.html").write_text("""<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8"></head>
<body style="margin:0;font-family:Arial,sans-serif;color:#222;">
<div style="background:linear-gradient(135deg,#1b5e20,#43a047);color:#fff;padding:48px;">
  <div style="font-size:38px;font-weight:bold;">Итоги недели ЖКХ Московской области</div>
  <div style="font-size:18px;margin-top:8px;">Неделя 33 · 10–16 августа 2026</div>
</div>
<div style="padding:36px 48px;">
  <div style="font-size:28px;font-weight:bold;color:#1b5e20;">Аварийность</div>
  <div style="font-size:16px;line-height:1.6;">За неделю 14 технологических нарушений: 5 — ХВС, 4 — отопление, 3 — электроснабжение.</div>
  <div style="font-size:16px;line-height:1.6;">Среднее время устранения — 4,2 часа при плане 4,0. Затянутые: Серпухов (6,5 ч), Клин (7,1 ч).</div>
  <div style="font-size:28px;font-weight:bold;color:#1b5e20;">Диспетчеризация</div>
  <div style="font-size:16px;line-height:1.6;">Передано 1 245 заявок, закрыто 1 108 (89%). Повторных обращений 3,4% (норма ≤ 4%).</div>
  <div style="font-size:28px;font-weight:bold;color:#1b5e20;">Поручения</div>
  <div style="font-size:16px;line-height:1.6;">Кураторам до 20.08 закрыть просрочку; подготовить свод к совещанию у министра.</div>
</div>
</body></html>""", encoding="utf-8")

# ── 2. Имитация pdf2htmlEX: страницы .pf, текст в div.t ──
(T / "pro2_pdf.html").write_text("""<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>
<div id="page1" class="pf w0" data-page-no="1"><div class="pc pc1 w0">
  <div class="t m0 x0 h1 y0 ff0 fs3 fc0">ОТЧЁТ О КАЧЕСТВЕ ВОДОСНАБЖЕНИЯ</div>
  <div class="t m0 x0 h2 y1 ff0 fs1 fc0">Московская область · август 2026</div>
  <div class="t m0 x0 h3 y2 ff0 fs1 fc0">Подготовлен ЖКХ-Центром</div>
</div></div>
<div id="page2" class="pf w0" data-page-no="2"><div class="pc pc2 w0">
  <div class="t m0 x0 h1 y0 ff0 fs2 fc0">1. Ключевые показатели</div>
  <div class="t m0 x0 h4 y1 ff0 fs1 fc0">Выполнение промывок — 69% от плана.</div>
  <div class="t m0 x0 h4 y2 ff0 fs1 fc0">Просроченные задачи — 864, из них 132 свыше 20 дней.</div>
  <div class="t m0 x0 h4 y3 ff0 fs1 fc0">Собираемость НВОС — 80% при плане 95%.</div>
</div></div>
<div id="page3" class="pf w0" data-page-no="3"><div class="pc pc3 w0">
  <div class="t m0 x0 h1 y0 ff0 fs2 fc0">2. Проблемные округа</div>
  <div class="t m0 x0 h4 y1 ff0 fs1 fc0">Красногорск — явка 38%, просрочка 112 задач.</div>
  <div class="t m0 x0 h4 y2 ff0 fs1 fc0">Клин — явка 40%, просрочка 98 задач.</div>
  <div class="t m0 x0 h4 y3 ff0 fs1 fc0">Вывод: усилить контроль, совещание с кураторами до 15.08.</div>
</div></div>
</body></html>""", encoding="utf-8")

# ── 3. Глубокая вложенность + табличная верстка, strong-шапки ──
(T / "pro3_nested.html").write_text("""<html><body>
<table width="100%"><tr><td>
<table><tr><td><div><div><div><strong>Контроль воды</strong></div></div></div></td></tr>
<tr><td><div><div>Промывки выполнены на 69% от плана, дезинфекция — 77%.</div></div></td></tr>
<tr><td><div><div>Лабораторный контроль — 91% проб в норме.</div></div></td></tr></table>
<table><tr><td><div><div><div><strong>Просроченные задачи</strong></div></div></div></td></tr>
<tr><td><div><div>Всего 864; более 20 дней — 132; от 1 до 20 дней — 418.</div></div></td></tr></table>
<table><tr><td><div><div><div><strong>Видеонаблюдение</strong></div></div></div></td></tr>
<tr><td><div><div>Онлайн 178 камер, офлайн 22, не подключены 14.</div></div></td></tr></table>
</td></tr></table>
</body></html>""", encoding="utf-8")

# ── 4. Микс: стили + списки + таблица + цитата + картинка + заметки ──
(T / "pro4_mixed.html").write_text("""<html><body>
<div style="font-size:34px;font-weight:bold;">Сводный отчёт по платформе</div>
<p style="font-size:15px;">Август 2026 · Нейрона ИИ</p>
<div style="font-size:26px;font-weight:bold;">Модули мониторинга</div>
<ul>
  <li>Заполненность данных — 92%</li>
  <li>Просрочки — 864 задачи</li>
  <li>Камеры — 83% онлайн</li>
</ul>
<div style="font-size:26px;font-weight:bold;">Таблица округов</div>
<table border="1">
 <tr><th>Округ</th><th>Явка</th></tr>
 <tr><td>Красногорск</td><td>38%</td></tr>
 <tr><td>Клин</td><td>40%</td></tr>
</table>
<blockquote>Все камеры, не работающие более 30 суток, подлежат замене.</blockquote>
<img src="tests_html/chart_green.png">
<div class="notes">Докладывает зам. министра; регламент 10 минут.</div>
</body></html>""", encoding="utf-8")

print("✅ 4 сложных HTML созданы\n")

# ── Прогон: умный режим ──
for f in sorted(T.glob("pro*.html")):
    html = f.read_text(encoding="utf-8")
    slides = c.parse_html_to_slides_pro(html)
    print(f"{f.name}: слайдов={len(slides)}")
    for i, s in enumerate(slides[:6], 1):
        print(f"   {i}. {s['title'][:55] or '(без заголовка)'} | буллетов={len(s['bullets'])}, таблиц={len(s['tables'])}")
    out = c.build_pptx(slides, None, f"pro_{f.stem}.pptx")
    print(f"   → {out}\n")

# ── Прогон: скриншот-режим для pdf-подобного ──
try:
    imgs = c.render_slides_images((T / "pro2_pdf.html").read_text(encoding="utf-8"), "pro2_pdf")
    out = c.build_pptx_from_images(imgs, "pro2_shots.pptx")
    print(f"✅ shots-режим: {len(imgs)} слайдов → {out}")
except Exception as e:
    print("⚠️ shots-режим недоступен:", e)

print("\nОткрывай: start data\\tools\\output")