import json
import re
import time

from playwright.sync_api import sync_playwright

from services.water_dashboard.config import (
    DEBUG_DIR, HEADLESS, PAGE_WAIT_SECONDS, PLAYWRIGHT_PROFILE_DIR, SOURCES,
)

TABLES_JS = """
() => {
    const norm = (s) => (s || '').trim();
    const out = [];
    for (const table of Array.from(document.querySelectorAll('table'))) {
        let heads = Array.from(table.querySelectorAll('thead th'));
        if (!heads.length) {
            const fr = table.querySelector('tr');
            if (fr) heads = Array.from(fr.querySelectorAll('th, td'));
        }
        const headers = heads.map((h) => norm(h.textContent).toLowerCase());
        if (!headers.length) continue;

        const rows = [];
        for (const tr of Array.from(table.querySelectorAll('tbody tr'))) {
            const cells = Array.from(tr.querySelectorAll('td')).map((td) => norm(td.textContent));
            if (cells.length >= 2) rows.push(cells);
        }
        if (rows.length >= 3) out.push({ headers, rows });
    }
    return out;
}
"""

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _select_penultimate_date(page, sid):
    """Открывает фильтр «Дата» и выбирает предпоследнюю дату (для НВОС)."""
    def emit(msg):
        print(f"[{sid}] {msg}", flush=True)

    # 1. Находим контрол даты по лейблу «Дата» и помечаем его data-pw-id
    emit("Открываю фильтр «Дата»...")
    try:
        control_id = page.evaluate(
            """
            () => {
                const labels = Array.from(document.querySelectorAll('[data-qa="chartkit-control-title"]'));
                const dateLabel = labels.find(l => (l.textContent || '').trim().toLowerCase().startsWith('дата'));
                if (!dateLabel) return null;
                let node = dateLabel.parentElement;
                for (let i = 0; i < 5 && node; i++) {
                    const c = node.querySelector('[data-qa="chartkit-control-select"]')
                            || node.querySelector('.yc-select-control');
                    if (c) {
                        c.setAttribute('data-pw-id', 'nvos-date');
                        return 'nvos-date';
                    }
                    node = node.parentElement;
                }
                return null;
            }
            """
        )
        if not control_id:
            emit("⚠️ Контрол даты не найден")
            return
        emit(f"Контрол найден")
    except Exception as e:
        emit(f"⚠️ Ошибка поиска контрола: {e}")
        return

    # 2. Кликаем по контролу — открывается дропдаун
    try:
        page.locator('[data-pw-id="nvos-date"]').first.click(timeout=5000)
        page.wait_for_timeout(2000)
        emit("Dropdown открыт")
    except Exception as e:
        emit(f"⚠️ Не удалось кликнуть контрол: {e}")
        return

    # 3. Собираем опции дат — ищем в .yc-select-popup (работает в DataLens)
    options = []
    seen = set()

    # Стратегия A: yc-select-popup
    try:
        portals = page.locator('.yc-select-popup')
        for p in range(min(portals.count(), 5)):
            inner = portals.nth(p).locator('*')
            for i in range(min(inner.count(), 500)):
                item = inner.nth(i)
                try:
                    text = (item.inner_text() or '').strip()
                except Exception:
                    continue
                if DATE_RE.match(text) and text not in seen:
                    seen.add(text)
                    options.append((text, item))
    except Exception:
        pass

    # Стратегия B: фолбэк по всему документу
    if not options:
        for sel in ('[role="option"]', '.yc-select-option', '.popup *', 'li'):
            try:
                locs = page.locator(sel)
                for i in range(min(locs.count(), 800)):
                    item = locs.nth(i)
                    try:
                        text = (item.inner_text() or '').strip()
                    except Exception:
                        continue
                    if DATE_RE.match(text) and text not in seen:
                        seen.add(text)
                        options.append((text, item))
            except Exception:
                continue
            if options:
                break

    emit(f"Найдено опций дат: {len(options)}")
    if len(options) < 2:
        emit("⚠️ Недостаточно опций — оставляю фильтр как есть")
        return

    # 4. Сортируем по дате и берём ПРЕДПОСЛЕДНЮЮ
    options.sort(key=lambda x: x[0])
    target_text, target_item = options[-2]
    emit(f"Выбираю предпоследнюю дату: {target_text}")

    # 5. Клик (обычный + fallback по bounding_box)
    try:
        target_item.scroll_into_view_if_needed(timeout=3000)
        page.wait_for_timeout(200)
        target_item.click(timeout=3000)
        emit("✓ Клик по опции сработал")
    except Exception:
        try:
            box = target_item.bounding_box()
            if box:
                page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                emit("✓ Клик по bounding_box сработал")
            else:
                emit("⚠️ У опции нет bounding_box")
                return
        except Exception as e:
            emit(f"⚠️ Не удалось выбрать дату: {e}")
            return

    # 6. Ждём пересчёт виджетов
    page.wait_for_timeout(3000)
    emit("✓ Пересчёт данных завершён")


def _wait_for_content(page):
    """Умное ожидание: ждём появления таблиц/виджетов или fallback на таймер."""
    try:
        page.wait_for_selector(
            'table, .dl-widget, [class*="chart"], [class*="widget"]',
            timeout=15000,
        )
        # + небольшой буфер для финальной отрисовки чисел
        page.wait_for_timeout(1500)
    except Exception:
        # fallback: ждём фиксированное время
        page.wait_for_timeout(PAGE_WAIT_SECONDS * 1000)


def scrape_all():
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    PLAYWRIGHT_PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    extractions = {}

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PLAYWRIGHT_PROFILE_DIR),
            headless=HEADLESS,
            viewport={"width": 1440, "height": 1100},
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = context.pages[0] if context.pages else context.new_page()

        for src in SOURCES:
            sid = src["id"]
            print(f"STAGE: {src['name']}")
            try:
                page.goto(src["url"], wait_until="domcontentloaded", timeout=90000)

                # Умное ожидание появления контента
                _wait_for_content(page)

                # Специальная обработка для НВОС: выбор предпоследней даты
                if sid == "nvos":
                    _select_penultimate_date(page, sid)

                tables = page.evaluate(TABLES_JS)
                text = page.evaluate("() => document.body.innerText")

                extractions[sid] = {"tables": tables, "text": text}

                with open(DEBUG_DIR / f"{sid}.json", "w", encoding="utf-8") as f:
                    json.dump({"url": src["url"], "tables": tables, "text": text},
                              f, ensure_ascii=False, indent=2)

                print(f"[saved] {sid}: таблиц={len(tables)}")
            except Exception as e:
                print(f"[warn] {sid}: {e}")
                extractions[sid] = {"tables": [], "text": ""}

        context.close()

    return extractions