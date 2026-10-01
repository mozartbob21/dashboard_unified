import json
import re
import sys
import datetime as _dt

from playwright.sync_api import sync_playwright

from services.water_dashboard.browser import launch_context
from services.water_dashboard.diagnostics import MESSAGES, SourceReadError, error_code

from services.water_dashboard.config import (
    DEBUG_DIR, HEADLESS, PLAYWRIGHT_PROFILE_DIR, SOURCES,
)

# Серверная консоль Windows (cp1251) не переживает символы вроде "✓" —
# переводим stdout в UTF-8, иначе этап nvos прерывался UnicodeEncodeError.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

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
        if (rows.length >= 1) out.push({ headers, rows });
    }
    return out;
}
"""

def _select_latest_date(page, sid):
    """Выбирает одну последнюю наступившую дату, не снимая текущий выбор."""
    def parse_date(value):
        value = (value or '').strip()
        for pattern, date_format in ((r'\d{4}-\d{2}-\d{2}', '%Y-%m-%d'),
                                     (r'\d{2}\.\d{2}\.\d{4}', '%d.%m.%Y')):
            if re.fullmatch(pattern, value):
                try:
                    return _dt.datetime.strptime(value, date_format).date()
                except ValueError:
                    pass
        return None

    def selected_date():
        return parse_date(control.inner_text(timeout=3000))

    try:
        control_id = page.evaluate(
            """
            () => {
                const labels = Array.from(document.querySelectorAll('[data-qa="chartkit-control-title"]'));
                const label = labels.find(l => (l.textContent || '').trim().toLowerCase().startsWith('дата'));
                const container = label && label.closest('[data-qa="chartkit-control"]');
                const select = container && container.querySelector('[data-qa="chartkit-control-select"]');
                if (!select) return null;
                select.setAttribute('data-pw-id', 'nvos-date');
                return 'nvos-date';
            }
            """
        )
        if not control_id:
            raise SourceReadError('date')
        control = page.locator('[data-pw-id="nvos-date"]').first
        current = selected_date()
        control.click(timeout=5000)
        popup = page.locator('.yc-select-popup:visible').first
        popup.wait_for(state='visible', timeout=5000)
        items = popup.locator('.yc-select-item[data-value]')
        options = []
        for index in range(items.count()):
            item = items.nth(index)
            date = parse_date(item.get_attribute('data-value'))
            if date is None:
                title = item.locator('.yc-select-item__title')
                if title.count():
                    date = parse_date(title.first.inner_text(timeout=3000))
            if date is not None and date <= _dt.date.today():
                options.append((date, index))
        if not options:
            raise SourceReadError('date')
        target_date, target_index = max(options, key=lambda option: option[0])
        if current == target_date:
            # DataLens uses toggles: clicking the selected item would clear it.
            page.keyboard.press('Escape')
            if selected_date() != target_date:
                raise SourceReadError('date')
            print(f'[{sid}] Последняя наступившая дата уже выбрана', flush=True)
            return

        clear = popup.get_by_text('Очистить', exact=True)
        if clear.count() and clear.first.is_visible():
            clear.first.click(timeout=3000)
            if not popup.is_visible():
                control.click(timeout=5000)
                popup.wait_for(state='visible', timeout=5000)
        elif current is not None:
            # Do not add another date to an existing multi-selection.
            raise SourceReadError('date')
        target_item = popup.locator('.yc-select-item[data-value]').nth(target_index)
        target_item.scroll_into_view_if_needed(timeout=3000)
        target_item.click(timeout=3000)
        for _ in range(20):
            if selected_date() == target_date:
                page.keyboard.press('Escape')
                page.wait_for_timeout(3000)
                print(f'[{sid}] Последняя наступившая дата выбрана', flush=True)
                return
            page.wait_for_timeout(250)
        raise SourceReadError('date')
    except SourceReadError:
        raise
    except Exception:
        # Browser exceptions can contain page contents; expose a stable code only.
        raise SourceReadError('date') from None


def _wait_for_content(page):
    """Ждём реальную готовность виджетов: сеть спокойна, спиннеры исчезли, пауза."""
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    try:
        page.wait_for_selector(
            ".dc-loader, .loader, [class*='spinner'], [class*='loading'], [class*='progress']",
            state="detached", timeout=8000)
    except Exception:
        pass
    page.wait_for_timeout(3500)


def stab_wait_after_date(page, max_rounds=15):
    """Ждём пересчёта виджета после выбора даты: текст страницы стабилен."""
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    print("[nvos-stab] дата выбрана, жду пересчёта...", flush=True)
    prev = None
    for _ in range(max_rounds):
        page.wait_for_timeout(2000)
        try:
            cur = page.evaluate("document.body.innerText")
        except Exception:
            cur = None
        if cur and cur == prev:
            break
        prev = cur
    print("[nvos-stab] пересчёт стабилен", flush=True)


def scrape_all():
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    PLAYWRIGHT_PROFILE_DIR.mkdir(parents=True, exist_ok=True)

    extractions = {}

    with sync_playwright() as p:
        context, browser_name = launch_context(
            p, profile_dir=PLAYWRIGHT_PROFILE_DIR, headless=HEADLESS,
        )
        print(f"STAGE: Браузер для сбора: {browser_name}", flush=True)
        try:
            page = context.pages[0] if context.pages else context.new_page()

            for src in SOURCES:
                sid = src["id"]
                print(f"STAGE: {src['name']}", flush=True)
                try:
                    page.goto(src["url"], wait_until="domcontentloaded", timeout=90000)

                    _wait_for_content(page)

                    if sid == "nvos":
                        _select_latest_date(page, sid)
                        stab_wait_after_date(page)

                    tables = page.evaluate(TABLES_JS)
                    text = page.evaluate("() => document.body.innerText")
                    for frame in page.frames[1:]:
                        try:
                            tables.extend(frame.evaluate(TABLES_JS))
                            text += "\n" + frame.evaluate("() => document.body.innerText")
                        except Exception:
                            pass
                    if re.search(r"Войдите в аккаунт|Нет доступа к|Авторизуйтесь", text, re.I) and not tables:
                        raise SourceReadError("access")
                    if not tables and re.search(r"Внутренняя ошибка|Something went wrong", text, re.I):
                        raise SourceReadError("source_error")
                    if not tables and re.search(r"Я не робот|Подтвердите.{0,50}человек|SmartCaptcha", text, re.I):
                        raise SourceReadError("captcha")

                    extractions[sid] = {"tables": tables, "text": text}

                    with open(DEBUG_DIR / f"{sid}.json", "w", encoding="utf-8") as f:
                        json.dump({"url": src["url"], "tables": tables, "text": text},
                                  f, ensure_ascii=False, indent=2)

                    print(f"[saved] {sid}: таблиц={len(tables)}", flush=True)
                except Exception as e:
                    print(f"[warn] {sid}: {e}", flush=True)
                    extractions[sid] = {"tables": [], "text": "", "error": MESSAGES[error_code(e)]}

        finally:
            context.close()

    return extractions
