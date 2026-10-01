import json
import re
import sys
import datetime as _dt
from urllib.parse import urlsplit

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
    const seen = new Set();
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
        const key = JSON.stringify({headers, rows});
        if (rows.length >= 1 && !seen.has(key)) {
            seen.add(key);
            out.push({ headers, rows });
        }
    }
    return out;
}
"""

# Read values inside their own indicator. Adjacent body text also contains
# organisation rows and chart axes and must never be treated as KPI cards.
WIDGETS_JS = """
() => Array.from(document.querySelectorAll('[data-qa="chart-widget"]')).flatMap(widget => {
    const label = (widget.querySelector('.widget-header')?.innerText || '')
        .split(/\\n/).map(s => s.trim()).filter(s => s && !/^Ещ[её]\\s+\\d+$/.test(s)).join(' ');
    if (!label || widget.getAttribute('aria-busy') === 'true') return [];
    return Array.from(widget.querySelectorAll('.chartkit-indicator__item')).map(item => {
        const value = (item.querySelector('.chartkit-indicator__item-value')?.innerText || '').trim();
        const caption = (widget.querySelector('.dl-widget__description')?.innerText ||
                         (item.innerText || '').replace(value, '')).trim();
        return {label, value, caption};
    }).filter(item => item.value);
})
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
            return target_date.isoformat()

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
                return target_date.isoformat()
            page.wait_for_timeout(250)
        raise SourceReadError('date')
    except SourceReadError:
        raise
    except Exception:
        # Browser exceptions can contain page contents; expose a stable code only.
        raise SourceReadError('date') from None


def _wait_for_content(page):
    """Wait for rendered charts, not an empty document with idle network."""
    page.wait_for_selector('[data-qa="chart-widget"]', state='attached', timeout=45000)
    page.wait_for_function("""() => {
        const charts = Array.from(document.querySelectorAll('[data-qa="chart-widget"]'));
        return charts.length > 0 && charts.every(chart => chart.getAttribute('aria-busy') !== 'true');
    }""", timeout=45000)
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass


def _save_browser_diagnostics(page, sid, response, errors):
    try:
        diagnostics = {
            'status': response.status if response else None,
            'title': page.title(), 'errors': errors[-30:],
            'body': page.evaluate("() => document.body.outerHTML.slice(0,2000)"),
        }
        (DEBUG_DIR / f'{sid}-browser.json').write_text(
            json.dumps(diagnostics, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception:
        # Optional diagnostics must not discard a successful source extraction.
        pass


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


def _read_edo_rankings(page, source_url):
    """Read the full organisation table separately from overview KPI widgets.

    The overview's eight-row ECP table is a filtered list of lagging entries,
    so it cannot provide a ranking of the best organisations.
    """
    from services.water_dashboard.details import is_edo_ranking_table
    table_url = urlsplit(source_url)._replace(query='tab=EL', fragment='').geturl()
    print('STAGE: ЭДО — полная таблица РСО', flush=True)
    try:
        page.goto(table_url, wait_until='domcontentloaded', timeout=90000)
        _wait_for_content(page)
        tables = page.evaluate(TABLES_JS)
        for frame in page.frames[1:]:
            try:
                tables.extend(frame.evaluate(TABLES_JS))
            except Exception:
                pass
        tables = [table for table in tables if is_edo_ranking_table(table)]
        if not tables or not any(table.get('rows') for table in tables):
            raise SourceReadError('source_error')
        return {'ranking_tables': tables, 'ranking_url': table_url, 'ranking_error': ''}
    except Exception as exc:
        # The successfully read overview remains valid. Do not turn its ECP
        # laggards into a fabricated "best" ranking when this tab fails.
        print(f'[warn] edo_rso rankings: {exc}', flush=True)
        return {'ranking_tables': [], 'ranking_url': table_url,
                'ranking_error': 'Полная таблица РСО не получена. ' + MESSAGES[error_code(exc)]}


def scrape_all(source_ids=None):
    selected = {source['id'] for source in SOURCES} if source_ids is None else set(source_ids)
    known = {source['id'] for source in SOURCES}
    if not selected or not selected <= known:
        raise ValueError('Unknown or empty water dashboard source selection')
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
            browser_errors = []

            def failed_request(request):
                url = urlsplit(request.url)
                browser_errors.append({'kind': 'request', 'host': url.netloc,
                                       'path': url.path, 'resource': request.resource_type,
                                       'error': str(request.failure)[:300]})

            page.on('requestfailed', failed_request)
            page.on('pageerror', lambda exc: browser_errors.append({'kind': 'script', 'error': str(exc)[:500]}))

            for src in SOURCES:
                sid = src["id"]
                if sid not in selected:
                    continue
                print(f"STAGE: {src['name']}", flush=True)
                browser_errors.clear()
                response = None
                text = ""
                try:
                    response = page.goto(src["url"], wait_until="domcontentloaded", timeout=90000)

                    _wait_for_content(page)

                    data_date = None
                    if sid == "nvos":
                        data_date = _select_latest_date(page, sid)
                        stab_wait_after_date(page)
                        _wait_for_content(page)

                    tables = page.evaluate(TABLES_JS)
                    widgets = page.evaluate(WIDGETS_JS)
                    text = page.evaluate("() => document.body.innerText")
                    for frame in page.frames[1:]:
                        try:
                            tables.extend(frame.evaluate(TABLES_JS))
                            widgets.extend(frame.evaluate(WIDGETS_JS))
                            text += "\n" + frame.evaluate("() => document.body.innerText")
                        except Exception:
                            pass
                    if re.search(r"Войдите в аккаунт|Нет доступа к|Авторизуйтесь", text, re.I) and not tables:
                        raise SourceReadError("access")
                    if not tables and re.search(r"Внутренняя ошибка|Something went wrong", text, re.I):
                        raise SourceReadError("source_error")
                    if not tables and re.search(r"Я не робот|Подтвердите.{0,50}человек|SmartCaptcha", text, re.I):
                        raise SourceReadError("captcha")

                    extractions[sid] = {"tables": tables, "widgets": widgets, "text": text, "data_date": data_date}
                    if sid == "edo_rso":
                        extractions[sid].update(_read_edo_rankings(page, src["url"]))

                    with open(DEBUG_DIR / f"{sid}.json", "w", encoding="utf-8") as f:
                        json.dump({"url": src["url"], **extractions[sid]},
                                  f, ensure_ascii=False, indent=2)

                    print(f"[saved] {sid}: таблиц={len(tables)}", flush=True)
                except Exception as e:
                    print(f"[warn] {sid}: {e}", flush=True)
                    code = error_code(e)
                    if any(item.get('resource') == 'script' and item.get('host') == 'yastatic.net'
                           for item in browser_errors):
                        code = 'assets'
                    # Frequency notes can remain readable when chart queries
                    # fail. Read them from this source only, never the previous
                    # tab after a failed navigation. They are not metric data.
                    refresh_text = text
                    if response is not None and not refresh_text:
                        try:
                            refresh_text = page.evaluate("() => document.body.innerText")
                        except Exception:
                            refresh_text = ""
                    extractions[sid] = {"tables": [], "widgets": [], "text": "", "error": MESSAGES[code],
                                        "refresh_text": refresh_text if isinstance(refresh_text, str) else ""}
                finally:
                    _save_browser_diagnostics(page, sid, response, browser_errors)

        finally:
            context.close()

    return extractions
