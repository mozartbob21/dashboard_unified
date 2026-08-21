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
    """РћС‚РєСЂС‹РІР°РµС‚ С„РёР»СЊС‚СЂ В«Р”Р°С‚Р°В» Рё РІС‹Р±РёСЂР°РµС‚ РїСЂРµРґРїРѕСЃР»РµРґРЅСЋСЋ РґР°С‚Сѓ (РґР»СЏ РќР’РћРЎ)."""
    def emit(msg):
        print(f"[{sid}] {msg}", flush=True)

    # 1. РќР°С…РѕРґРёРј РєРѕРЅС‚СЂРѕР» РґР°С‚С‹ РїРѕ Р»РµР№Р±Р»Сѓ В«Р”Р°С‚Р°В» Рё РїРѕРјРµС‡Р°РµРј РµРіРѕ data-pw-id
    emit("РћС‚РєСЂС‹РІР°СЋ С„РёР»СЊС‚СЂ В«Р”Р°С‚Р°В»...")
    try:
        control_id = page.evaluate(
            """
            () => {
                const labels = Array.from(document.querySelectorAll('[data-qa="chartkit-control-title"]'));
                const dateLabel = labels.find(l => (l.textContent || '').trim().toLowerCase().startsWith('РґР°С‚Р°'));
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
            emit("вљ пёЏ РљРѕРЅС‚СЂРѕР» РґР°С‚С‹ РЅРµ РЅР°Р№РґРµРЅ")
            return
        emit(f"РљРѕРЅС‚СЂРѕР» РЅР°Р№РґРµРЅ")
    except Exception as e:
        emit(f"вљ пёЏ РћС€РёР±РєР° РїРѕРёСЃРєР° РєРѕРЅС‚СЂРѕР»Р°: {e}")
        return

    # 2. РљР»РёРєР°РµРј РїРѕ РєРѕРЅС‚СЂРѕР»Сѓ вЂ” РѕС‚РєСЂС‹РІР°РµС‚СЃСЏ РґСЂРѕРїРґР°СѓРЅ
    try:
        page.locator('[data-pw-id="nvos-date"]').first.click(timeout=5000)
        page.wait_for_timeout(2000)
        emit("Dropdown РѕС‚РєСЂС‹С‚")
    except Exception as e:
        emit(f"вљ пёЏ РќРµ СѓРґР°Р»РѕСЃСЊ РєР»РёРєРЅСѓС‚СЊ РєРѕРЅС‚СЂРѕР»: {e}")
        return

    # 3. РЎРѕР±РёСЂР°РµРј РѕРїС†РёРё РґР°С‚ вЂ” РёС‰РµРј РІ .yc-select-popup (СЂР°Р±РѕС‚Р°РµС‚ РІ DataLens)
    options = []
    seen = set()

    # РЎС‚СЂР°С‚РµРіРёСЏ A: yc-select-popup
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

    # РЎС‚СЂР°С‚РµРіРёСЏ B: С„РѕР»Р±СЌРє РїРѕ РІСЃРµРјСѓ РґРѕРєСѓРјРµРЅС‚Сѓ
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

    emit(f"РќР°Р№РґРµРЅРѕ РѕРїС†РёР№ РґР°С‚: {len(options)}")
    if len(options) < 2:
        emit("вљ пёЏ РќРµРґРѕСЃС‚Р°С‚РѕС‡РЅРѕ РѕРїС†РёР№ вЂ” РѕСЃС‚Р°РІР»СЏСЋ С„РёР»СЊС‚СЂ РєР°Рє РµСЃС‚СЊ")
        return

    # 4. РЎРѕСЂС‚РёСЂСѓРµРј РїРѕ РґР°С‚Рµ Рё Р±РµСЂС‘Рј РџР Р•Р”РџРћРЎР›Р•Р”РќР®Р®
    options.sort(key=lambda x: x[0])
    target_text, target_item = options[-2]
    emit(f"Р’С‹Р±РёСЂР°СЋ РїСЂРµРґРїРѕСЃР»РµРґРЅСЋСЋ РґР°С‚Сѓ: {target_text}")

    # 5. РљР»РёРє (РѕР±С‹С‡РЅС‹Р№ + fallback РїРѕ bounding_box)
    try:
        target_item.scroll_into_view_if_needed(timeout=3000)
        page.wait_for_timeout(200)
        target_item.click(timeout=3000)
        emit("вњ“ РљР»РёРє РїРѕ РѕРїС†РёРё СЃСЂР°Р±РѕС‚Р°Р»")
    except Exception:
        try:
            box = target_item.bounding_box()
            if box:
                page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                emit("вњ“ РљР»РёРє РїРѕ bounding_box СЃСЂР°Р±РѕС‚Р°Р»")
            else:
                emit("вљ пёЏ РЈ РѕРїС†РёРё РЅРµС‚ bounding_box")
                return
        except Exception as e:
            emit(f"вљ пёЏ РќРµ СѓРґР°Р»РѕСЃСЊ РІС‹Р±СЂР°С‚СЊ РґР°С‚Сѓ: {e}")
            return

    # 6. Р–РґС‘Рј РїРµСЂРµСЃС‡С‘С‚ РІРёРґР¶РµС‚РѕРІ
    page.wait_for_timeout(3000)
    emit("вњ“ РџРµСЂРµСЃС‡С‘С‚ РґР°РЅРЅС‹С… Р·Р°РІРµСЂС€С‘РЅ")


def _wait_for_content(page):
    """Р–РґС‘Рј СЂРµР°Р»СЊРЅСѓСЋ РіРѕС‚РѕРІРЅРѕСЃС‚СЊ РІРёРґР¶РµС‚РѕРІ: СЃРµС‚СЊ СЃРїРѕРєРѕР№РЅР°, СЃРїРёРЅРЅРµСЂС‹ РёСЃС‡РµР·Р»Рё, РїР°СѓР·Р°."""
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

                # РЈРјРЅРѕРµ РѕР¶РёРґР°РЅРёРµ РїРѕСЏРІР»РµРЅРёСЏ РєРѕРЅС‚РµРЅС‚Р°
                _wait_for_content(page)

                # РЎРїРµС†РёР°Р»СЊРЅР°СЏ РѕР±СЂР°Р±РѕС‚РєР° РґР»СЏ РќР’РћРЎ: РІС‹Р±РѕСЂ РїСЂРµРґРїРѕСЃР»РµРґРЅРµР№ РґР°С‚С‹
                if sid == "nvos":
                    _select_penultimate_date(page, sid)
                    try:  # stab-wait: Р¶РґС‘Рј РїСЂРёРјРµРЅРµРЅРёСЏ С„РёР»СЊС‚СЂР° РґР°С‚С‹
                        page.wait_for_load_state("networkidle", timeout=15000)
                    except Exception:
                        pass
                    page.wait_for_timeout(5000)

                tables = page.evaluate(TABLES_JS)
                text = page.evaluate("() => document.body.innerText")

                extractions[sid] = {"tables": tables, "text": text}

                with open(DEBUG_DIR / f"{sid}.json", "w", encoding="utf-8") as f:
                    json.dump({"url": src["url"], "tables": tables, "text": text},
                              f, ensure_ascii=False, indent=2)

                print(f"[saved] {sid}: С‚Р°Р±Р»РёС†={len(tables)}")
            except Exception as e:
                print(f"[warn] {sid}: {e}")
                extractions[sid] = {"tables": [], "text": ""}

        context.close()

    return extractions