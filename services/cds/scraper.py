"""
CDS Scraper v9 — период + "Вывести список" → Сохранить как → xlsx → pandas.
Fixes:
  - Meta+a → кроссплатформенный SELECT_ALL (Mac/Win/Linux)
  - "Еще"/"Ещё" — поддержка обоих вариантов
  - Улучшена обработка ошибок
"""
import asyncio
import calendar
import json
import platform
import re
import traceback
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from services.cds.config import SCREENSHOT_DIR
from services.cds.auth import login_to_cds
from services.cds.navigate import go_to_obrasheniya

RESULT_DIR = Path("data/cds")
RESULT_DIR.mkdir(parents=True, exist_ok=True)
XLSX_PATH = RESULT_DIR / "appeals.xlsx"

DEBUG = False

# Mac = Cmd+A, Windows/Linux = Ctrl+A
SELECT_ALL = "Meta+a" if platform.system() == "Darwin" else "Control+a"

# Кнопка ОК: 1С где-то рисует кириллицу "ОК", где-то латиницу "OK".
OK_PATTERN = re.compile(r"^\s*[ОO][КK]\s*$")


# ──────────────────────────── ХЕЛПЕРЫ ────────────────────────────


async def click_visible(page, text, exact: bool = False):
    """Кликает по ВИДИМОМУ элементу с текстом (1С дублирует пункты меню в DOM).
    text может быть строкой или re.Pattern."""
    items = (
        page.get_by_text(text, exact=exact)
        if isinstance(text, str)
        else page.get_by_text(text)
    )
    count = await items.count()
    for i in range(count - 1, -1, -1):
        el = items.nth(i)
        try:
            if await el.is_visible():
                await el.click()
                return True
        except Exception:
            continue
    raise RuntimeError(f"Видимый элемент '{text}' не найден (всего в DOM: {count})")


async def click_ok(page):
    """Жмёт ОК/OK (кириллица или латиница)."""
    await click_visible(page, OK_PATTERN)


async def find_input_by_value(page, *needles: str):
    """Ищет видимый input, чьё ЖИВОЕ значение содержит подстроку."""
    inputs = page.locator("input:visible")
    for i in range(await inputs.count()):
        el = inputs.nth(i)
        try:
            val = await el.input_value()
        except Exception:
            continue
        if any(n in val for n in needles):
            return el
    return None


# ──────────────────────────── ПЕРИОД ────────────────────────────


def get_default_period() -> tuple[str, str]:
    today = date.today()
    if today.month == 1:
        start = date(today.year - 1, 12, 1)
    else:
        start = date(today.year, today.month - 1, 1)
    last = calendar.monthrange(today.year, today.month)[1]
    end = date(today.year, today.month, last)
    return start.strftime("%d.%m.%Y"), end.strftime("%d.%m.%Y")


async def open_more_menu(page):
    """Открывает меню 'Ещё'/'Еще' — поддержка обоих вариантов написания."""
    for text in ("Ещё", "Еще", "Ещe"):
        loc = page.locator(f'text="{text}"').last
        try:
            if await loc.is_visible(timeout=2000):
                await loc.click()
                await page.wait_for_timeout(800)
                return
        except Exception:
            continue
    # Последний fallback — без проверки видимости
    await page.locator('text="Еще"').last.click()
    await page.wait_for_timeout(800)


async def set_period(page, date_from: str, date_to: str) -> bool:
    print(f"[CDS] Устанавливаем период: {date_from} — {date_to}")

    await open_more_menu(page)
    await click_visible(page, "Установить период")
    await page.wait_for_timeout(1200)

    inputs = page.locator('input[id*="ериод"]:visible')
    if await inputs.count() < 2:
        inputs = page.locator("input:visible")
        n = await inputs.count()
        first_input, second_input = inputs.nth(n - 2), inputs.nth(n - 1)
    else:
        first_input, second_input = inputs.nth(0), inputs.nth(1)

    await first_input.click()
    await page.keyboard.press(SELECT_ALL)
    await page.keyboard.type(date_from, delay=60)

    await second_input.click()
    await page.keyboard.press(SELECT_ALL)
    await page.keyboard.type(date_to, delay=60)

    await click_visible(page, "Выбрать", exact=True)
    print("[CDS] ✅ Период применён")
    await page.wait_for_timeout(5000)
    return True


# ──────────────────────────── ЭКСПОРТ ────────────────────────────


async def export_to_excel(page) -> Path:
    """Ещё → Вывести список → ОК → Меню(⋮) → Файл → Сохранить как → xlsx → OK."""
    print("[CDS] Вывести список...")
    await open_more_menu(page)
    await click_visible(page, "Вывести список")
    await page.wait_for_timeout(1500)
    await click_ok(page)

    # Документ формируется с задержкой. Критерий готовности — "меню реально
    # открылось и в нём виден пункт 'Файл'". Пробуем циклом до 180 сек.
    print("[CDS] Ждём документ и открываем меню (до 180 сек)...")
    await page.wait_for_timeout(5000)

    menu_selectors = [
        '[title*="Меню"]:visible',
        '[title*="Ещё"]:visible',
        '[title*="Еще"]:visible',
        'button:has-text("Меню"):visible',
    ]

    async def _file_item_visible() -> bool:
        loc = page.locator('text="Файл"')
        n = await loc.count()
        for i in range(n):
            try:
                if await loc.nth(i).is_visible():
                    return True
            except Exception:
                pass
        return False

    menu_opened = False
    for attempt in range(35):
        candidates = []
        for sel in menu_selectors:
            try:
                loc = page.locator(sel)
                for i in range(await loc.count()):
                    candidates.append(loc.nth(i))
            except Exception:
                pass

        for cand in reversed(candidates):
            try:
                await cand.click(timeout=3000)
            except Exception:
                continue
            await page.wait_for_timeout(800)
            if await _file_item_visible():
                menu_opened = True
                print(f"[CDS] Меню документа открыто (попытка {attempt + 1})")
                break
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(400)

        if menu_opened:
            break

        try:
            await page.keyboard.press("Alt+Minus")
            await page.wait_for_timeout(800)
            if await _file_item_visible():
                menu_opened = True
                print(f"[CDS] Меню открыто через Alt+Minus (попытка {attempt + 1})")
                break
            await page.keyboard.press("Escape")
        except Exception:
            pass

        await page.wait_for_timeout(4000)

    if not menu_opened:
        screenshot_path = str(SCREENSHOT_DIR / "cds_menu_not_found.png")
        await page.screenshot(path=screenshot_path)
        raise RuntimeError(
            "Не удалось открыть меню документа за 180 сек "
            f"(скриншот: {screenshot_path})"
        )

    await click_visible(page, "Файл")
    await page.wait_for_timeout(600)
    await click_visible(page, "Сохранить как")
    await page.wait_for_timeout(1200)

    # Поле "Тип" → F4 → Лист Excel2007
    print("[CDS] Выбираем формат xlsx...")
    type_input = await find_input_by_value(page, ".mxl", "абличный")
    if type_input is None:
        vis = page.locator("input:visible")
        type_input = vis.nth(await vis.count() - 1)

    await type_input.click()
    await page.wait_for_timeout(400)
    await page.keyboard.press("F4")
    await page.wait_for_timeout(900)

    opt = page.get_by_text("Лист Excel2007", exact=False)
    opt_visible = False
    for i in range(await opt.count()):
        try:
            if await opt.nth(i).is_visible():
                opt_visible = True
                break
        except Exception:
            pass
    if not opt_visible:
        await page.keyboard.press("Alt+ArrowDown")
        await page.wait_for_timeout(900)

    await click_visible(page, "Лист Excel2007")
    await page.wait_for_timeout(800)
    await page.screenshot(path=str(SCREENSHOT_DIR / "save_dialog_xlsx.png"))

    # OK (кириллица ИЛИ латиница!) → скачивание
    print("[CDS] Жмём OK, ждём скачивание...")
    async with page.expect_download(timeout=90000) as dl_info:
        await click_ok(page)

    download = await dl_info.value
    print(f"[CDS] Имя файла от сервера: {download.suggested_filename}")
    await download.save_as(XLSX_PATH)
    print(f"[CDS] ✅ Скачан файл: {XLSX_PATH}")
    return XLSX_PATH


# ──────────────────────────── ПАРСИНГ XLSX ────────────────────────────


def parse_excel(path: Path) -> list[dict]:
    raw = pd.read_excel(path, header=None)
    header_idx = 0
    for i in range(min(len(raw), 15)):
        vals = [str(v) for v in raw.iloc[i].tolist()]
        if any("Номер" in v for v in vals) and any("Дата" in v for v in vals):
            header_idx = i
            break

    df = pd.read_excel(path, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").fillna("")
    print(f"[CDS] Колонки: {list(df.columns)}")

    rename_map = {
        "Дата": "date",
        "Подразделение": "department",
        "Номер": "number",
        "Состояние обращения": "status",
        "Адрес обращения": "address",
        "Тип обращения": "type",
        "Причина обращения": "reason",
        "Срок исполнения": "deadline",
        "Тип заявки": "request_type",
        "Источник поступления": "source",
        "Заявитель": "applicant",
        "Исполнитель": "executor",
    }
    records = []
    for _, row in df.iterrows():
        rec = {str(k): str(v).strip() for k, v in row.items()}
        for ru, en in rename_map.items():
            if ru in rec:
                rec[en] = rec[ru]
        records.append(rec)
    return records


# ──────────────────────────── ОРКЕСТРАТОР ────────────────────────────


async def scrape_cds_appeals(date_from="", date_to="", headless=True) -> dict:
    pw = browser = page = None
    try:
        if not date_from or not date_to:
            date_from, date_to = get_default_period()

        print("=" * 60)
        print(f"CDS SCRAPER v9 — EXPORT LIST | {date_from} — {date_to}")
        print("=" * 60)

        pw, browser, context, page = await login_to_cds(headless=headless)

        if not await go_to_obrasheniya(page):
            return {"success": False, "error": "Навигация не удалась", "data": []}
        await page.wait_for_timeout(3000)

        if DEBUG:
            await page.pause()

        await set_period(page, date_from, date_to)
        xlsx = await export_to_excel(page)
        appeals = parse_excel(xlsx)

        result = {
            "success": True,
            "data": appeals,
            "count": len(appeals),
            "date_from": date_from,
            "date_to": date_to,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        }
        (RESULT_DIR / "result.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"\n[CDS] ✅ Готово! {len(appeals)} обращений")
        return result

    except Exception as e:
        print(f"[CDS] ❌ Ошибка: {e}")
        traceback.print_exc()
        if page:
            try:
                await page.screenshot(path=str(SCREENSHOT_DIR / "error.png"))
            except Exception:
                pass
        return {"success": False, "error": str(e), "data": [], "count": 0}
    finally:
        if browser:
            await browser.close()
        if pw:
            await pw.stop()


if __name__ == "__main__":
    import sys

    d_from = sys.argv[1] if len(sys.argv) > 1 else ""
    d_to = sys.argv[2] if len(sys.argv) > 2 else ""
    result = asyncio.run(scrape_cds_appeals(d_from, d_to, headless=False))
    if result["success"]:
        print(f"\n{'=' * 60}")
        print(f"✅ УСПЕХ: {result['count']} обращений")
        for i, r in enumerate(result["data"][:10]):
            print(
                f"  [{i + 1}] {r.get('date', '')} | "
                f"{r.get('number', '')} | {r.get('status', '')}"
            )
    else:
        print(f"\n❌ ОШИБКА: {result['error']}")
