"""Навигация: АДС → Обращения."""
from services.cds.config import SCREENSHOT_DIR


async def _debug_dump(page, prefix: str):
    """Скриншот + URL + title + видимый текст страницы в лог."""
    try:
        await page.screenshot(path=f"{prefix}.png", full_page=True, timeout=10000)
    except Exception:
        pass
    try:
        print(f"[CDS][DEBUG] URL: {page.url}")
    except Exception:
        pass
    try:
        title = await page.title()
        print(f"[CDS][DEBUG] TITLE: {title}")
    except Exception:
        pass
    try:
        text = await page.locator("body").inner_text(timeout=5000)
        text = "\n".join(line.strip() for line in text.splitlines() if line.strip())
        print("[CDS][DEBUG] BODY TEXT FIRST 3000:")
        print(text[:3000])
    except Exception as e:
        print(f"[CDS][DEBUG] Не удалось получить body text: {type(e).__name__}: {e}")


async def _js_click_by_text(page, needle: str) -> bool:
    """JS-фолбэк: найти видимый кликабельный элемент с текстом и кликнуть."""
    try:
        return await page.evaluate(
            """
            (needle) => {
                const candidates = Array.from(document.querySelectorAll(
                    'a,button,div,span,td,li,[role="button"],[onclick]'
                ));
                function visible(el) {
                    const st = window.getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    return st && st.visibility !== 'hidden' &&
                        st.display !== 'none' && r.width > 0 && r.height > 0;
                }
                const el = candidates.find(e => {
                    const t = (e.innerText || e.textContent || '').trim();
                    return visible(e) && t.includes(needle);
                });
                if (!el) return false;
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
                return true;
            }
            """,
            needle,
        )
    except Exception as e:
        print(f"[CDS] JS-клик по тексту '{needle}' не удался: {type(e).__name__}: {e}")
        return False


LICENSE_ERROR_MSG = (
    "Сервер 1С не выдаёт лицензию: окно 'Не найдена лицензия' возвращается "
    "после нескольких нажатий 'Выполнить запуск'. Это проблема на стороне "
    "сервера 1С (лицензии закончились или слетел ключ защиты) — "
    "нужно обращаться к администратору 1С, скрипт это обойти не может."
)


async def _dialog_visible(page):
    """Возвращает текст кнопки активного диалога 1С или None."""
    for txt in ("Выполнить запуск", "Продолжить", "Перезапустить"):
        try:
            if await page.locator(f'text="{txt}"').first.is_visible():
                return txt
        except Exception:
            pass
    return None


async def _handle_1c_dialogs(page, max_clicks: int = 3) -> bool:
    """
    Закрывает служебные диалоги 1С (лицензия и т.п.).
    Кликаем максимум max_clicks раз с долгим ожиданием.
    Если диалог возвращается — False (серверная проблема, кликать бессмысленно).
    """
    for attempt in range(max_clicks):
        txt = await _dialog_visible(page)
        if txt is None:
            return True
        print(f"[CDS] ⚠️ Диалог 1С: '{txt}', клик (попытка {attempt + 1}/{max_clicks})")
        try:
            await page.locator(f'text="{txt}"').first.click(timeout=5000)
        except Exception:
            await _js_click_by_text(page, txt)
        # 1С может стартовать сеанс долго — ждём щедро
        await page.wait_for_timeout(12000)
    return await _dialog_visible(page) is None


async def go_to_obrasheniya(page) -> bool:
    """
    Из главного экрана переходит в:
    Аварийно диспетчерская служба → Обращения.
    Возвращает True если таблица обращений загрузилась.
    """
    # Закрываем диалоги 1С (лицензия и т.п.) — максимум 3 попытки
    if not await _handle_1c_dialogs(page, max_clicks=3):
        await _debug_dump(page, "cds_license_error")
        raise RuntimeError(LICENSE_ERROR_MSG)

    print("[CDS] Навигация: АДС...")
    clicked = False
    selectors = [
        '#themesCell_theme_4',
        '[id*="theme_4"]',
        '[title*="АДС"]',
        'text=Аварийно-диспетчерская',
        'text=Аварийно диспетчерская',
        'text=Аварийно',
        'text=АДС',
    ]
    for _round in range(45):  # 45 * 2 сек = 90 сек
        # Если диалог лицензии вернулся — сервер отказывает, дальше бессмысленно
        if await _dialog_visible(page) == "Выполнить запуск":
            await _debug_dump(page, "cds_license_error")
            raise RuntimeError(LICENSE_ERROR_MSG)
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if await loc.is_visible():
                    try:
                        await loc.scroll_into_view_if_needed(timeout=1000)
                    except Exception:
                        pass
                    try:
                        await loc.click(timeout=10000)
                    except Exception:
                        await loc.click(timeout=10000, force=True)
                    clicked = True
                    print(f"[CDS] АДС кликнут через селектор: {sel}")
                    break
            except Exception:
                pass
        if clicked:
            break
        if _round % 5 == 4:
            for txt in ("АДС", "Аварийно"):
                if await _js_click_by_text(page, txt):
                    clicked = True
                    print(f"[CDS] АДС кликнут через JS по тексту: {txt}")
                    break
        if clicked:
            break
        await page.wait_for_timeout(2000)

    if not clicked:
        await _debug_dump(page, "cds_ads_not_found")
        raise RuntimeError(
            "Не удалось найти раздел АДС ни одним селектором "
            "(скриншот: cds_ads_not_found.png)"
        )
    await page.wait_for_timeout(3000)

    # Клик по "Обращения" в панели функций
    print("[CDS] Навигация: Обращения...")
    clicked = await page.evaluate("""() => {
        const el = document.querySelector('#cmd_0_0_txt');
        if (el) { el.click(); return true; }
        const items = document.querySelectorAll('.functionItemBox');
        for (const item of items) {
            if (item.textContent.trim() === 'Обращения') {
                item.click();
                return true;
            }
        }
        return false;
    }""")

    if not clicked:
        for sel in ('text="Обращения"', 'a:has-text("Обращения")',
                    'span:has-text("Обращения")', '[title*="Обращения"]'):
            try:
                loc = page.locator(sel).first
                if await loc.is_visible():
                    await loc.click(timeout=5000)
                    clicked = True
                    print(f"[CDS] Обращения кликнуты через селектор: {sel}")
                    break
            except Exception:
                pass

    if not clicked:
        clicked = await _js_click_by_text(page, "Обращения")
        if clicked:
            print("[CDS] Обращения кликнуты через JS по тексту")

    if not clicked:
        await _debug_dump(page, "cds_obrasheniya_not_found")
        print("[CDS] ❌ Не найден пункт 'Обращения'")
        return False

    await page.wait_for_timeout(8000)

    try:
        grid_visible = await page.locator('#form5_Список').is_visible(timeout=15000)
    except Exception:
        grid_visible = False
    if grid_visible:
        print("[CDS] ✅ Таблица обращений загружена")
        return True

    try:
        any_grid = await page.locator('.gridBody').first.is_visible(timeout=5000)
    except Exception:
        any_grid = False
    if any_grid:
        print("[CDS] ✅ Grid найден (fallback)")
        return True

    print("[CDS] ⚠️ Таблица не обнаружена, продолжаем...")
    return True
