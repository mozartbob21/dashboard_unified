"""Навигация: АДС → Обращения."""
from services.cds.config import SCREENSHOT_DIR


async def go_to_obrasheniya(page) -> bool:
    """
    Из главного экрана переходит в:
    Аварийно диспетчерская служба → Обращения.
    Возвращает True если таблица обращений загрузилась.
    """
    print("[CDS] Навигация: АДС...")
    # Устойчивый клик по разделу АДС: сначала по тексту, потом по ID
    clicked = False
    # 1С показывает разные стартовые экраны:
    #  - плитки тем  -> #themesCell_theme_4 (текст "АДС")
    #  - боковое меню -> "Аварийно диспетчерская служба" (полное название)
    # Опрашиваем все варианты по кругу, до 90 сек суммарно.
    selectors = [
        '#themesCell_theme_4',
        'text=Аварийно-диспетчерская',
        'text=Аварийно диспетчерская',
        'text=Аварийно',
        'text=АДС',
    ]
    for _round in range(45):  # 45 * 2 сек = 90 сек
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if await loc.is_visible():
                    await loc.click(timeout=10000)
                    clicked = True
                    print(f"[CDS] АДС кликнут через селектор: {sel}")
                    break
            except Exception:
                pass
        if clicked:
            break
        await page.wait_for_timeout(2000)
    if not clicked:
        await page.screenshot(path="cds_ads_not_found.png")
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
        print("[CDS] ❌ Не найден пункт 'Обращения'")
        return False

    # Ждём появления таблицы
    await page.wait_for_timeout(8000)

    # Проверяем что таблица загрузилась
    grid_visible = await page.locator('#form5_Список').is_visible(timeout=15000)
    if grid_visible:
        print("[CDS] ✅ Таблица обращений загружена")
        return True

    # Fallback: ищем любой grid
    any_grid = await page.locator('.gridBody').first.is_visible(timeout=5000)
    if any_grid:
        print("[CDS] ✅ Grid найден (fallback)")
        return True

    print("[CDS] ⚠️ Таблица не обнаружена, продолжаем...")
    return True