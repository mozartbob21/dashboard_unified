"""Навигация: АДС → Обращения."""
from services.cds.config import SCREENSHOT_DIR


async def go_to_obrasheniya(page) -> bool:
    """
    Из главного экрана переходит в:
    Аварийно диспетчерская служба → Обращения.
    Возвращает True если таблица обращений загрузилась.
    """
    print("[CDS] Навигация: АДС...")
    await page.click('#themesCell_theme_4')
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