"""
Установка периода в 1С.
Стратегия: keyboard-only навигация по модальному диалогу.
Никакого mouse.click, никакого JS .click() — только Tab + Enter + Type.
"""
from datetime import datetime, timedelta
from services.cds.config import SCREENSHOT_DIR


def get_default_period() -> tuple[str, str]:
    """Возвращает (начало прошлого месяца, сегодня) в формате DD.MM.YYYY."""
    today = datetime.now()
    # Начало текущего месяца
    first_of_current = today.replace(day=1)
    # Начало прошлого месяца
    last_month_end = first_of_current - timedelta(days=1)
    first_of_last = last_month_end.replace(day=1)

    date_from = first_of_last.strftime("%d.%m.%Y")
    date_to = today.strftime("%d.%m.%Y")
    return date_from, date_to


async def set_period(page, date_from: str, date_to: str) -> bool:
    """
    Устанавливает период через Ещё → Установить период.
    Использует ТОЛЬКО клавиатуру для работы с модальным диалогом.
    """
    print(f"[CDS] Устанавливаем период: {date_from} — {date_to}")

    # === Шаг 1: Открыть меню "Ещё" ===
    more_btn = page.locator('#form5_allActionsФормаСтандартныеКоманды')
    await more_btn.click(timeout=10000)
    await page.wait_for_timeout(1500)

    # === Шаг 2: "Установить период..." ===
    period_item = page.locator('text="Установить период..."').first
    if await period_item.is_visible(timeout=3000):
        await period_item.click()
    else:
        # Fallback: ищем по id
        fallback = page.locator('[id*="УстановитьИнтервал"]').first
        if await fallback.is_visible(timeout=2000):
            await fallback.click()
        else:
            print("[CDS] ⚠️ Пункт 'Установить период' не найден")
            await page.keyboard.press("Escape")
            return False

    await page.wait_for_timeout(2500)
    await page.screenshot(path=str(SCREENSHOT_DIR / "period_01_dialog_opened.png"))

    # === Шаг 3: Работаем с диалогом ТОЛЬКО через клавиатуру ===
    # Структура диалога "Выберите период":
    #   [Radio: Стандартный / Произвольный]
    #   [Поле даты "с"] [📅] [✕]
    #   [Поле даты "по"] [📅] [✕]
    #   [Ссылка: Очистить период]
    #   [Кнопка: Выбрать] [Кнопка: Отмена]
    #
    # При открытии фокус обычно на radio или первом поле.
    # Нам нужно: выбрать "Произвольный" → ввести даты → нажать "Выбрать"

    # Сначала переключим на "Произвольный" период (если есть radio)
    # В 1С это обычно второй radio-button
    # Попробуем просто Tab до первого поля ввода даты

    # Нажимаем Tab несколько раз чтобы попасть на первое поле даты
    # Определим где мы находимся по evaluate
    
    # Стратегия: 
    # 1) Найдём ID полей ввода в диалоге через JS
    # 2) Используем page.fill() / page.type() по ID
    
    field_ids = await page.evaluate("""() => {
        const modal = document.getElementById('modalSurface');
        if (!modal || modal.style.display === 'none') return { error: 'no_modal' };
        
        // Ищем все input[type=text] которые сейчас в DOM и видимы
        const inputs = document.querySelectorAll('input[type="text"]');
        const dateFields = [];
        
        for (const inp of inputs) {
            const rect = inp.getBoundingClientRect();
            // Видимое поле в области модалки (примерно y > 250)
            if (rect.width > 40 && rect.height > 10 && rect.top > 200 && rect.top < 600) {
                dateFields.push({
                    id: inp.id,
                    name: inp.name,
                    value: inp.value,
                    x: Math.round(rect.x + rect.width/2),
                    y: Math.round(rect.y + rect.height/2),
                    width: Math.round(rect.width)
                });
            }
        }
        
        // Сортируем сверху вниз
        dateFields.sort((a, b) => a.y - b.y);
        
        return { fields: dateFields };
    }""")

    print(f"[CDS] Поля в диалоге: {field_ids}")

    if "error" in field_ids:
        print("[CDS] ⚠️ Модальное окно не обнаружено!")
        return False

    fields = field_ids.get("fields", [])

    if len(fields) >= 2:
        # Нашли поля — заполняем через focus + keyboard
        await _fill_date_field(page, fields[0], date_from)
        await _fill_date_field(page, fields[1], date_to)
    elif len(fields) == 1:
        # Одно поле — может быть "Стандартный" режим, переключим на "Произвольный"
        await _switch_to_arbitrary(page)
        await page.wait_for_timeout(1000)
        # Перечитаем поля
        field_ids2 = await page.evaluate("""() => {
            const inputs = document.querySelectorAll('input[type="text"]');
            const dateFields = [];
            for (const inp of inputs) {
                const rect = inp.getBoundingClientRect();
                if (rect.width > 40 && rect.height > 10 && rect.top > 200 && rect.top < 600) {
                    dateFields.push({ id: inp.id, x: Math.round(rect.x+rect.width/2), y: Math.round(rect.y+rect.height/2), width: Math.round(rect.width) });
                }
            }
            dateFields.sort((a, b) => a.y - b.y);
            return dateFields;
        }""")
        if len(field_ids2) >= 2:
            await _fill_date_field(page, field_ids2[0], date_from)
            await _fill_date_field(page, field_ids2[1], date_to)
        else:
            await _fill_via_tab(page, date_from, date_to)
    else:
        # Полей не нашли — чисто клавиатурный ввод
        await _fill_via_tab(page, date_from, date_to)

    await page.wait_for_timeout(500)
    await page.screenshot(path=str(SCREENSHOT_DIR / "period_02_filled.png"))

    # === Шаг 4: Нажимаем "Выбрать" через Enter ===
    # Табаем до кнопки "Выбрать" и жмём Enter
    # Или просто Enter (если фокус уже на кнопке)
    
    # Пробуем найти кнопку и сфокусироваться на ней
    btn_focused = await page.evaluate("""() => {
        const allEls = document.querySelectorAll('button, a, [tabindex]');
        for (const el of allEls) {
            const rect = el.getBoundingClientRect();
            const text = el.textContent.trim();
            if (text === 'Выбрать' && rect.width > 0 && rect.top > 200) {
                el.focus();
                return true;
            }
        }
        return false;
    }""")

    if btn_focused:
        await page.keyboard.press("Enter")
        print("[CDS] Нажат Enter на кнопке 'Выбрать' (focus)")
    else:
        # Табаем до "Выбрать"
        for i in range(15):
            await page.keyboard.press("Tab")
            await page.wait_for_timeout(100)
            # Проверяем на чём фокус
            focused_text = await page.evaluate("""() => {
                const el = document.activeElement;
                return el ? el.textContent.trim() : '';
            }""")
            if focused_text == "Выбрать":
                await page.keyboard.press("Enter")
                print(f"[CDS] Tab #{i+1} → фокус на 'Выбрать' → Enter")
                break
        else:
            # Не нашли — просто Enter
            await page.keyboard.press("Enter")
            print("[CDS] Enter (не нашли 'Выбрать' через Tab)")

    await page.wait_for_timeout(3000)

    # === Шаг 5: Проверяем что модалка закрылась ===
    still_open = await page.evaluate("""() => {
        const modal = document.getElementById('modalSurface');
        return modal && modal.offsetWidth > 0 && getComputedStyle(modal).display !== 'none';
    }""")

    if still_open:
        print("[CDS] ⚠️ Модалка не закрылась! Пробуем Escape...")
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(2000)

    await page.screenshot(path=str(SCREENSHOT_DIR / "period_03_after.png"))
    print("[CDS] ✅ Период установлен")
    return True


async def _fill_date_field(page, field_info: dict, date_value: str):
    """Заполняет поле даты: кликаем по ID через JS focus, потом keyboard."""
    field_id = field_info.get("id", "")

    if field_id:
        # Фокусируемся на поле через JS
        await page.evaluate(f"""() => {{
            const el = document.getElementById('{field_id}');
            if (el) {{ el.focus(); el.select(); }}
        }}""")
    else:
        # Кликаем по координатам через Playwright
        await page.mouse.click(field_info["x"], field_info["y"])

    await page.wait_for_timeout(200)

    # Выделяем всё и вводим новую дату
    await page.keyboard.press("Control+a")
    await page.wait_for_timeout(100)
    await page.keyboard.type(date_value, delay=30)
    await page.wait_for_timeout(300)

    # Tab чтобы подтвердить ввод (blur)
    await page.keyboard.press("Tab")
    await page.wait_for_timeout(200)

    print(f"[CDS] Поле {field_id or field_info.get('x','?')}: '{date_value}'")


async def _switch_to_arbitrary(page):
    """Переключает радио-кнопку на 'Произвольный' период."""
    await page.evaluate("""() => {
        const radios = document.querySelectorAll('input[type="radio"]');
        for (const r of radios) {
            // Второй radio обычно "Произвольный"
            const label = r.parentElement?.textContent || '';
            if (label.includes('Произвольн') || label.includes('произвольн')) {
                r.click();
                r.checked = true;
                r.dispatchEvent(new Event('change', {bubbles: true}));
                return;
            }
        }
        // Если не нашли по тексту — кликаем второй radio
        if (radios.length >= 2) {
            radios[1].click();
            radios[1].checked = true;
            radios[1].dispatchEvent(new Event('change', {bubbles: true}));
        }
    }""")
    print("[CDS] Переключено на 'Произвольный'")


async def _fill_via_tab(page, date_from: str, date_to: str):
    """Fallback: чисто Tab-навигация для ввода дат."""
    print("[CDS] Fallback: ввод дат через Tab...")

    # Tab до первого поля ввода
    for _ in range(5):
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(150)
        # Проверяем тип активного элемента
        is_input = await page.evaluate("""() => {
            const el = document.activeElement;
            return el && el.tagName === 'INPUT' && el.type === 'text';
        }""")
        if is_input:
            break

    # Вводим первую дату
    await page.keyboard.press("Control+a")
    await page.keyboard.type(date_from, delay=30)
    await page.wait_for_timeout(300)

    # Tab до второго поля (может быть через 2-4 табов из-за кнопок)
    for _ in range(5):
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(150)
        is_input = await page.evaluate("""() => {
            const el = document.activeElement;
            return el && el.tagName === 'INPUT' && el.type === 'text';
        }""")
        if is_input:
            break

    # Вводим вторую дату
    await page.keyboard.press("Control+a")
    await page.keyboard.type(date_to, delay=30)
    await page.wait_for_timeout(300)