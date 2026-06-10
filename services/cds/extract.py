"""
Извлечение данных: Ещё → Вывести список → Парсинг DOM табличного документа.
"""
from services.cds.config import SCREENSHOT_DIR


async def extract_via_output_list(page) -> list[dict]:
    """
    Выполняет 'Вывести список...' и парсит результат из DOM.
    """
    print("[CDS] Вывести список...")

    # === Убираем модалку если есть ===
    await _dismiss_modal(page)

    # === Открываем меню "Ещё" ===
    more_btn = page.locator('#form5_allActionsФормаСтандартныеКоманды')
    try:
        await more_btn.click(timeout=10000)
    except Exception as e:
        print(f"[CDS] Первый клик 'Ещё' не удался: {e}")
        await _dismiss_modal(page)
        await more_btn.click(timeout=10000)

    await page.wait_for_timeout(1500)

    # === Клик "Вывести список..." ===
    output_btn = page.locator('text="Вывести список..."').first
    if not await output_btn.is_visible(timeout=3000):
        output_btn = page.locator('#form5_popup_ФормаВывестиСписок')

    if not await output_btn.is_visible(timeout=3000):
        print("[CDS] ⚠️ Пункт 'Вывести список...' не найден")
        await page.keyboard.press("Escape")
        return []

    await output_btn.click()
    await page.wait_for_timeout(2000)

    # Скриншот
    await page.screenshot(path=str(SCREENSHOT_DIR / "output_list_dialog.png"))

    # === Нажимаем ОК через mouse.click по координатам ===
    ok_clicked = await _click_button_by_text(page, "ОК")
    if not ok_clicked:
        print("[CDS] ОК не нажата координатами, пробуем Enter...")
        await page.keyboard.press("Enter")

    await page.wait_for_timeout(3000)

    # Проверяем закрылся ли диалог
    still_open = await _is_modal_open(page)
    if still_open:
        print("[CDS] Диалог не закрылся после ОК, пробуем Enter...")
        await page.keyboard.press("Enter")
        await page.wait_for_timeout(3000)

    still_open = await _is_modal_open(page)
    if still_open:
        print("[CDS] Диалог ВСЕГДА НЕ закрывается, пробуем Tab+Enter...")
        # Фокус может быть не на кнопке — табаем до ОК
        for _ in range(10):
            await page.keyboard.press("Tab")
            await page.wait_for_timeout(100)
        await page.keyboard.press("Enter")
        await page.wait_for_timeout(3000)

    # Ждём формирования
    print("[CDS] Ожидаем формирования табличного документа...")
    await page.wait_for_timeout(12000)

    await page.screenshot(path=str(SCREENSHOT_DIR / "output_list_result.png"))

    # === Парсим ===
    rows = await _parse_spreadsheet_document(page)

    if not rows:
        print("[CDS] Табличный документ не распарсился, fallback: grid...")
        rows = await _parse_grid_direct(page)

    return rows


async def _click_button_by_text(page, text: str) -> bool:
    """
    Кликает по кнопке 1С через page.mouse.click() по координатам.
    1С веб-клиент НЕ реагирует на el.click() из JS!
    """
    # Получаем координаты элемента с нужным текстом
    coords = await page.evaluate("""(targetText) => {
        const allElements = document.querySelectorAll('button, a, span, div, td');
        const candidates = [];
        
        for (const el of allElements) {
            // Точное совпадение текста (без дочерних)
            const directText = Array.from(el.childNodes)
                .filter(n => n.nodeType === 3)
                .map(n => n.textContent.trim())
                .join('');
            
            const fullText = el.textContent.trim();
            const rect = el.getBoundingClientRect();
            
            if (rect.width === 0 || rect.height === 0) continue;
            if (rect.width > 200 || rect.height > 60) continue;
            
            // Проверяем и прямой текст, и полный
            if (directText === targetText || (fullText === targetText && el.children.length === 0)) {
                candidates.push({
                    x: rect.x + rect.width / 2,
                    y: rect.y + rect.height / 2,
                    w: rect.width,
                    h: rect.height,
                    tag: el.tagName,
                    id: el.id,
                    cls: el.className.substring(0, 50)
                });
            }
        }
        
        if (candidates.length === 0) return null;
        
        // Предпочитаем button/a с меньшей шириной (более конкретный элемент)
        candidates.sort((a, b) => {
            if (a.tag === 'BUTTON' && b.tag !== 'BUTTON') return -1;
            if (a.tag === 'A' && b.tag !== 'A' && b.tag !== 'BUTTON') return -1;
            return a.w - b.w;
        });
        
        return candidates[0];
    }""", text)

    if not coords:
        print(f"[CDS] Кнопка '{text}' не найдена в DOM")
        return False

    print(f"[CDS] Кликаем '{text}' по координатам ({coords['x']:.0f}, {coords['y']:.0f}) [{coords['tag']} id={coords['id']}]")

    # Реальный клик мышью через Playwright
    await page.mouse.click(coords["x"], coords["y"])
    await page.wait_for_timeout(1000)

    return True


async def _is_modal_open(page) -> bool:
    """Проверяет открыто ли модальное окно."""
    try:
        result = await page.evaluate("""() => {
            const modal = document.getElementById('modalSurface');
            if (!modal) return false;
            const style = getComputedStyle(modal);
            return style.display !== 'none' && modal.offsetWidth > 0;
        }""")
        return result
    except Exception:
        return False


async def _dismiss_modal(page):
    """Закрывает модальное окно если оно есть."""
    if await _is_modal_open(page):
        # Сначала Escape
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(1500)

        if await _is_modal_open(page):
            # Пробуем кликнуть "Отмена"
            await _click_button_by_text(page, "Отмена")
            await page.wait_for_timeout(1500)

        if await _is_modal_open(page):
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(1000)


async def _parse_spreadsheet_document(page) -> list[dict]:
    """Парсит табличный документ 1С."""
    # Проверяем что модалка закрыта
    if await _is_modal_open(page):
        print("[CDS] ⚠️ Модалка ещё открыта при парсинге!")
        return []

    result = await page.evaluate("""() => {
        // === Вариант 1: Новая форма (form6, form7...) с табличным документом ===
        const forms = document.querySelectorAll('[id^="form"]');
        let newestForm = null;
        let maxFormNum = 5;
        for (const f of forms) {
            const match = f.id.match(/form(\\d+)/);
            if (match) {
                const num = parseInt(match[1]);
                if (num > maxFormNum) {
                    maxFormNum = num;
                    newestForm = f;
                }
            }
        }

        if (newestForm) {
            // Moxel cells
            const moxelCells = newestForm.querySelectorAll('td');
            if (moxelCells.length > 10) {
                const data = [];
                for (const c of moxelCells) {
                    const r = c.getBoundingClientRect();
                    const text = c.textContent.trim();
                    if (r.width > 5 && r.height > 5 && text) {
                        data.push({ top: Math.round(r.top), left: Math.round(r.left), text });
                    }
                }
                if (data.length > 5) {
                    const rowsMap = {};
                    for (const cell of data) {
                        let key = null;
                        for (const k of Object.keys(rowsMap)) {
                            if (Math.abs(Number(k) - cell.top) <= 3) { key = k; break; }
                        }
                        if (!key) { key = String(cell.top); rowsMap[key] = []; }
                        rowsMap[key].push(cell);
                    }
                    const tops = Object.keys(rowsMap).map(Number).sort((a, b) => a - b);
                    const rows = tops.map(t =>
                        rowsMap[String(t)].sort((a, b) => a.left - b.left).map(c => c.text)
                    );
                    if (rows.length > 2) return { source: 'new_form_moxel', rows };
                }
            }

            // Grid в новой форме
            const headCells = newestForm.querySelectorAll('.gridHead .gridBoxText');
            const headers = [];
            headCells.forEach(c => headers.push(c.textContent.trim()));

            const lines = newestForm.querySelectorAll('.gridBody .gridLine');
            const rows = [];
            if (headers.length > 0) rows.push(headers);
            for (const line of lines) {
                const cells = [];
                line.querySelectorAll('.gridBoxText').forEach(c => cells.push(c.textContent.trim()));
                if (cells.length > 0 && cells.some(c => c)) rows.push(cells);
            }
            if (rows.length > 2) return { source: 'new_form_grid', rows };
        }

        // === Вариант 2: Основной grid form5 ===
        const mainHead = document.querySelectorAll('.gridHead .gridBoxText');
        const headers = [];
        mainHead.forEach(c => headers.push(c.textContent.trim()));

        const mainLines = document.querySelectorAll('.gridBody .gridLine');
        const rows = [];
        if (headers.length > 0) rows.push(headers);
        for (const line of mainLines) {
            const cells = [];
            line.querySelectorAll('.gridBoxText').forEach(c => cells.push(c.textContent.trim()));
            if (cells.length > 0 && cells.some(c => c)) rows.push(cells);
        }
        if (rows.length > 2) return { source: 'main_grid', rows };

        return { source: 'none', rows: [] };
    }""")

    if not result["rows"]:
        return []

    print(f"[CDS] Парсинг ({result['source']}): {len(result['rows'])} строк")
    return _rows_to_records(result["rows"])


def _rows_to_records(rows: list[list[str]]) -> list[dict]:
    """Преобразует массив строк в список словарей."""
    if len(rows) < 2:
        return []

    headers = rows[0]
    data = []
    for row in rows[1:]:
        if not any(cell for cell in row):
            continue
        record = {}
        for i, val in enumerate(row):
            key = headers[i] if i < len(headers) else f"col_{i}"
            record[key] = val
        record["date"] = record.get("Дата", "")
        record["number"] = record.get("Номер", "")
        record["status"] = record.get("Состояние обращения", record.get("Состояние", record.get("Состо...", "")))
        record["address"] = record.get("Адрес обращения", record.get("Адрес", ""))
        record["type"] = record.get("Тип обращения", "")
        record["deadline"] = record.get("Срок исполнения", "")
        record["request_type"] = record.get("Тип заявки", "")
        record["source"] = record.get("Источник поступления", record.get("Источник", ""))
        record["applicant"] = record.get("Заявитель", "")
        record["executor"] = record.get("Исполнитель", "")
        record["department"] = record.get("Подразделение", record.get("По...", ""))
        data.append(record)

    return data


async def _parse_grid_direct(page) -> list[dict]:
    """Fallback: парсит текущую таблицу + пагинация."""
    all_rows = []
    seen_keys = set()
    max_pages = 50

    for page_num in range(max_pages):
        rows = await page.evaluate("""() => {
            const headers = [];
            document.querySelectorAll('.gridHead .gridBoxText').forEach(c => {
                headers.push(c.textContent.trim());
            });
            const dataRows = [];
            document.querySelectorAll('.gridBody .gridLine').forEach(line => {
                const cells = [];
                line.querySelectorAll('.gridBoxText').forEach(c => cells.push(c.textContent.trim()));
                if (cells.length > 0) dataRows.push(cells);
            });
            return { headers, dataRows };
        }""")

        headers = rows["headers"]
        new_count = 0

        for row in rows["dataRows"]:
            key = "|".join(row[:5])
            if key in seen_keys:
                continue
            seen_keys.add(key)
            new_count += 1

            record = {}
            for i, val in enumerate(row):
                h = headers[i] if i < len(headers) else f"col_{i}"
                record[h] = val
            record["date"] = record.get("Дата", "")
            record["number"] = record.get("Номер", "")
            record["status"] = record.get("Состояние обращения", record.get("Состояние", ""))
            record["address"] = record.get("Адрес обращения", record.get("Адрес", ""))
            record["type"] = record.get("Тип обращения", "")
            record["deadline"] = record.get("Срок исполнения", "")
            record["request_type"] = record.get("Тип заявки", "")
            record["source"] = record.get("Источник поступления", "")
            record["applicant"] = record.get("Заявитель", "")
            record["executor"] = record.get("Исполнитель", "")
            record["department"] = record.get("Подразделение", "")
            all_rows.append(record)

        print(f"  Стр. {page_num + 1}: +{new_count} новых, всего {len(all_rows)}")

        if new_count == 0 and page_num > 0:
            break

        has_next = await page.evaluate("""() => {
            const btns = document.querySelectorAll('[id*="Down"], [id*="Next"], [id*="forward"]');
            for (const btn of btns) {
                if (btn.offsetWidth > 0 && !btn.classList.contains('disabled')) {
                    btn.click();
                    return true;
                }
            }
            return false;
        }""")

        if not has_next:
            break

        await page.wait_for_timeout(3000)

    return all_rows