"""
CDS Export: сохранение списка обращений в Excel через интерфейс 1С.

Последовательность:
1. На странице "Список" (обращения уже открыты)
2. Клик по "⋮" (Ещё / три точки) в верхней панели инструментов
3. Файл → Сохранить как...
4. В окне "Сохранить" меняем тип на "Лист Excel (*.xls)"
5. Задаём имя файла
6. Нажимаем OK
7. Ловим скачанный файл
8. Парсим .xls → JSON → сохраняем в data/cds/result.json
"""

import json
import traceback
from pathlib import Path
from datetime import datetime

RESULT_DIR = Path("data/cds")
DOWNLOADS_DIR = RESULT_DIR / "downloads"


async def set_period_and_export(page, date_from: str = "", date_to: str = "") -> bool:
    """
    Экспорт списка обращений в Excel.
    page — уже на странице со списком обращений.
    Возвращает True если файл успешно скачан и распарсен.
    """
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)

    try:
        # --- ШАГ 1: Установка периода (если поля есть на странице) ---
        if date_from and date_to:
            await _try_set_period(page, date_from, date_to)

        # --- ШАГ 2: Нажимаем "Ещё" (три точки) в панели инструментов ---
        print("STAGE: Открываем меню экспорта")

        # Ищем кнопку "Ещё" — это может быть кнопка с текстом "Ещё" или иконка "⋮"
        more_button = None

        # Вариант 1: кнопка с текстом "Ещё"
        more_button = page.locator('button:has-text("Ещё")').first
        if not await more_button.is_visible(timeout=3000):
            # Вариант 2: три вертикальные точки в тулбаре
            more_button = page.locator('[title="Ещё"]').first
        if not await more_button.is_visible(timeout=3000):
            # Вариант 3: иконка меню "⋮" в правом верхнем углу
            more_button = page.locator('.more-actions-button, .toolbar-more, [data-action="more"]').first
        if not await more_button.is_visible(timeout=3000):
            # Вариант 4: последняя кнопка в тулбаре
            more_button = page.locator('.toolbar button').last

        await more_button.click()
        await page.wait_for_timeout(1000)

        # --- ШАГ 3: Файл → Сохранить как... ---
        print("STAGE: Файл → Сохранить как...")

        # Ищем пункт меню "Файл"
        file_menu = page.locator('text="Файл"').first
        if await file_menu.is_visible(timeout=3000):
            await file_menu.click()
            await page.wait_for_timeout(500)

        # Ищем "Сохранить как..."
        save_as = page.locator('text="Сохранить как..."').first
        if not await save_as.is_visible(timeout=3000):
            save_as = page.locator('text="Сохранить как"').first

        await save_as.click()
        await page.wait_for_timeout(2000)

        # --- ШАГ 4: В окне "Сохранить" выбираем формат Excel ---
        print("STAGE: Выбираем формат Лист Excel (*.xls)")

        # Окно "Сохранить" открылось. Ищем выпадающий список "Тип:"
        type_dropdown = page.locator('select').first
        if await type_dropdown.is_visible(timeout=3000):
            # Это обычный <select>
            await type_dropdown.select_option(label="Лист Excel (*.xls)")
        else:
            # 1С может использовать кастомный dropdown
            # Кликаем по полю с типом файла
            type_field = page.locator('[class*="type"], [class*="format"]').first
            if not await type_field.is_visible(timeout=2000):
                # Ищем по тексту текущего значения
                type_field = page.locator('text="Табличный документ (*.mxl)"').first

            if await type_field.is_visible(timeout=2000):
                await type_field.click()
                await page.wait_for_timeout(1000)

                # Выбираем "Лист Excel (*.xls)" из списка
                excel_option = page.locator('text="Лист Excel (*.xls)"').first
                await excel_option.click()
                await page.wait_for_timeout(500)

        # --- ШАГ 5: Задаём имя файла ---
        print("STAGE: Задаём имя файла")

        filename = f"cds_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        # Ищем поле "Имя файла"
        name_input = page.locator('input[type="text"]').first
        if await name_input.is_visible(timeout=2000):
            await name_input.triple_click()  # Выделяем всё
            await name_input.fill(filename)
            await page.wait_for_timeout(300)

        # --- ШАГ 6: Нажимаем OK и ловим скачивание ---
        print("STAGE: Сохраняем файл")

        # Устанавливаем перехват скачивания
        async with page.expect_download(timeout=60000) as download_info:
            # Нажимаем "ОК" или "Сохранить"
            ok_button = page.locator('button:has-text("OK")').first
            if not await ok_button.is_visible(timeout=2000):
                ok_button = page.locator('button:has-text("ОК")').first  # Кириллица
            if not await ok_button.is_visible(timeout=2000):
                ok_button = page.locator('button:has-text("Сохранить")').first

            await ok_button.click()

        download = await download_info.value

        # Сохраняем скачанный файл
        download_path = DOWNLOADS_DIR / (filename + ".xls")
        await download.save_as(str(download_path))
        print(f"STAGE: Файл скачан: {download_path}")

        # --- ШАГ 7: Парсим Excel → JSON ---
        print("STAGE: Парсим Excel в JSON")

        rows = parse_xls_to_rows(download_path)

        # --- ШАГ 8: Сохраняем результат ---
        result = {
            "rows": rows,
            "total": len(rows),
            "exported_at": datetime.now().isoformat(timespec="seconds"),
            "source_file": str(download_path.name),
            "date_from": date_from,
            "date_to": date_to,
        }

        result_file = RESULT_DIR / "result.json"
        result_file.write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        print(f"STAGE: Готово. Записей: {len(rows)}")
        return True

    except Exception as e:
        print(f"[CDS Export] Error: {e}")
        print(traceback.format_exc())

        # Сохраняем скриншот ошибки
        try:
            error_screenshot = RESULT_DIR / "export_error.png"
            await page.screenshot(path=str(error_screenshot))
            print(f"[CDS Export] Screenshot saved: {error_screenshot}")
        except:
            pass

        return False


async def _try_set_period(page, date_from: str, date_to: str):
    """Попытка установить период фильтрации (если на странице есть поля дат)."""
    try:
        # Ищем поля периода на странице списка
        period_from = page.locator('[placeholder*="начал"], [title*="Начало"], [name*="from"], [name*="begin"]').first
        period_to = page.locator('[placeholder*="оконч"], [title*="Окончание"], [name*="to"], [name*="end"]').first

        if await period_from.is_visible(timeout=2000):
            await period_from.triple_click()
            await period_from.fill(date_from)
            print(f"  Период с: {date_from}")

        if await period_to.is_visible(timeout=2000):
            await period_to.triple_click()
            await period_to.fill(date_to)
            print(f"  Период по: {date_to}")

        # Нажимаем Enter или кнопку "Обновить" для применения фильтра
        await page.keyboard.press("Enter")
        await page.wait_for_timeout(3000)

    except Exception as e:
        print(f"[CDS] Не удалось установить период: {e}")
        # Не фатально — продолжаем без фильтра


def parse_xls_to_rows(file_path: Path) -> list[dict]:
    """
    Парсим .xls файл в список словарей.
    Пробуем xlrd (для .xls), если не установлен — openpyxl (для .xlsx).
    """
    rows = []

    try:
        import xlrd

        workbook = xlrd.open_workbook(str(file_path))
        sheet = workbook.sheet_by_index(0)

        if sheet.nrows < 2:
            return rows

        # Первая строка — заголовки
        headers = []
        for col_idx in range(sheet.ncols):
            cell_value = str(sheet.cell_value(0, col_idx)).strip()
            headers.append(cell_value)

        # Остальные строки — данные
        for row_idx in range(1, sheet.nrows):
            row_data = {}
            for col_idx in range(sheet.ncols):
                key = headers[col_idx] if col_idx < len(headers) else f"col_{col_idx}"
                value = sheet.cell_value(row_idx, col_idx)

                # xlrd возвращает float для чисел
                if isinstance(value, float) and value == int(value):
                    value = int(value)

                row_data[key] = str(value).strip() if value else ""

            # Добавляем нормализованные ключи для фронтенда
            row_data["date"] = row_data.get("Дата", row_data.get("Есть", ""))
            row_data["number"] = row_data.get("Номер", "")
            row_data["status"] = row_data.get("Состояние обращения", "")
            row_data["address"] = row_data.get("Адрес обращения", "")
            row_data["type"] = row_data.get("Тип обращения", "")
            row_data["deadline"] = row_data.get("Срок исполнения", "")
            row_data["request_type"] = row_data.get("Тип заявки", "")
            row_data["source"] = row_data.get("Источник поступления", "")
            row_data["applicant"] = row_data.get("Заявитель", "")
            row_data["executor"] = row_data.get("Исполнитель", "")
            row_data["department"] = row_data.get("Подразделение", "")
            row_data["is_overdue"] = row_data.get("Немедленно", row_data.get("Немедлен", ""))

            rows.append(row_data)

        print(f"  xlrd: прочитано {len(rows)} строк из {file_path.name}")
        return rows

    except ImportError:
        print("  xlrd не установлен, пробуем openpyxl...")

    try:
        import openpyxl

        workbook = openpyxl.load_workbook(str(file_path), read_only=True, data_only=True)
        sheet = workbook.active

        all_rows = list(sheet.iter_rows(values_only=True))

        if len(all_rows) < 2:
            return rows

        headers = [str(cell or "").strip() for cell in all_rows[0]]

        for row_values in all_rows[1:]:
            row_data = {}
            for col_idx, value in enumerate(row_values):
                key = headers[col_idx] if col_idx < len(headers) else f"col_{col_idx}"
                row_data[key] = str(value).strip() if value else ""

            row_data["date"] = row_data.get("Дата", "")
            row_data["number"] = row_data.get("Номер", "")
            row_data["status"] = row_data.get("Состояние обращения", "")
            row_data["address"] = row_data.get("Адрес обращения", "")
            row_data["type"] = row_data.get("Тип обращения", "")
            row_data["deadline"] = row_data.get("Срок исполнения", "")
            row_data["request_type"] = row_data.get("Тип заявки", "")
            row_data["source"] = row_data.get("Источник поступления", "")
            row_data["applicant"] = row_data.get("Заявитель", "")
            row_data["executor"] = row_data.get("Исполнитель", "")
            row_data["department"] = row_data.get("Подразделение", "")

            rows.append(row_data)

        print(f"  openpyxl: прочитано {len(rows)} строк из {file_path.name}")
        return rows

    except ImportError:
        print("  ОШИБКА: ни xlrd, ни openpyxl не установлены!")
        print("  Выполните: pip install xlrd openpyxl")
        return rows

    except Exception as e:
        print(f"  Ошибка парсинга: {e}")
        return rows


async def test_export(headless=True) -> dict:
    """Тест полного цикла: авторизация → навигация → экспорт."""
    try:
        from services.cds.auth import login_to_cds
        from services.cds.navigate import go_to_obrasheniya

        pw, browser, context, page = await login_to_cds(headless=headless)

        try:
            await page.wait_for_timeout(3000)

            nav_ok = await go_to_obrasheniya(page)
            if not nav_ok:
                return {"success": False, "error": "Навигация не удалась"}

            await page.wait_for_timeout(2000)

            export_ok = await set_period_and_export(page, date_from="01.05.2026", date_to="31.05.2026")

            if export_ok:
                result_file = RESULT_DIR / "result.json"
                if result_file.exists():
                    data = json.loads(result_file.read_text(encoding="utf-8"))
                    return {
                        "success": True,
                        "total_rows": data.get("total", 0),
                        "file": str(result_file),
                    }

            return {"success": False, "error": "Экспорт вернул False"}

        finally:
            await browser.close()
            await pw.stop()

    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "trace": traceback.format_exc(),
        }
