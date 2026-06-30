"""Обработка окна непредвиденной ошибки 1С."""


async def handle_1c_error_dialog(page, max_retries: int = 2) -> bool:
    """
    Проверяет наличие окна 'К сожалению, возникла непредвиденная ошибка'.
    Если есть — жмёт 'Перезапустить'. Возвращает True если окно было обработано.
    """
    for attempt in range(max_retries):
        # Ищем текст ошибки на странице
        error_visible = await page.locator(
            "text=возникла непредвиденная ошибка"
        ).count()

        if not error_visible:
            return attempt > 0  # True если хоть раз чинили

        print(f"[CDS] ⚠️ Обнаружено окно ошибки 1С (попытка {attempt + 1}/{max_retries})")

        # Жмём 'Перезапустить'
        restart_btn = page.locator("text=Перезапустить").first
        if await restart_btn.count():
            await restart_btn.click()
            print("[CDS] Нажата кнопка 'Перезапустить', ждём перезагрузку...")
            await page.wait_for_timeout(5000)
            try:
                await page.wait_for_load_state("networkidle", timeout=30000)
            except Exception:
                pass
        else:
            print("[CDS] ❌ Кнопка 'Перезапустить' не найдена")
            return False

    # После всех попыток снова проверяем
    still_error = await page.locator(
        "text=возникла непредвиденная ошибка"
    ).count()
    if still_error:
        raise RuntimeError(
            "1С возвращает непредвиденную ошибку даже после перезапусков. "
            "Вероятно, сервер 1С на обслуживании или есть баг конфигурации. "
            "Попробуйте позже."
        )
    return True
