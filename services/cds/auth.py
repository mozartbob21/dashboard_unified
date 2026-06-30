"""
Авторизация в CDS.
Сервер использует NTLM (Windows-аутентификация) + форму 1С.
"""
from services.cds.error_handler import handle_1c_error_dialog
from playwright.async_api import async_playwright

from services.cds.config import (
    CDS_BASE_URL,
    CDS_HTTP_USER,
    CDS_HTTP_PASSWORD,
    CDS_1C_USER,
    CDS_1C_PASSWORD,
    SCREENSHOT_DIR,
)


async def login_to_cds(headless: bool = True):
    """
    Запускает браузер и логинится в 1С CDS.
    Возвращает (playwright, browser, context, page).
    """
    pw = await async_playwright().start()

    browser = await pw.chromium.launch(
        headless=headless,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--auth-server-allowlist=176.100.216.181",
            "--auth-negotiate-delegate-allowlist=176.100.216.181",
        ],
    )

    # Контекст с HTTP credentials для NTLM
    context = await browser.new_context(
        viewport={"width": 1400, "height": 900},
        locale="ru-RU",
        http_credentials={
            "username": CDS_HTTP_USER,
            "password": CDS_HTTP_PASSWORD,
        },
        ignore_https_errors=True,
    )

    page = await context.new_page()

    print(f"[CDS] Переходим на {CDS_BASE_URL}...")
    print(f"[CDS] HTTP User: {CDS_HTTP_USER}")

    # Подключение с ретраями: сеть/VPN могут моргать, даём до 4 попыток
    _last_err = None
    for _attempt, _pause in [(1, 5), (2, 15), (3, 30), (4, 0)]:
        try:
            await page.goto(CDS_BASE_URL, wait_until="domcontentloaded", timeout=90000)
            if _attempt > 1:
                print(f"[CDS] ✅ Подключились с попытки {_attempt}")
            _last_err = None
            break
        except Exception as e:
            _last_err = e
            print(f"[CDS] Попытка {_attempt}/4 не удалась: {type(e).__name__}: {e}")
            if _pause:
                print(f"[CDS] Пауза {_pause} сек перед повтором...")
                await page.wait_for_timeout(_pause * 1000)
    if _last_err is not None:
        raise RuntimeError(
            f"Сервер 1С недоступен после 4 попыток (проверьте сеть/VPN): {_last_err}"
        ) from _last_err

    await page.wait_for_timeout(5000)
    try:
        await page.screenshot(
            path=str(SCREENSHOT_DIR / "after_goto.png"),
            timeout=10000,
        )
    except Exception as e:
        print(f"[CDS] ⚠️ Скриншот after_goto не удался (не критично): {type(e).__name__}")
    current_url = page.url
    print(f"[CDS] URL после goto: {current_url}")

    # ── Детекция формы логина 1С ──
    # 1С рисует форму через JS уже после domcontentloaded, поэтому ждём.
    await page.wait_for_timeout(3000)

    async def _detect_login_form():
        try:
            pwd = page.locator(
                'input[type="password"], [name="UserPassword"], #passText, '
                'input[name="pwd"]'
            )
            if await pwd.count() > 0:
                for i in range(await pwd.count()):
                    if await pwd.nth(i).is_visible():
                        return True
        except Exception:
            pass

        try:
            login_btn = page.get_by_text("Войти", exact=False)
            txt_inputs = page.locator('input:visible')
            if await login_btn.count() > 0 and await txt_inputs.count() > 0:
                for i in range(await login_btn.count()):
                    if await login_btn.nth(i).is_visible():
                        return True
        except Exception:
            pass

        return False

    has_login_form = False
    for _ in range(10):
        has_login_form = await _detect_login_form()
        if has_login_form:
            break
        await page.wait_for_timeout(2000)

    if has_login_form:
        print("[CDS] Найдена форма логина 1С, вводим credentials...")

        login_filled = False
        login_selectors = [
            'input[name="UserName"]',
            'input[name="usr"]',
            '#loginText',
            'input[type="text"]:visible',
        ]
        for sel in login_selectors:
            try:
                login_input = page.locator(sel).first
                if await login_input.is_visible(timeout=2000):
                    await login_input.click()
                    await login_input.fill(CDS_1C_USER)
                    val = await login_input.input_value()
                    print(f"[CDS] Логин 1С введён в {sel} -> [{val}]")
                    login_filled = True
                    break
            except Exception:
                continue
        if not login_filled:
            print("[CDS] \u26a0\ufe0f Поле логина не найдено")

        pwd_filled = False
        pwd_selectors = [
            'input[name="UserPassword"]',
            'input[name="pwd"]',
            '#passText',
            'input[type="password"]:visible',
        ]
        for sel in pwd_selectors:
            try:
                pwd_input = page.locator(sel).first
                if await pwd_input.is_visible(timeout=2000):
                    await pwd_input.click()
                    await pwd_input.fill(CDS_1C_PASSWORD)
                    val = await pwd_input.input_value()
                    print(f"[CDS] Пароль 1С введён в {sel} -> длина {len(val)}")
                    pwd_filled = True
                    break
            except Exception:
                continue
        if not pwd_filled:
            print("[CDS] \u26a0\ufe0f Поле пароля не найдено")

        clicked = False
        submit_selectors = [
            'input[type="submit"]',
            'button[type="submit"]',
            '#buttonOK',
            'text="Войти"',
            'text="OK"',
        ]
        for sel in submit_selectors:
            try:
                btn = page.locator(sel).first
                if await btn.is_visible(timeout=1500):
                    await btn.click()
                    print(f"[CDS] Кнопка входа: {sel}")
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            await page.keyboard.press("Enter")
            print("[CDS] Enter для входа")

        await page.wait_for_timeout(8000)

        still_login = await _detect_login_form()
        if still_login:
            shot = str(SCREENSHOT_DIR / "cds_login_failed.png")
            try:
                await page.screenshot(path=shot, timeout=10000)
            except Exception:
                pass
            raise RuntimeError(
                f"Форма логина 1С всё ещё видна после входа — "
                f"проверь логин/пароль 1С (скриншот: {shot})"
            )
        print("[CDS] \u2705 Форма логина пройдена")
    else:
        print("[CDS] Форма логина 1С не найдена — NTLM прошёл, уже внутри")
        await page.wait_for_timeout(3000)

    # Даём 1С дорисоваться после входа

    await page.wait_for_timeout(5000)

    # Страховка: если выскочило окно непредвиденной ошибки 1С — обработаем
    await handle_1c_error_dialog(page)

    # Отладочный скриншот — НЕ должен ронять логин при таймауте

    try:

        await page.screenshot(

            path=str(SCREENSHOT_DIR / "after_login.png"),

            timeout=10000,

        )

    except Exception as e:

        print(f"[CDS] ⚠️ Скриншот after_login не удался (не критично): {type(e).__name__}")
    print(f"[CDS] ✅ Авторизация завершена. URL: {page.url}")

    return pw, browser, context, page