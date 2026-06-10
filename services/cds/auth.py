"""
Авторизация в CDS.
Сервер использует NTLM (Windows-аутентификация) + форму 1С.
"""
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

    # Пробуем с domcontentloaded (быстрее чем networkidle)
    try:
        await page.goto(CDS_BASE_URL, wait_until="domcontentloaded", timeout=90000)
    except Exception as e:
        print(f"[CDS] Первая попытка: {e}")
        await page.wait_for_timeout(3000)
        await page.goto(CDS_BASE_URL, wait_until="domcontentloaded", timeout=90000)

    await page.wait_for_timeout(5000)
    await page.screenshot(path=str(SCREENSHOT_DIR / "after_goto.png"))

    current_url = page.url
    print(f"[CDS] URL после goto: {current_url}")

    # Проверяем — есть ли форма логина 1С (после прохождения NTLM)
    has_login_form = await page.evaluate("""() => {
        return !!document.querySelector('input[type="password"]') ||
               !!document.querySelector('[name="UserPassword"]') ||
               !!document.querySelector('#passText');
    }""")

    if has_login_form:
        print("[CDS] Найдена форма логина 1С, вводим credentials...")

        # Ищем поле логина
        login_selectors = [
            'input[name="UserName"]',
            'input[name="usr"]',
            '#loginText',
            'input[type="text"]',
        ]
        for sel in login_selectors:
            login_input = page.locator(sel).first
            try:
                if await login_input.is_visible(timeout=2000):
                    await login_input.fill(CDS_1C_USER)
                    print(f"[CDS] Логин 1С введён в {sel}")
                    break
            except Exception:
                continue

        # Ищем поле пароля
        pwd_selectors = [
            'input[name="UserPassword"]',
            'input[name="pwd"]',
            '#passText',
            'input[type="password"]',
        ]
        for sel in pwd_selectors:
            pwd_input = page.locator(sel).first
            try:
                if await pwd_input.is_visible(timeout=2000):
                    await pwd_input.fill(CDS_1C_PASSWORD)
                    print(f"[CDS] Пароль 1С введён в {sel}")
                    break
            except Exception:
                continue

        # Нажимаем кнопку входа
        submit_selectors = [
            'input[type="submit"]',
            'button[type="submit"]',
            '#buttonOK',
            'text="Войти"',
            'text="OK"',
        ]
        for sel in submit_selectors:
            btn = page.locator(sel).first
            try:
                if await btn.is_visible(timeout=1000):
                    await btn.click()
                    print(f"[CDS] Кнопка входа: {sel}")
                    break
            except Exception:
                continue
        else:
            await page.keyboard.press("Enter")
            print("[CDS] Enter для входа")

        await page.wait_for_timeout(8000)
    else:
        print("[CDS] Форма логина 1С не найдена — NTLM прошёл, уже внутри")
        await page.wait_for_timeout(3000)

    await page.screenshot(path=str(SCREENSHOT_DIR / "after_login.png"))
    print(f"[CDS] ✅ Авторизация завершена. URL: {page.url}")

    return pw, browser, context, page