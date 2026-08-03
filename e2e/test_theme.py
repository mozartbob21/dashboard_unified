"""Тест переключения тем оформления."""
import pytest

pytestmark = pytest.mark.e2e


def test_theme_toggle_changes_attribute(admin_page, app_url):
    """Переключение темы меняет data-theme на <html>."""
    admin_page.goto(app_url)

    html = admin_page.locator("html")
    initial_theme = html.get_attribute("data-theme")

    # Кликаем по кнопке темы (открывает dropdown)
    admin_page.locator("#themeToggle").click()

    # Выбираем «Тёмная»
    admin_page.locator('[data-theme-mode="dark"]').click()

    # Проверяем, что data-theme изменился
    new_theme = html.get_attribute("data-theme")
    assert new_theme != initial_theme or new_theme == "dark"


def test_theme_persists_after_reload(admin_page, app_url):
    """Выбранная тема сохраняется после перезагрузки."""
    admin_page.goto(app_url)
    admin_page.locator("#themeToggle").click()
    admin_page.locator('[data-theme-mode="dark"]').click()

    # Перезагружаем
    admin_page.reload()
    admin_page.wait_for_load_state("networkidle")

    theme = admin_page.locator("html").get_attribute("data-theme")
    assert theme == "dark"
    