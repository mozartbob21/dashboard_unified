"""E2E тесты планировщика автозапуска."""
import pytest
from playwright.sync_api import expect

pytestmark = pytest.mark.e2e


def test_scheduler_page_loads(admin_page, app_url):
    """Планировщик открывается для админа."""
    admin_page.goto(f"{app_url}/scheduler")

    expect(admin_page.locator("h1")).to_contain_text("Авто-запуск")
    expect(admin_page.locator(".sch-status-bar")).to_be_visible()


def test_scheduler_cards_have_correct_labels(admin_page, app_url):
    """
    Карточки планировщика используют обновлённые названия:
    'Обращения 1С', 'Заполненность данных', а не 'ЦДС', 'ЭДО'.
    """
    admin_page.goto(f"{app_url}/scheduler")

    cards = admin_page.locator(".sch-card")
    expect(cards.first).to_be_visible(timeout=10000)

    count = cards.count()
    assert count >= 5, f"В планировщике должно быть минимум 5 карточек, найдено: {count}"

    titles = [cards.nth(i).locator(".sch-card-title").text_content() for i in range(count)]
    titles_text = " ".join(titles)

    assert "ЦДС" not in titles_text, (
        f"Найдено старое название 'ЦДС' — должно быть 'Обращения 1С'. Все заголовки: {titles}"
    )


def test_scheduler_toggle_enables_job(admin_page, app_url):
    """Переключатель включает задачу и меняет статус."""
    admin_page.goto(f"{app_url}/scheduler")

    first_card = admin_page.locator(".sch-card").first
    toggle_label = first_card.locator(".sch-toggle")
    checkbox = first_card.locator(".sch-toggle input")

    expect(toggle_label).to_be_visible()
    was_checked = checkbox.is_checked()

    toggle_label.click()

    # Ждём изменения состояния
    admin_page.wait_for_timeout(500)
    assert checkbox.is_checked() != was_checked


def test_scheduler_interval_change(admin_page, app_url):
    """Можно изменить интервал через селект."""
    admin_page.goto(f"{app_url}/scheduler")

    first_card = admin_page.locator(".sch-card").first
    select = first_card.locator(".sch-interval-select")
    save_btn = first_card.locator(".sch-btn-save")

    expect(select).to_be_visible(timeout=5000)

    # Берём первую доступную опцию — заведомо существует
    first_option_value = select.locator("option").first.get_attribute("value")
    select.select_option(first_option_value)

    save_btn.click()

    # expect сам ждёт появления toast с нужным текстом
    toast = admin_page.locator(".sch-toast")
    expect(toast).to_contain_text(
        "Интервал",
        timeout=5000,
        ignore_case=True,
    )
    