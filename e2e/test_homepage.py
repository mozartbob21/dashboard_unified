"""Тесты главной страницы (home.html)."""
import pytest
import re


pytestmark = pytest.mark.e2e




def test_info_switch_between_cards_resets_previous(admin_page, app_url):
    """
    БАГ-РЕГРЕССИЯ: если открыть описание на одной карточке и,
    не нажимая «Назад», нажать (i) на соседней — первая карточка
    должна вернуться в активное состояние (не блёклая, кликабельная).
    """
    admin_page.goto(app_url)

    cards = admin_page.locator(".module-card")
    expect(cards.nth(1)).to_be_visible()

    first_card = cards.nth(0)
    second_card = cards.nth(1)

    # 1. Открываем описание на первой карточке
    first_card.locator(".info-btn").click()
    expect(first_card).to_have_class(re.compile(r"card-info-open"))
    expect(first_card.locator(".card-info-overlay")).to_be_visible()

    # 2. Не нажимая «Назад», жмём (i) на второй карточке
    second_card.locator(".info-btn").click()
    expect(second_card).to_have_class(re.compile(r"card-info-open"))
    expect(second_card.locator(".card-info-overlay")).to_be_visible()

    # 3. ПРОВЕРКА БАГА: первая карточка обязана вернуться в норму
    expect(first_card.locator(".card-info-overlay")).to_be_hidden()
    expect(first_card).not_to_have_class(re.compile(r"card-info-open"))

def test_homepage_loads(page, app_url):
    """Главная открывается и рендерит базовые элементы."""
    # Без авторизации — редирект на логин
    page.goto(app_url)
    assert "/login" in page.url


def test_homepage_after_login(admin_page, app_url):
    """После логина — главная с карточками."""
    admin_page.goto(app_url)

    # Шапка
    assert admin_page.locator(".topbar").is_visible()
    assert admin_page.locator(".brand-title").text_content() == "Нейрона ИИ"

    # Hero-секция
    assert admin_page.locator(".hero").is_visible()
    assert "Локальная AI-платформа" in admin_page.locator(".hero h1").text_content()


def test_all_module_cards_have_titles(admin_page, app_url):
    """Каждая карточка модуля должна иметь заголовок."""
    admin_page.goto(app_url)

    cards = admin_page.locator(".module-card").all()
    assert len(cards) > 0, "На главной должны быть карточки модулей"

    for card in cards:
        title = card.locator(".module-title").text_content()
        assert title and len(title.strip()) > 0, (
            f"Карточка без заголовка: {card.inner_html()[:200]}"
        )


def test_card_info_button_opens_overlay(admin_page, app_url):
    """Кнопка (i) открывает оверлей с описанием."""
    admin_page.goto(app_url)

    # Берём первую карточку с кнопкой info-btn
    first_card = admin_page.locator(".module-card").first
    info_btn = first_card.locator(".info-btn")
    
    if not info_btn.is_visible():
        pytest.skip("info-btn не найдена на первой карточке")

    info_btn.click()

    # Оверлей должен появиться
    overlay = first_card.locator(".card-info-overlay")
    assert overlay.is_visible(timeout=3000)

    # Должен быть заголовок и кнопка «Назад»
    assert overlay.locator(".card-info-overlay-title").is_visible()
    assert overlay.locator(".card-info-back").is_visible()


def test_card_info_closes_on_back_button(admin_page, app_url):
    """Кнопка «Назад» закрывает оверлей."""
    admin_page.goto(app_url)

    first_card = admin_page.locator(".module-card").first
    info_btn = first_card.locator(".info-btn")
    if not info_btn.is_visible():
        pytest.skip("info-btn не найдена")

    info_btn.click()
    overlay = first_card.locator(".card-info-overlay")
    overlay.wait_for(state="visible")

    overlay.locator(".card-info-back").click()
    overlay.wait_for(state="hidden", timeout=3000)


def test_quick_actions_panel_visible_for_admin(admin_page, app_url):
    """Панель быстрого доступа видна только админу."""
    admin_page.goto(app_url)
    assert admin_page.locator(".quick-actions-bar").is_visible()


def test_whats_new_modal_opens_and_closes(admin_page, app_url):
    """Модалка «Что нового» открывается и закрывается по Escape."""
    admin_page.goto(app_url)

    # Открываем через меню профиля
    admin_page.locator(".profile-menu-trigger").click()
    admin_page.locator("#whatsNewButton").click()

    modal = admin_page.locator("#whatsNewModal")
    modal.wait_for(state="visible", timeout=3000)

    # Закрываем по Escape
    admin_page.keyboard.press("Escape")
    modal.wait_for(state="hidden", timeout=3000)
    