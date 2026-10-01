import datetime

import pytest

from services.water_dashboard import scraper
from services.water_dashboard.diagnostics import SourceReadError


class Items:
    def __init__(self, items):
        self.items = items

    @property
    def first(self):
        return self.items[0]

    def nth(self, index):
        return self.items[index]

    def count(self):
        return len(self.items)


class Title:
    def __init__(self, text):
        self.text = text

    def inner_text(self, **kwargs):
        return self.text


class Item:
    def __init__(self, page, value, title=None):
        self.page, self.value = page, value
        self.title = value if title is None else title

    def get_attribute(self, name):
        assert name == 'data-value'
        return self.value

    def locator(self, selector):
        assert selector == '.yc-select-item__title'
        return Items([Title(self.title)])

    def scroll_into_view_if_needed(self, **kwargs):
        pass

    def click(self, **kwargs):
        self.page.clicked.append(self.value)
        if not self.page.apply_click:
            return
        if self.value in self.page.selected:
            self.page.selected.remove(self.value)
        else:
            self.page.selected.append(self.value)


class Control:
    def __init__(self, page):
        self.page = page

    def inner_text(self, **kwargs):
        return ', '.join(self.page.selected) or 'Нет выбранных значений'

    def click(self, **kwargs):
        self.page.popup_open = True


class Clear:
    def __init__(self, page):
        self.page = page

    def is_visible(self):
        return self.page.popup_open

    def click(self, **kwargs):
        self.page.clear_count += 1
        self.page.selected = []
        if self.page.close_on_clear:
            self.page.popup_open = False


class Popup:
    def __init__(self, page):
        self.page = page

    def wait_for(self, **kwargs):
        if not self.page.popup_open:
            raise RuntimeError('Popup is not visible')

    def is_visible(self):
        return self.page.popup_open

    def locator(self, selector):
        assert selector == '.yc-select-item[data-value]'
        return Items(self.page.items)

    def get_by_text(self, text, **kwargs):
        assert text == 'Очистить' and kwargs == {'exact': True}
        return Items([Clear(self.page)] if self.page.has_clear else [])


class Keyboard:
    def __init__(self, page):
        self.page = page

    def press(self, key):
        assert key == 'Escape'
        self.page.popup_open = False


class Page:
    def __init__(self, dates, selected=(), *, has_control=True,
                 has_clear=True, apply_click=True, close_on_clear=False):
        self.items = [Item(self, value) for value in dates]
        self.selected = list(selected)
        self.has_control, self.has_clear = has_control, has_clear
        self.apply_click, self.close_on_clear = apply_click, close_on_clear
        self.popup_open = False
        self.clicked = []
        self.clear_count = 0
        self.keyboard = Keyboard(self)

    def evaluate(self, script):
        assert 'chartkit-control-title' in script and 'chartkit-control-select' in script
        return 'nvos-date' if self.has_control else None

    def locator(self, selector):
        if selector == '[data-pw-id="nvos-date"]':
            return Items([Control(self)])
        assert selector == '.yc-select-popup:visible'
        return Items([Popup(self)])

    def wait_for_timeout(self, milliseconds):
        pass


@pytest.fixture(autouse=True)
def fixed_today(monkeypatch):
    class FixedDate(datetime.date):
        @classmethod
        def today(cls):
            return cls(2026, 10, 1)
    monkeypatch.setattr(scraper._dt, 'date', FixedDate)


def test_latest_selected_date_is_not_toggled_off():
    page = Page(['2026-08-10', '2026-09-10', '2026-10-10'], ['2026-09-10'])
    assert scraper._select_latest_date(page, 'nvos') == '2026-09-10'
    assert page.selected == ['2026-09-10']
    assert page.clicked == [] and page.clear_count == 0
    assert page.popup_open is False


@pytest.mark.parametrize('close_on_clear', [False, True])
def test_old_selection_is_cleared_before_selecting_latest(close_on_clear):
    page = Page(['2026-09-10', '2026-08-10', '2026-10-10'],
                ['2026-08-10'], close_on_clear=close_on_clear)
    scraper._select_latest_date(page, 'nvos')
    assert page.selected == ['2026-09-10']
    assert page.clicked == ['2026-09-10'] and page.clear_count == 1


def test_multiple_dates_become_only_latest_date():
    page = Page(['2026-09-10', '2026-08-10'], ['2026-08-10', '2026-09-10'])
    scraper._select_latest_date(page, 'nvos')
    assert page.selected == ['2026-09-10'] and page.clear_count == 1


def test_mixed_date_formats_are_sorted_chronologically_not_lexically():
    page = Page(['2026-08-31', '10.09.2026', '02.10.2026'])
    scraper._select_latest_date(page, 'nvos')
    assert page.selected == ['10.09.2026']


def test_matching_display_date_with_another_format_does_not_toggle():
    page = Page(['2026-09-10'], ['10.09.2026'])
    scraper._select_latest_date(page, 'nvos')
    assert page.clicked == [] and page.selected == ['10.09.2026']


@pytest.mark.parametrize('page', [
    Page(['2026-09-10'], apply_click=False),
    Page(['2026-09-10'], ['2026-08-10'], has_clear=False),
    Page(['2026-10-02']),
    Page(['2026-02-31', '31.09.2026']),
    Page(['2026-09-10'], has_control=False),
])
def test_missing_or_unconfirmed_date_is_a_source_error(page):
    with pytest.raises(SourceReadError) as caught:
        scraper._select_latest_date(page, 'nvos')
    assert caught.value.code == 'date'
