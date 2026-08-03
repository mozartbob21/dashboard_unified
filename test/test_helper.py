# test/test_helper.py
"""Unit-тесты для чистых хелперов app.py и services.run_history."""
import sys
from pathlib import Path

# Корень проекта в sys.path, чтобы импорты app и services работали
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import (
    normalize_key,
    normalize_text,
    to_int,
    calculate_overdue_metrics,
    calculate_standard_status_metrics,
)
from services.run_history import _format_duration, _plural


# ─── normalize_text ───

def test_normalize_text_strips_whitespace():
    assert normalize_text("  привет  ") == "привет"


def test_normalize_text_none_returns_empty():
    assert normalize_text(None) == ""


# ─── normalize_key ───

def test_normalize_key_uppercases_municipality():
    m, o = normalize_key("  Красногорск  ", "ООО Ромашка")
    assert m == "КРАСНОГОРСК"
    assert o == "ооо ромашка"


# ─── to_int ───

def test_to_int_plain_number():
    assert to_int("123") == 123


def test_to_int_spaces_removed():
    assert to_int("1 234") == 1234


def test_to_int_comma_as_decimal():
    assert to_int("45,6") == 45


def test_to_int_none_default():
    assert to_int(None, default=7) == 7


def test_to_int_garbage_returns_default():
    assert to_int("abc") == 0
    assert to_int("abc", default=3) == 3


# ─── calculate_overdue_metrics ───

def test_calculate_overdue_metrics_classifies():
    raw = {
        "items": [
            {"overdue_count": 0},
            {"overdue_count": 5},
            {"overdue_count": 25},
        ]
    }
    m = calculate_overdue_metrics(raw)
    assert m["total"] == 3
    assert m["ok"] == 1
    assert m["risk"] == 1
    assert m["critical"] == 1


def test_calculate_overdue_metrics_empty():
    m = calculate_overdue_metrics(None)
    assert m == {"total": 0, "critical": 0, "risk": 0, "ok": 0}


# ─── calculate_standard_status_metrics ───

def test_standard_metrics_counts_statuses():
    result = {
        "rows": [
            {"status": "critical"},
            {"status": "risk"},
            {"status": "ok"},
            {"status": "норма"},
        ]
    }
    m = calculate_standard_status_metrics(result)
    assert m["total"] == 4
    assert m["critical"] == 1
    assert m["risk"] == 1
    assert m["ok"] == 2


# ─── _format_duration ───

def test_format_duration_seconds():
    assert _format_duration(45) == "45 с"


def test_format_duration_minutes_seconds():
    assert _format_duration(125) == "2 мин 05 с"


def test_format_duration_hours():
    assert _format_duration(3661) == "1 ч 01 мин"


def test_format_duration_none():
    assert _format_duration(None) == ""


# ─── _plural ───

def test_plural_one():
    assert _plural(1, "час", "часа", "часов") == "час"
    assert _plural(21, "час", "часа", "часов") == "час"


def test_plural_few():
    assert _plural(2, "час", "часа", "часов") == "часа"
    assert _plural(24, "час", "часа", "часов") == "часа"


def test_plural_many():
    assert _plural(5, "час", "часа", "часов") == "часов"
    assert _plural(11, "час", "часа", "часов") == "часов"
    assert _plural(112, "час", "часа", "часов") == "часов"
    