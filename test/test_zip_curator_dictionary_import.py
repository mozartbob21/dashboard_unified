"""Uploaded dictionary additions must persist across requests and source imports."""
import ast
import copy
import json
from pathlib import Path

import pytest
from fastapi import FastAPI, Request
from fastapi.testclient import TestClient

from services.zip_curator import core, dictionary


HEADER = ["Наименование (норм.)", "Категория", "Группа", "Вода"]
IMPORT_URL = "/zip_curator/api/dictionary/import"
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@pytest.fixture
def store(monkeypatch, tmp_path):
    base = {
        "categories": ["Приборы", "Трубы"],
        "groupsByCat": {"Приборы": ["Датчики"], "Трубы": ["Стальные трубы"]},
        "dict": {"базовая позиция": ["Трубы", "Стальные трубы", 0],
                 "переназначаемая позиция": ["Трубы", "Стальные трубы", 0]},
        "kws": [["тестключ", "Приборы", "Датчики"]],
    }
    base_path = tmp_path / "base.json"
    base_path.write_text(json.dumps(base, ensure_ascii=False), encoding="utf-8")
    monkeypatch.setattr(dictionary, "BASE_FILE", base_path)
    monkeypatch.setattr(dictionary, "CUSTOM_FILE", tmp_path / "dictionary.json")
    for name, path in {
        "DATA": tmp_path,
        "STATE_FILE": tmp_path / "state.json",
        "PUBLISHED_FILE": tmp_path / "published.json",
        "MUNICIPALITY_OVERRIDES_FILE": tmp_path / "municipalities.json",
    }.items():
        monkeypatch.setattr(core, name, path)
    dictionary._merged.cache_clear()
    dictionary.add_category("Ручная категория", "Ручная группа")
    dictionary.remember("ручная позиция", "Ручная категория", "Ручная группа", True)
    yield tmp_path
    dictionary._merged.cache_clear()


def workbook(store, rows):
    path = store / "upload.xlsx"
    core.write_xlsx(rows, path)
    return path.read_bytes()


def api_client():
    # Compile the production routes without application startup, schedulers, or .env.
    tree = ast.parse((core.BASE / "app.py").read_text(encoding="utf-8-sig"))
    names = {"zc_dictionary", "zc_import_dictionary", "zc_dict"}
    selected = [node for node in tree.body
                if isinstance(node, ast.AsyncFunctionDef) and node.name in names]
    assert {node.name for node in selected} == names
    app = FastAPI()
    module = ast.Module(body=selected, type_ignores=[])
    exec(compile(module, "app.py", "exec"), {"app": app, "zc": core, "Request": Request})
    return TestClient(app)


def upload(client, data, filename="dictionary.xlsx"):
    return client.post(IMPORT_URL, files={"file": (filename, data, MIME)})


def test_import_persists_for_fresh_clients_classification_ingestion_and_export(store):
    base_before = dictionary.BASE_FILE.read_bytes()
    initial = dictionary.load_dictionary()
    data = workbook(store, [
        ["Вода", "Группа", "Наименование", "Категория"],
        ["ДА", "датчики", "  НОВЫЙ   ДАТЧИК  ", "приборы"],
        ["нет", "Новая группа", "Специальная позиция", "Новый раздел"],
        ["yes", "датчики", "Переназначаемая позиция", "ПРИБОРЫ"],
    ])
    with api_client() as client:
        response = upload(client, data)
    assert response.status_code == 200
    assert response.json()["ok"] is True
    assert response.json()["imported"] == 3
    assert response.json()["dictionary"]["dict"]["новый датчик"] == ["Приборы", "Датчики", 1]
    assert dictionary.CUSTOM_FILE.exists()

    dictionary._merged.cache_clear()
    with api_client() as next_client:
        fresh = next_client.get("/zip_curator/api/dictionary")
        assert fresh.status_code == 200
        loaded = fresh.json()
        assert loaded["dict"]["новый датчик"] == ["Приборы", "Датчики", 1]
        assert loaded["dict"]["переназначаемая позиция"] == ["Приборы", "Датчики", 1]
        assert loaded["dict"]["ручная позиция"] == initial["dict"]["ручная позиция"]
        assert loaded["dict"]["базовая позиция"] == initial["dict"]["базовая позиция"]
        assert loaded["kws"] == initial["kws"]
        assert loaded["groupsByCat"]["Новый раздел"] == ["Новая группа"]
        assert loaded["groupsByCat"]["Ручная категория"] == ["Ручная группа"]
        assert dictionary.BASE_FILE.read_bytes() == base_before
        assert core.classify("  Новый   датчик ") == {
            "cat": "Приборы", "grp": "Датчики", "water": True, "via": "match",
        }
        core.ingest([{"rso": "Синтетическое РСО", "items": [
            {"name": "НОВЫЙ ДАТЧИК", "qty": 2, "unitRaw": "шт"},
        ]}])
        item = core.load_state()["pending"][0]["items"][0]
        assert (item["cat"], item["grp"], item["water"], item["via"]) == (
            "Приборы", "Датчики", True, "match",
        )
        exported = next_client.get("/zip_curator/api/dict")
        assert exported.status_code == 200
        rows = core.read_xlsx_rows(exported.content)
        assert rows[0] == HEADER
        assert ["новый датчик", "Приборы", "Датчики", "да"] in rows
        result = upload(next_client, exported.content, "exported.xlsx")
        assert result.status_code == 200
        assert result.json()["imported"] == len(loaded["dict"])
        assert result.json()["dictionary"] == loaded


def test_identical_normalized_duplicates_are_one_entry(store):
    result = dictionary.import_workbook("dictionary.xlsx", workbook(store, [
        HEADER,
        ["  НОВЫЙ   ДАТЧИК ", "приборы", "датчики", "да"],
        ["новый датчик", "ПРИБОРЫ", "ДАТЧИКИ", 1],
    ]))
    assert result["imported"] == 1
    assert result["dictionary"]["dict"]["новый датчик"] == ["Приборы", "Датчики", 1]


def test_optional_water_column_and_blank_classification_round_trip(store):
    result = dictionary.import_workbook("dictionary.xlsx", workbook(store, [
        HEADER[:3],
        ["Без классификации", "", ""],
        ["Только категория", "приборы", ""],
    ]))
    assert result["imported"] == 2
    assert result["dictionary"]["dict"]["без классификации"] == [None, None, 0]
    assert result["dictionary"]["dict"]["только категория"] == ["Приборы", None, 0]
    with api_client() as client:
        exported = client.get("/zip_curator/api/dict")
        assert exported.status_code == 200
        again = upload(client, exported.content)
        assert again.status_code == 200
        assert again.json()["dictionary"] == result["dictionary"]


@pytest.mark.parametrize("token,expected", [
    ("да", 1), ("1", 1), ("true", 1), ("yes", 1), ("+", 1),
    ("нет", 0), ("0", 0), ("false", 0), ("no", 0), ("-", 0), ("", 0),
])
def test_supported_water_values(store, token, expected):
    result = dictionary.import_workbook("dictionary.xlsx", workbook(store, [
        HEADER, ["Новая позиция", "Приборы", "Датчики", token],
    ]))
    assert result["dictionary"]["dict"]["новая позиция"][2] == expected


@pytest.mark.parametrize("rows", [
    [],
    [HEADER],
    [["Что-то", "Категория", "Группа"], ["Позиция", "Приборы", "Датчики"]],
    [["Наименование", "Группа"], ["Позиция", "Датчики"]],
    [HEADER, ["Хорошая позиция", "Новый раздел", "Новая группа", "да"],
     ["Плохая позиция", "", "Группа без категории", "нет"]],
    [HEADER, ["Хорошая позиция", "Новый раздел", "Новая группа", "да"],
     ["Плохая позиция", "Приборы", "Датчики", "возможно"]],
    [HEADER, ["Повтор", "Приборы", "Датчики", "да"],
     [" ПОВТОР ", "Трубы", "Стальные трубы", "да"]],
    [HEADER, ["Повтор", "Приборы", "Датчики", "да"],
     [" ПОВТОР ", "Приборы", "Датчики", "нет"]],
    [["Наименование", "Категория", "Группа", "Группа"], ["Позиция", "Приборы", "Датчики", "Датчики"]],
    [HEADER, ["", "Приборы", "Датчики", "да"]],
], ids=["empty-sheet", "header-only", "unknown-name-header", "missing-category-header",
        "group-without-category", "invalid-water", "conflicting-classification", "conflicting-water",
        "duplicate-header", "empty-name"])
def test_invalid_rows_reject_entire_import_without_changing_saved_dictionary(store, rows):
    before = dictionary.CUSTOM_FILE.read_bytes()
    before_merged = dictionary.load_dictionary()
    with api_client() as client:
        response = upload(client, workbook(store, rows))
    assert response.status_code == 400
    assert response.json()["ok"] is False
    assert response.json()["error"]
    assert dictionary.CUSTOM_FILE.read_bytes() == before
    dictionary._merged.cache_clear()
    assert dictionary.load_dictionary() == before_merged


@pytest.mark.parametrize("filename,body", [
    ("dictionary.xlsx", b"not-an-excel-workbook"),
    ("dictionary.xlsx", b""),
    ("dictionary.txt", b"not-an-excel-workbook"),
])
def test_wrong_file_returns_safe_400_without_changes(store, filename, body):
    before = dictionary.CUSTOM_FILE.read_bytes()
    with api_client() as client:
        response = upload(client, body, filename)
    assert response.status_code == 400
    assert response.json()["ok"] is False
    assert str(store) not in response.text
    assert dictionary.CUSTOM_FILE.read_bytes() == before


def test_missing_upload_returns_400(store):
    before = dictionary.CUSTOM_FILE.read_bytes()
    with api_client() as client:
        response = client.post(IMPORT_URL)
    assert response.status_code == 400
    assert response.json()["ok"] is False
    assert dictionary.CUSTOM_FILE.read_bytes() == before


def test_large_upload_returns_413_without_changes(store):
    before = dictionary.CUSTOM_FILE.read_bytes()
    with api_client() as client:
        response = upload(client, b"x" * (10 * 1024 * 1024 + 1))
    assert response.status_code == 413
    assert response.json()["ok"] is False
    assert dictionary.CUSTOM_FILE.read_bytes() == before


def test_atomic_save_failure_keeps_old_dictionary_and_returns_safe_500(monkeypatch, store):
    before = dictionary.CUSTOM_FILE.read_bytes()
    before_merged = copy.deepcopy(dictionary.load_dictionary())
    data = workbook(store, [HEADER, ["Новая позиция", "Приборы", "Датчики", "да"]])

    def fail_replace(src, dst):
        raise OSError("save denied at /private/server/path?token=SECRET")

    monkeypatch.setattr(dictionary.os, "replace", fail_replace)
    with api_client() as client:
        response = upload(client, data)
    assert response.status_code == 500
    assert response.json()["ok"] is False
    assert "SECRET" not in response.text
    assert "/private/server/path" not in response.text
    assert dictionary.CUSTOM_FILE.read_bytes() == before
    dictionary._merged.cache_clear()
    assert dictionary.load_dictionary() == before_merged
