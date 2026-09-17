import json

from services.zip_curator import core


def configure_files(monkeypatch, tmp_path):
    input_dir = tmp_path / "input"
    monkeypatch.setattr(core, "DATA", tmp_path)
    monkeypatch.setattr(core, "INPUT_DIR", input_dir)
    monkeypatch.setattr(core, "FOLDER_A_DIR", input_dir / "folder_a")
    monkeypatch.setattr(core, "STATE_FILE", tmp_path / "state.json")
    monkeypatch.setattr(core, "PUBLISHED_FILE", tmp_path / "published.json")
    monkeypatch.setattr(core, "MUNICIPALITY_OVERRIDES_FILE", tmp_path / "municipality_overrides.json")
    return input_dir


def item(name="Труба", qty=1):
    return {
        "name": name,
        "nn": core.norm(name),
        "cat": "Трубы",
        "grp": "Труба",
        "qty": qty,
        "unit": "шт",
        "unitRaw": "шт",
        "unitUnknown": False,
        "water": False,
        "via": "match",
    }


def registry(rso):
    return {"rso": rso, "okrug": "", "date": "17.09.2026", "items": [item()]}


def write_registry(path, rso):
    core.write_xlsx([
        [rso],
        ["Наименование", "Ед.", "Количество"],
        ["Труба", "шт", 1],
    ], path)


def test_folder_scan_skips_already_approved_rso(monkeypatch, tmp_path):
    input_dir = configure_files(monkeypatch, tmp_path)
    input_dir.mkdir()
    approved = registry("Уже согласованное РСО")
    core.save_state({
        # Имитируем дубль, который успел попасть в очередь до обновления логики.
        "pending": [approved],
        "clean": {core.norm(approved["rso"]): approved},
        "excluded_rso": {},
    })
    write_registry(input_dir / "approved.xlsx", approved["rso"])
    write_registry(input_dir / "new.xlsx", "Новое РСО")

    added, skipped = core.scan_folder(input_dir)

    state = core.load_state()
    assert added == 1
    assert skipped == 1
    assert [entry["rso"] for entry in state["pending"]] == ["Новое РСО"]


def test_delete_rso_removes_only_target_and_suppresses_folder_restore(monkeypatch, tmp_path):
    input_dir = configure_files(monkeypatch, tmp_path)
    input_dir.mkdir()
    first = registry("РСО с ошибочным названием")
    second = registry("Правильное РСО")
    core.save_state({
        "pending": [first],
        "clean": {core.norm(first["rso"]): first, core.norm(second["rso"]): second},
        "excluded_rso": {},
    })

    result = core.delete_rso(first["rso"])

    state = core.load_state()
    assert result["removed"] is True
    assert list(state["clean"]) == [core.norm(second["rso"])]
    assert state["pending"] == []
    assert core.norm(first["rso"]) in state["excluded_rso"]
    published = json.loads(core.PUBLISHED_FILE.read_text(encoding="utf-8"))["rows"]
    assert {row[0] for row in published[1:]} == {second["rso"]}

    write_registry(input_dir / "deleted.xlsx", first["rso"])
    added, skipped = core.scan_folder(input_dir)
    assert (added, skipped) == (0, 1)
    assert core.load_state()["pending"] == []
