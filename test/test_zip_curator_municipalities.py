import json

from services.zip_curator import core


def configure_files(monkeypatch, tmp_path):
    monkeypatch.setattr(core, "DATA", tmp_path)
    monkeypatch.setattr(core, "STATE_FILE", tmp_path / "state.json")
    monkeypatch.setattr(core, "MUNICIPALITY_OVERRIDES_FILE", tmp_path / "municipality_overrides.json")


def test_manual_municipality_is_normalized_and_has_priority(monkeypatch, tmp_path):
    configure_files(monkeypatch, tmp_path)
    core.save_municipality_override("  МУП   Водоканал  ", "ГО Лесной")

    assert core.resolve_municipality("муп водоканал") == "ГО Лесной"
    saved = json.loads((tmp_path / "municipality_overrides.json").read_text(encoding="utf-8"))
    assert saved["муп водоканал"]["organization"] == "МУП   Водоканал"


def test_saving_mapping_updates_pending_and_clean_state(monkeypatch, tmp_path):
    configure_files(monkeypatch, tmp_path)
    core.save_state({
        "pending": [{"rso": "МУП Водоканал", "okrug": "", "items": []}],
        "clean": {"муп водоканал": {"rso": "муп водоканал", "okrug": "", "items": []}},
    })

    core.save_municipality_override("МУП Водоканал", "ИСТРА")

    state = core.load_state()
    assert state["pending"][0]["okrug"] == "ИСТРА"
    assert state["clean"]["муп водоканал"]["okrug"] == "ИСТРА"


def test_delete_falls_back_to_builtin_dictionary(monkeypatch, tmp_path):
    configure_files(monkeypatch, tmp_path)
    organization = "МУП Истринский водоканал"
    core.save_municipality_override(organization, "РУЧНОЙ")
    assert core.delete_municipality_override(organization) is True
    assert core.resolve_municipality(organization.lower()) == "ИСТРА"
