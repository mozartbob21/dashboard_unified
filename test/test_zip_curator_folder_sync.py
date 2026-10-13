"""Folder A failures must not destroy cached registries or stop unrelated files."""
import copy
import hashlib
import json
import os
from pathlib import Path

import pytest
import requests

from services.zip_curator import core


API = "https://cloud-api.yandex.net/v1/disk/public/resources"
PUBLIC_URL = "https://disk.yandex.ru/d/synthetic-folder-a"
DOWNLOAD = "https://downloader.disk.yandex.ru/"


class FakeResponse:
    def __init__(self, *, payload=None, content=b"", status=200, url="", headers=None):
        self.payload = payload
        self.content = content
        self.status_code = status
        self.url = url
        self.headers = dict(headers or {})
        if content:
            self.headers.setdefault("Content-Length", str(len(content)))
        self.closed = False

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    def close(self):
        self.closed = True

    def json(self):
        return copy.deepcopy(self.payload)

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(
                f"{self.status_code} Server Error for url: {self.url}", response=self
            )

    def iter_content(self, chunk_size=8192):
        for offset in range(0, len(self.content), chunk_size):
            yield self.content[offset:offset + chunk_size]


class FakeSession:
    """Explicit routes: any unexpected request fails, with no real network fallback."""
    def __init__(self, routes):
        self.routes = {key: list(values) for key, values in routes.items()}
        self.calls = []
        self.headers = {}

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False

    def close(self):
        pass

    def get(self, url, **kwargs):
        self.calls.append((url, kwargs))
        params = kwargs.get("params") or {}
        key = (url, params.get("path")) if url == API + "/download" else url
        if url == API:
            key = (url, params.get("offset", 0))
        values = self.routes.get(key)
        assert values, f"Unexpected request: {url} {params}"
        result = values.pop(0)
        if isinstance(result, Exception):
            raise result
        result.url = url
        return result


@pytest.fixture
def local_store(monkeypatch, tmp_path):
    input_dir = tmp_path / "input"
    for name, value in {
        "DATA": tmp_path,
        "INPUT_DIR": input_dir,
        "FOLDER_A_DIR": input_dir / "folder_a",
        "STATE_FILE": tmp_path / "state.json",
        "PUBLISHED_FILE": tmp_path / "published.json",
        "FOLDER_A_MANIFEST_FILE": tmp_path / "folder_a_manifest.json",
        "MUNICIPALITY_OVERRIDES_FILE": tmp_path / "municipalities.json",
    }.items():
        monkeypatch.setattr(core, name, value)
    from services.zip_curator import yandex_disk
    monkeypatch.setattr(yandex_disk.time, "sleep", lambda _: None)
    return tmp_path


def workbook(tmp_path, rso="Тестовое РСО", qty=3):
    path = tmp_path / "fixture.xlsx"
    core.write_xlsx([
        [rso],
        ["Наименование", "Ед.", "Количество"],
        ["Труба", "шт", qty],
    ], path)
    return path.read_bytes()


def remote_file(name, data, *, href=None, modified="2026-10-13T10:00:00Z"):
    return {
        "type": "file", "name": name, "path": "/" + name,
        "file": href or DOWNLOAD + name,
        "md5": hashlib.md5(data).hexdigest(), "size": len(data), "modified": modified,
    }


def listing(items, *, total=None, offset=0):
    return FakeResponse(payload={
        "type": "dir",
        "_embedded": {"items": items, "total": len(items) if total is None else total,
                      "offset": offset, "limit": 200},
    })


def use_session(monkeypatch, routes):
    session = FakeSession(routes)
    monkeypatch.setattr(core.requests, "Session", lambda: session)
    return session


def cache(items_and_data):
    core.FOLDER_A_DIR.mkdir(parents=True, exist_ok=True)
    entries = []
    for item, data in items_and_data:
        local = core.FOLDER_A_DIR / item["name"]
        local.write_bytes(data)
        os.utime(local, (1700000000, 1700000000))
        entries.append({
            "name": item["name"], "local_name": item["name"],
            "remote_path": item["path"], "md5": item["md5"],
            "modified": item["modified"], "size": item["size"],
        })
    core.FOLDER_A_MANIFEST_FILE.write_text(json.dumps({
        "source": PUBLIC_URL, "synced_at": "2026-10-01T10:00:00", "files": entries,
    }), encoding="utf-8")
    return entries


def read_manifest():
    return json.loads(core.FOLDER_A_MANIFEST_FILE.read_text(encoding="utf-8"))


@pytest.mark.parametrize("failure", [
    FakeResponse(status=500), FakeResponse(status=403), requests.Timeout("signed-url-timeout"),
], ids=["server-error", "expired-link", "network-timeout"])
def test_download_failure_obtains_fresh_link_and_finishes(monkeypatch, local_store, failure):
    data = workbook(local_store)
    item = remote_file("registry.xlsx", data, href=DOWNLOAD + "expired?token=SECRET")
    fresh = DOWNLOAD + "fresh?token=FRESH_SECRET"
    session = use_session(monkeypatch, {
        (API, 0): [listing([item])],
        item["file"]: [failure],
        (API + "/download", item["path"]): [FakeResponse(payload={"href": fresh})],
        fresh: [FakeResponse(content=data)],
    })

    result = core.sync_and_scan_folder_a(PUBLIC_URL)

    assert result["downloaded"] == 1
    assert result.get("failed", 0) == 0
    assert not result.get("partial", False)
    assert (core.FOLDER_A_DIR / item["name"]).read_bytes() == data
    assert result["state"]["pending"][0]["rso"] == "Тестовое РСО"
    urls = [url for url, _ in session.calls]
    assert urls == [API, item["file"], API + "/download", fresh]


def test_failed_changed_file_keeps_old_cache_metadata_and_processes_good_file(monkeypatch, local_store):
    old_bytes = workbook(local_store, "РСО со старой копией", 1)
    new_bytes = workbook(local_store, "РСО со старой копией", 5)
    good_bytes = workbook(local_store, "Другое РСО", 7)
    old_item = remote_file("a-cached.xlsx", old_bytes, modified="2026-10-01T10:00:00Z")
    orphan_item = remote_file("not-listed.xlsx", old_bytes)
    previous = cache([(old_item, old_bytes), (orphan_item, old_bytes)])
    old_registry = core.parse_reestr(core.read_xlsx_rows(old_bytes), old_item["name"])
    core.ingest([old_registry])
    state = core.load_state()
    # A curator's correction must survive failure to refresh this source.
    state["pending"][0]["items"][0]["qty"] = 77
    core.save_state(state)
    pending_before = copy.deepcopy(state["pending"][0])
    changed = remote_file("a-cached.xlsx", new_bytes)
    good = remote_file("b-good.xlsx", good_bytes)
    fresh = DOWNLOAD + "fresh-failing?token=SECRET"
    use_session(monkeypatch, {
        (API, 0): [listing([changed, good])],
        changed["file"]: [FakeResponse(status=500)],
        (API + "/download", changed["path"]): [FakeResponse(payload={"href": fresh})] * 2,
        fresh: [FakeResponse(status=500), FakeResponse(status=500)],
        good["file"]: [FakeResponse(content=good_bytes)],
    })

    result = core.sync_and_scan_folder_a(PUBLIC_URL)

    assert result["ok"] is True
    assert result["partial"] is True
    assert result["files"] == 2
    assert result["downloaded"] == 1
    assert result["failed"] == 1
    assert result["removed"] == 0
    assert result["errors"][0]["name"] == changed["name"]
    assert result["errors"][0]["retained"] is True
    assert "SECRET" not in json.dumps(result["errors"])
    assert "downloader.disk.yandex.ru" not in json.dumps(result["errors"])
    cached_path = core.FOLDER_A_DIR / changed["name"]
    assert cached_path.read_bytes() == old_bytes
    assert cached_path.stat().st_mtime == 1700000000
    assert (core.FOLDER_A_DIR / orphan_item["name"]).exists()
    manifest = {entry["remote_path"]: entry for entry in read_manifest()["files"]}
    for field in ("md5", "modified", "size"):
        assert manifest[changed["path"]][field] == previous[0][field]
    pending = {entry["rso"]: entry for entry in result["state"]["pending"]}
    assert pending["РСО со старой копией"] == pending_before
    assert pending["Другое РСО"]["items"][0]["qty"] == 7


def test_failed_new_file_does_not_stop_later_files(monkeypatch, local_store):
    data = workbook(local_store)
    bad = remote_file("a-bad.xlsx", data)
    good = remote_file("b-good.xlsx", data)
    fresh = DOWNLOAD + "still-unavailable"
    use_session(monkeypatch, {
        (API, 0): [listing([bad, good])],
        bad["file"]: [FakeResponse(status=503)],
        (API + "/download", bad["path"]): [FakeResponse(payload={"href": fresh})] * 2,
        fresh: [FakeResponse(status=503), FakeResponse(status=503)],
        good["file"]: [FakeResponse(content=data)],
    })

    result = core.sync_and_scan_folder_a(PUBLIC_URL)

    assert result["partial"] is True
    assert result["failed"] == result["downloaded"] == 1
    assert result["errors"][0]["retained"] is False
    assert not (core.FOLDER_A_DIR / bad["name"]).exists()
    assert (core.FOLDER_A_DIR / good["name"]).read_bytes() == data
    assert len(result["state"]["pending"]) == 1


@pytest.mark.parametrize("payload", [
    {},
    {"type": "dir"},
    {"type": "file", "_embedded": {"items": [], "total": 0, "offset": 0, "limit": 200}},
    {"type": "dir", "_embedded": {"items": [], "total": 1, "offset": 0, "limit": 200}},
    {"type": "dir", "_embedded": {"items": "bad", "total": 1, "offset": 0, "limit": 200}},
], ids=["empty-payload", "missing-items", "not-folder", "incomplete-page", "invalid-items"])
def test_invalid_listing_does_not_prune_or_modify_existing_data(monkeypatch, local_store, payload):
    data = workbook(local_store)
    item = remote_file("cached.xlsx", data)
    cache([(item, data)])
    core.save_state({"pending": [], "clean": {}, "excluded_rso": {}})
    before_manifest = core.FOLDER_A_MANIFEST_FILE.read_bytes()
    before_state = core.STATE_FILE.read_bytes()
    use_session(monkeypatch, {(API, 0): [FakeResponse(payload=payload)] * 3})

    with pytest.raises(core.DiskSyncError):
        core.sync_and_scan_folder_a(PUBLIC_URL)

    assert (core.FOLDER_A_DIR / item["name"]).read_bytes() == data
    assert core.FOLDER_A_MANIFEST_FILE.read_bytes() == before_manifest
    assert core.STATE_FILE.read_bytes() == before_state


def test_http_200_invalid_workbook_cannot_replace_previous_cache(monkeypatch, local_store):
    old_data = workbook(local_store)
    invalid_data = b"<html>Temporary download service failure</html>"
    old = remote_file("cached.xlsx", old_data)
    previous = cache([(old, old_data)])[0]
    # The advertised checksum matches the invalid response: structure must also be checked.
    changed = remote_file("cached.xlsx", invalid_data)
    fresh = DOWNLOAD + "fresh-invalid"
    use_session(monkeypatch, {
        (API, 0): [listing([changed])],
        changed["file"]: [FakeResponse(content=invalid_data)],
        (API + "/download", changed["path"]): [FakeResponse(payload={"href": fresh})] * 2,
        fresh: [FakeResponse(content=invalid_data), FakeResponse(content=invalid_data)],
    })

    result = core.sync_folder_a(PUBLIC_URL)

    assert result["partial"] is True
    assert result["failed"] == 1
    assert result["downloaded"] == 0
    assert result["errors"][0]["retained"] is True
    assert (core.FOLDER_A_DIR / changed["name"]).read_bytes() == old_data
    assert read_manifest()["files"][0]["md5"] == previous["md5"]
    assert not list(core.FOLDER_A_DIR.glob("*.part*"))


def test_corrupt_local_file_is_redownloaded_even_when_manifest_matches(monkeypatch, local_store):
    data = workbook(local_store)
    item = remote_file("cached.xlsx", data)
    cache([(item, data)])
    local_path = core.FOLDER_A_DIR / item["name"]
    # Same byte length prevents a size-only cache check from falsely passing.
    local_path.write_bytes(b"x" * len(data))
    session = use_session(monkeypatch, {
        (API, 0): [listing([item])], item["file"]: [FakeResponse(content=data)],
    })

    result = core.sync_folder_a(PUBLIC_URL)

    assert result["downloaded"] == 1
    assert result["reused"] == 0
    assert local_path.read_bytes() == data
    assert [url for url, _ in session.calls] == [API, item["file"]]


def test_verified_unchanged_cache_requires_no_download(monkeypatch, local_store):
    data = workbook(local_store)
    item = remote_file("cached.xlsx", data)
    cache([(item, data)])
    session = use_session(monkeypatch, {(API, 0): [listing([item])]})

    result = core.sync_folder_a(PUBLIC_URL)

    assert result["downloaded"] == 0
    assert result["reused"] == 1
    assert [url for url, _ in session.calls] == [API]


def test_sync_does_not_requeue_or_replace_approved_rso(monkeypatch, local_store):
    existing = {
        "rso": "Согласованное РСО", "okrug": "Тестовый округ", "date": "01.10.2026",
        "items": [{"name": "Труба", "nn": "труба", "qty": 1, "unitRaw": "шт",
                   "unit": "шт", "unitUnknown": False, "cat": "Трубы", "grp": "Труба",
                   "water": False, "via": "match"}],
    }
    key = core.norm(existing["rso"])
    core.save_state({"pending": [], "clean": {key: existing}, "excluded_rso": {}})
    core.publish_clean()
    published_before = core.PUBLISHED_FILE.read_bytes()
    approved_before = copy.deepcopy(core.load_state()["clean"][key])
    data = workbook(local_store, existing["rso"], 999)
    item = remote_file("approved.xlsx", data)
    use_session(monkeypatch, {(API, 0): [listing([item])], item["file"]: [FakeResponse(content=data)]})

    result = core.sync_and_scan_folder_a(PUBLIC_URL)

    assert result["downloaded"] == 1
    assert result["state"]["pending"] == []
    assert result["state"]["clean"][key] == approved_before
    assert core.PUBLISHED_FILE.read_bytes() == published_before


def test_parallel_refresh_is_rejected_and_lock_released_after_failure(monkeypatch, local_store):
    session = use_session(monkeypatch, {(API, 0): [FakeResponse(payload={}), listing([])]})
    with core._folder_sync_lock():
        with pytest.raises(core.FolderSyncBusy):
            core.sync_folder_a(PUBLIC_URL)
    assert session.calls == []
    with pytest.raises(core.DiskSyncError):
        core.sync_folder_a(PUBLIC_URL)
    # A failed first run must not leave the next refresh permanently busy.
    result = core.sync_folder_a(PUBLIC_URL)
    assert result["ok"] is True
    assert result["files"] == 0


def test_complete_empty_folder_prunes_only_previously_tracked_files(monkeypatch, local_store):
    data = workbook(local_store)
    old = remote_file("removed.xlsx", data)
    cache([(old, data)])
    untracked = core.FOLDER_A_DIR / "manually-staged.xlsx"
    untracked.write_bytes(data)
    use_session(monkeypatch, {(API, 0): [listing([])]})

    result = core.sync_folder_a(PUBLIC_URL)

    assert result["files"] == 0
    assert not result["partial"]
    assert result["removed"] == 1
    assert not (core.FOLDER_A_DIR / old["name"]).exists()
    assert untracked.read_bytes() == data
    assert read_manifest()["files"] == []


def test_storage_lookup_failure_is_explained_without_raw_response():
    from services.zip_curator.yandex_disk import DiskClient, DiskSyncError
    response = FakeResponse(status=500, content=b'file IO failure: no successful lookup result\n')
    with pytest.raises(DiskSyncError, match='не может прочитать файл'):
        DiskClient._check_status(response, download=True)


def test_name_collision_checks_generated_suffix_too():
    item = {'name': 'registry.xlsx', 'path': '/synthetic'}
    digest = hashlib.sha1(item['path'].encode()).hexdigest()[:8]
    used = {'registry.xlsx', f'registry-{digest}.xlsx'}
    result = core._safe_remote_name(item, used)
    assert result == f'registry-{digest}-2.xlsx'


def test_deadline_stops_requests(monkeypatch):
    from services.zip_curator import yandex_disk
    session = FakeSession({})
    client = yandex_disk.DiskClient(session, PUBLIC_URL)
    client.deadline = 0
    with pytest.raises(core.DiskSyncError, match='Время обновления истекло'):
        client.download({'path': '/file.xlsx', 'file': DOWNLOAD + 'file.xlsx'}, core.read_xlsx_rows)
    assert not session.calls


def test_partial_sync_keeps_absent_cache_tracked_for_later_cleanup(monkeypatch, local_store):
    old_data = workbook(local_store, rso='РСО старое')
    absent = remote_file('absent.xlsx', old_data)
    changing = remote_file('changing.xlsx', old_data)
    cache([(absent, old_data), (changing, old_data)])
    new_data = workbook(local_store, rso='РСО новое', qty=9)
    updated = remote_file('changing.xlsx', new_data)
    fresh = DOWNLOAD + 'fresh-changing'
    use_session(monkeypatch, {
        (API, 0): [listing([updated])],
        updated['file']: [FakeResponse(status=500)],
        (API + '/download', updated['path']): [FakeResponse(payload={'href':fresh}), FakeResponse(payload={'href':fresh})],
        fresh: [FakeResponse(status=500),FakeResponse(status=500)],
    })
    result=core.sync_folder_a(PUBLIC_URL)
    assert result['partial']
    assert {item['name'] for item in read_manifest()['files']} == {'absent.xlsx','changing.xlsx'}
    assert (core.FOLDER_A_DIR/'absent.xlsx').exists()
    use_session(monkeypatch, {(API,0):[listing([updated])], updated['file']:[FakeResponse(content=new_data)]})
    result=core.sync_folder_a(PUBLIC_URL)
    assert not result['partial']
    assert result['removed']==1
    assert not (core.FOLDER_A_DIR/'absent.xlsx').exists()
