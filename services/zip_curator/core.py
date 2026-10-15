"""Куратор ЗиП: загрузка исходников, согласование и локальная публикация."""
import hashlib, io, json, math, os, re, tempfile, threading, unicodedata, zipfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import requests
from services.zip_curator import dictionary
from services.zip_curator.yandex_disk import DiskClient, DiskSyncError, MAX_FILE_BYTES, validate_workbook

BASE = Path(__file__).resolve().parents[2]
DATA = BASE / "data" / "zip_curator"
INPUT_DIR = Path(__import__("os").getenv("ZIP_INPUT_DIR", str(DATA / "input")))
FOLDER_A_DIR = INPUT_DIR / "folder_a"
STATE_FILE = DATA / "state.json"
PUBLISHED_FILE = DATA / "published.json"
FOLDER_A_MANIFEST_FILE = DATA / "folder_a_manifest.json"
MUNICIPALITY_OVERRIDES_FILE = DATA / "municipality_overrides.json"
DICT_JSON = Path(__file__).with_name("zip_dict.json")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SOURCE_DATE_FIELDS = ("source_created_at", "source_modified_at")
SOURCE_INFO_FIELDS = (*SOURCE_DATE_FIELDS, "source_provider", "source_identity", "source_path",
                      "source_md5", "source_etag", "source_size")
FINGERPRINT_FIELDS = ("source_sha256", "source_record_hash")
SOURCE_FIELDS = ("loaded_at", "uploaded", "fname", "local_file", "rso_explicit", *FINGERPRINT_FIELDS, *SOURCE_INFO_FIELDS)

D = json.loads(DICT_JSON.read_text(encoding="utf-8"))

_MUTATION_CONTEXT = threading.local()


@contextmanager
def _owned_curator_mutation():
    previous = getattr(_MUTATION_CONTEXT, "active", False)
    _MUTATION_CONTEXT.active = True
    try:
        yield
    finally:
        _MUTATION_CONTEXT.active = previous


@contextmanager
def _curator_mutation_lock():
    if getattr(_MUTATION_CONTEXT, "active", False):
        yield
        return
    try:
        with _folder_sync_lock(), _owned_curator_mutation():
            yield
    except FolderSyncBusy:
        raise CuratorConflictError("Данные ЗиП сейчас обновляются. Дождитесь завершения и повторите действие.") from None


def _col(s):
    n = 0
    for ch in s: n = n*26 + (ord(ch)-64)
    return n-1

def read_xlsx_rows(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        r = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in r.findall(NS+"si"):
            shared.append("".join(t.text or "" for t in si.iter(NS+"t")))
    sheets = sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", n))
    if not sheets: return []
    r = ET.fromstring(z.read(sheets[0]))
    rows = []
    for row in r.iter(NS+"row"):
        cells = {}
        for c in row.findall(NS+"c"):
            ref = c.get("r") or ""
            ci = _col(re.match(r"([A-Z]+)", ref).group(1)) if ref and re.match(r"([A-Z]+)", ref) else (max(cells)+1 if cells else 0)
            t = c.get("t"); v = c.find(NS+"v"); is_ = c.find(NS+"is")
            if t == "s" and v is not None: val = shared[int(v.text)] if v.text else ""
            elif t == "inlineStr" and is_ is not None: val = "".join(x.text or "" for x in is_.iter(NS+"t"))
            elif v is not None: val = v.text or ""
            else: val = ""
            cells[ci] = val
        if cells:
            rows.append([cells.get(i) for i in range(max(cells)+1)])
    return rows

def norm(s):
    s = str(s or "").lower().strip()
    s = re.sub(r"^\(\d+[.\d]*\)", "", s)
    s = re.sub(r"^\d+[а-яё]?\s", "", s)
    s = re.sub(r"[,\s]+$", "", s)
    return re.sub(r"\s+", " ", s).strip()

def load_municipality_overrides():
    """Return curator-maintained organization -> municipality mappings."""
    if MUNICIPALITY_OVERRIDES_FILE.exists():
        try:
            data = json.loads(MUNICIPALITY_OVERRIDES_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, json.JSONDecodeError):
            pass
    return {}

def resolve_municipality(organization):
    """Resolve an organization, preferring the curator's additional dictionary."""
    key = norm(organization)
    override = load_municipality_overrides().get(key)
    if override:
        return str(override.get("municipality") or "").strip()
    for name, municipality in D.get("rso2omsu", {}).items():
        if norm(name) == key:
            return municipality
    return ""

@_curator_mutation_lock()
def save_municipality_override(organization, municipality):
    organization = str(organization or "").strip()
    municipality = str(municipality or "").strip()
    key = norm(organization)
    if not key or not municipality:
        raise ValueError("Укажите организацию и муниципалитет")
    entries = load_municipality_overrides()
    entry = {
        "organization": organization,
        "municipality": municipality,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    entries[key] = entry
    DATA.mkdir(parents=True, exist_ok=True)
    MUNICIPALITY_OVERRIDES_FILE.write_text(
        json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    _apply_municipality_to_state(key, municipality)
    return entry

@_curator_mutation_lock()
def delete_municipality_override(organization):
    key = norm(organization)
    entries = load_municipality_overrides()
    if key not in entries:
        return False
    del entries[key]
    DATA.mkdir(parents=True, exist_ok=True)
    MUNICIPALITY_OVERRIDES_FILE.write_text(
        json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    _apply_municipality_to_state(key, resolve_municipality(organization))
    return True

def _apply_municipality_to_state(organization_key, municipality):
    """Keep already loaded and approved data consistent with a changed mapping."""
    st = load_state()
    changed = False
    for item in st.get("pending", []):
        if norm(item.get("rso")) == organization_key:
            item["okrug"] = municipality
            changed = True
    for item in st.get("clean", {}).values():
        if norm(item.get("rso")) == organization_key:
            item["okrug"] = municipality
            changed = True
    if changed:
        if any(norm(item.get("rso")) == organization_key and item.get("items") for item in st.get("clean", {}).values()):
            publish_clean(st)
        else:
            save_state(st)

def _unit_key(value):
    # Excel exports use dots, non-breaking spaces and both м3 / м³.
    value = unicodedata.normalize("NFKC", str(value)).casefold()
    return re.sub(r"[\s.]+", "", value)


UNIT_ALIASES = {
    "шт": ("штука", "штуки", "штук"),
    "м": ("метр", "метры", "метров"),
    "пог. м": ("пог м", "м п", "погонный метр", "погонных метров"),
    "кг": ("килограмм", "килограммы", "килограммов"),
    "т": ("тн", "тонна", "тонны", "тонн"),
    "л": ("литр", "литры", "литров"),
    "м³": ("м3", "м куб", "куб м", "кубометр", "кубический метр"),
    "пар": ("пара", "пары"),
    "компл": ("комплект", "комплекты", "комплектов"),
    "упак": ("упаковка", "упаковки", "упаковок"),
    "рул": ("рулон", "рулоны", "рулонов"),
    "пачк": ("пачка", "пачки", "пачек"),
    "бал": ("баллон", "баллоны", "баллонов"),
}
UNIT_LOOKUP = {
    _unit_key(alias): canonical
    for canonical, aliases in UNIT_ALIASES.items()
    for alias in (canonical, *aliases)
}
UNIT_LOOKUP.update({_unit_key(alias): canonical for alias, canonical in D["umap"].items()})


def norm_unit(u):
    if u is None or not str(u).strip():
        return None, True
    canonical = UNIT_LOOKUP.get(_unit_key(u))
    return canonical, canonical is None


_UNSET = object()


class ApprovalValidationError(ValueError):
    """Incomplete registries remain in the queue; callers can display each issue."""
    def __init__(self, issues):
        self.issues = issues
        messages = [issue["message"] for issue in issues]
        summary = " ".join(messages[:6])
        if len(messages) > 6:
            summary += f" Ещё замечаний: {len(messages) - 6}."
        super().__init__("Нельзя согласовать реестр. " + summary)


class CuratorConflictError(ValueError):
    pass


def _missing_label(value):
    return not isinstance(value, str) or norm(value) in {
        "", "-", "—", "–", "не указано", "не указан", "не определен", "не определён",
        "муниципалитет не определён", "муниципалитет не определен",
    }


def _quantity(value):
    if value is None or isinstance(value, str) and not value.strip():
        return None
    if isinstance(value, bool) or not isinstance(value, (str, int, float)):
        raise ValueError("Количество должно быть числом не меньше нуля")
    try:
        number = float(re.sub(r"\s+", "", value).replace(",", ".")) if isinstance(value, str) else float(value)
    except (ValueError, OverflowError):
        raise ValueError("Количество должно быть числом не меньше нуля") from None
    if not math.isfinite(number) or number < 0:
        raise ValueError("Количество должно быть конечным числом не меньше нуля")
    return number


def approval_issues(record, classification=None):
    """Authoritative completeness checks, with zero-based row references for UI."""
    current = classification if classification is not None else dictionary.load_dictionary()
    issues = []
    def add(field, message, index=None):
        issue = {"field": field, "message": message}
        if index is not None:
            issue["item_index"] = index
        issues.append(issue)

    if _missing_label(record.get("rso")):
        add("rso", "Не указано название РСО.")
    municipality = record.get("okrug")
    if _missing_label(municipality):
        municipality = resolve_municipality(record.get("rso"))
    if _missing_label(municipality):
        add("okrug", "Не указан ОМСУ: нажмите «Указать» рядом с названием РСО.")
    items = record.get("items")
    if not isinstance(items, list) or not items:
        add("items", "В реестре нет позиций остатков.")
        return issues
    for index, item in enumerate(items):
        prefix = f"Позиция {index + 1}: "
        if not isinstance(item, dict):
            add("item", prefix + "данные позиции повреждены.", index)
            continue
        if _missing_label(item.get("name")):
            add("name", prefix + "не указано наименование.", index)
        category, group = item.get("cat"), item.get("grp")
        if _missing_label(category):
            add("cat", prefix + "не выбрана категория.", index)
        elif category not in current["categories"]:
            add("cat", prefix + "категория отсутствует в справочнике.", index)
        else:
            groups = current["groupsByCat"].get(category, [])
            if groups and _missing_label(group):
                add("grp", prefix + "не выбрана группа.", index)
            elif group and group not in groups:
                add("grp", prefix + "группа не принадлежит выбранной категории.", index)
        _, unknown = norm_unit(item.get("unit"))
        if unknown or item.get("unitUnknown"):
            add("unit", prefix + "не указана или не распознана единица измерения.", index)
        try:
            quantity = _quantity(item.get("qty"))
            if quantity is None:
                add("qty", prefix + "не указано количество.", index)
        except ValueError:
            add("qty", prefix + "количество должно быть конечным числом не меньше нуля.", index)
    return issues


def _record_revision(record):
    data = {key: value for key, value in record.items() if key not in {"revision", "approval_issues"}}
    return hashlib.sha256(json.dumps(data, sort_keys=True, ensure_ascii=False, default=str).encode("utf-8")).hexdigest()


def _with_approval_status(state):
    current = dictionary.load_dictionary()
    for record in state.get("pending", []):
        if not isinstance(record, dict):
            continue
        record["approval_issues"] = approval_issues(record, current)
        record["revision"] = _record_revision(record)
    return state


def _pending_record(state, index, expected_revision=None):
    pending = state.get("pending", [])
    if not isinstance(index, int) or isinstance(index, bool) or index < 0 or index >= len(pending):
        if expected_revision is not None:
            raise CuratorConflictError("Очередь изменилась. Проверьте актуальные данные и повторите действие.")
        return None
    record = pending[index]
    if expected_revision is not None and expected_revision != _record_revision(record):
        raise CuratorConflictError("Реестр уже изменён в другой вкладке или при обновлении. Проверьте актуальные данные и повторите действие.")
    return record


def _restore_pending_units(state):
    """Repair old imports from retained raw values without reclassifying items."""
    changed = False
    for registry in state.get("pending", []):
        for item in registry.get("items", []):
            if item.get("unit") and not item.get("unitUnknown"):
                continue
            unit, unknown = norm_unit(item.get("unitRaw"))
            if not unknown:
                item.update(unit=unit, unitUnknown=False)
                changed = True
    return changed

def classify(name):
    return dictionary.lookup(norm(name))

def parse_reestr(rows, fname):
    hi = -1
    for i, r in enumerate(rows[:15]):
        low = [str(x or "").lower() for x in (r or [])]
        if any("наимен" in c for c in low) and any("кол" in c for c in low): hi = i; break
    rso, date = "", ""
    for i in range(0, hi if hi >= 0 else 3):
        r = rows[i] if i < len(rows) else []
        cell = r[0] if r else None
        if cell is None: continue
        s = str(cell).strip()
        if not s: continue
        dm = re.search(r"(\d{4})-(\d{2})-(\d{2})", s) or re.search(r"(\d{2})\.(\d{2})\.(\d{4})", s)
        if dm and not date:
            a,b,c = dm.groups()
            date = f"{c}.{b}.{a}" if "-" in dm.group(0) else dm.group(0)
        elif not rso: rso = s
    rso_explicit = bool(rso)
    if not rso: rso = re.sub(r"\.(xlsx|xls)$", "", fname or "РСО", flags=re.I)
    if hi < 0: hi = 0
    items = []
    for r in rows[hi+1:]:
        r = r or []
        if all(x in (None, "") for x in r[:3]): continue
        name = str(r[0] or "").strip()
        if not name: continue
        q = r[2] if len(r) > 2 else None
        if isinstance(q, str):
            try: q = float(q.replace(",", "."))
            except Exception: q = None
        items.append({"name": name, "unitRaw": str(r[1] or "").strip() if len(r) > 1 else "", "qty": q})
    return {"rso": rso, "rso_explicit": rso_explicit, "date": date, "items": items}

def _atomic_write_json(path, payload):
    path.parent.mkdir(parents=True, exist_ok=True)
    _atomic_write_bytes(path, json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"))


def _clean_from_rows(rows):
    if len(rows) < 2:
        return {}
    header = [norm(v) for v in rows[0]]

    def column(fragment):
        return next((i for i, value in enumerate(header) if fragment in value), -1)

    indexes = {
        "rso": column("рсо"), "okrug": column("округ"), "name": column("наимен"),
        "cat": column("категор"), "grp": column("групп"), "qty": column("кол"),
        "unit": column("ед"), "water": column("вод"), "date": column("дата"),
    }
    if indexes["rso"] < 0 or indexes["name"] < 0:
        return {}

    clean = {}
    for row in rows[1:]:
        if not isinstance(row, list):
            continue
        value = lambda key, default="": row[indexes[key]] if 0 <= indexes[key] < len(row) else default
        rso = str(value("rso") or "").strip()
        name = str(value("name") or "").strip()
        if not rso or not name:
            continue
        key = norm(rso)
        entry = clean.setdefault(key, {
            "rso": rso,
            "okrug": str(value("okrug") or "").strip(),
            "date": str(value("date") or "").strip(),
            "items": [],
        })
        qty = value("qty", None)
        try:
            qty = float(str(qty).replace(",", ".")) if qty not in (None, "") else None
        except (TypeError, ValueError):
            qty = None
        item_norm = norm(name)
        entry["items"].append({
            "name": name,
            "nn": item_norm,
            "cat": str(value("cat") or "").strip() or None,
            "grp": str(value("grp") or "").strip() or None,
            "qty": qty,
            "unit": str(value("unit") or "").strip() or None,
            "unitRaw": str(value("unit") or "").strip(),
            "unitUnknown": not bool(str(value("unit") or "").strip()),
            "water": bool(str(value("water") or "").strip()),
            "via": "published",
        })
    return clean


def _clean_from_published():
    """Restore the approved registry when upgrading from published.json-only storage."""
    if not PUBLISHED_FILE.exists():
        return {}
    try:
        payload = json.loads(PUBLISHED_FILE.read_text(encoding="utf-8"))
        rows = payload.get("rows") or []
    except (OSError, json.JSONDecodeError, AttributeError):
        return {}
    clean = _clean_from_rows(rows)
    for key, entry in clean.items():
        source = (payload.get("rso_metadata") or {}).get(key) or {}
        entry.update({field: source[field] for field in SOURCE_FIELDS if field in source})
    return clean


def _source_timestamp(value):
    """Accept source-provided ISO datetimes only; local file timestamps are not evidence."""
    if not isinstance(value, str) or not re.fullmatch(
            r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?", value):
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    return parsed.isoformat(timespec="seconds")


def _source_identity(provider, source, account=None, folder_path=None):
    # Different accounts can expose identical URLs/paths; never share their cache.
    # Persist only this digest, never the login or password.
    owner = str((account or {}).get("username") or "") if provider == "mosreg" else ""
    scope = [provider, str(source), owner, str(folder_path or "")]
    return hashlib.sha256(json.dumps(scope, ensure_ascii=False).encode("utf-8")).hexdigest()


def _manifest_sources():
    """Read legacy single-source manifests and the source-scoped registry together."""
    try:
        saved = json.loads(FOLDER_A_MANIFEST_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return {}
    if not isinstance(saved, dict):
        return {}
    candidates = list((saved.get("sources") or {}).values()) if isinstance(saved.get("sources"), dict) else []
    # The top-level copy stays compatible with legacy tools and upgrades. It wins
    # over the same registry entry if a previous release updated just this copy.
    candidates.append(saved)
    sources = {}
    for entry in candidates:
        if not isinstance(entry, dict) or not entry.get("source") or not isinstance(entry.get("files"), list):
            continue
        provider = entry.get("provider") or "yandex"
        identity = entry.get("source_identity")
        if not identity:
            if provider != "yandex":
                continue
            identity = _source_identity(provider, entry["source"])
        sources[identity] = {key: value for key, value in entry.items() if key not in {"sources", "version"}}
        sources[identity].update(provider=provider, source_identity=identity)
    return sources


def _valid_cache_name(name):
    return isinstance(name, str) and name not in {"", ".", ".."} and "/" not in name and "\\" not in name


def _manifest_source_metadata():
    """Index unambiguous cache ownership across every source, including legacy data."""
    indexed, duplicates = {}, set()
    for identity, saved in _manifest_sources().items():
        for item in saved["files"]:
            if not isinstance(item, dict):
                continue
            name = item.get("local_name")
            if not _valid_cache_name(name) or not item.get("remote_path"):
                continue
            try:
                local_file = (FOLDER_A_DIR / name).relative_to(INPUT_DIR).as_posix()
            except ValueError:
                continue
            info = {"source_provider": saved["provider"], "source_identity": identity,
                    "source_path": item["remote_path"]}
            for field, origin in [("source_created_at", "created"), ("source_modified_at", "modified")]:
                value = _source_timestamp(item.get(origin))
                if value:
                    info[field] = value
            if item.get("md5"):
                info["source_md5"] = str(item["md5"]).lower()
            if item.get("sha256"):
                info["source_sha256"] = str(item["sha256"])
            if item.get("etag"):
                info["source_etag"] = str(item["etag"])
            if type(item.get("size")) is int and item["size"] >= 0:
                info["source_size"] = item["size"]
            if local_file in indexed and indexed[local_file] != info:
                duplicates.add(local_file)
            indexed[local_file] = info
    return {key: value for key, value in indexed.items() if key not in duplicates}


def _record_fingerprint(record):
    """Original rows only: never hash classifications or curator corrections."""
    # Without an explicit organization, identical inventory rows may belong to
    # unrelated organizations. Only byte-identical workbooks are safe to merge.
    if not record.get("rso_explicit"):
        return None
    rows = [[re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(item.get("name") or "")).casefold()).strip(),
             _unit_key(item.get("unitRaw") or ""), item.get("qty")]
            for item in record.get("items", [])]
    rows.sort(key=lambda row: json.dumps(row, ensure_ascii=False, sort_keys=True))
    payload = [norm(record.get("rso")), rows]
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def _matching_fingerprint(first, second):
    return any(first.get(field) and first.get(field) == second.get(field) for field in FINGERPRINT_FIELDS)


def _source_date_value(record):
    for field in ("source_modified_at", "source_created_at"):
        value = _source_timestamp(record.get(field))
        if value:
            try:
                return datetime.fromisoformat(value).timestamp()
            except (ValueError, OverflowError):
                pass
    return None


def _older_source(candidate, existing):
    """A healthy old source cannot roll back a newer source that is unavailable."""
    if not candidate.get("source_identity") or not existing.get("source_identity"):
        return False
    incoming, previous = _source_date_value(candidate), _source_date_value(existing)
    same_origin = (candidate.get("source_identity") == existing.get("source_identity")
                   and candidate.get("source_path") == existing.get("source_path"))
    if previous is not None and incoming is not None:
        return incoming < previous or (incoming == previous and not same_origin)
    # Unknown dates are not evidence of a newer version from a different source.
    return not same_origin and incoming is None


def _same_source_version(first, second):
    if not first.get("local_file") or first.get("local_file") != second.get("local_file"):
        return False
    if first.get("source_identity") or second.get("source_identity"):
        if (not first.get("source_identity") or first.get("source_identity") != second.get("source_identity")
                or first.get("source_path") != second.get("source_path")):
            return False
        if first.get("source_md5") and second.get("source_md5"):
            return first["source_md5"] == second["source_md5"]
        return bool(first.get("source_etag") and first.get("source_etag") == second.get("source_etag")
                    and type(first.get("source_size")) is int
                    and first.get("source_size") == second.get("source_size"))
    return bool(first.get("uploaded") and first.get("uploaded") == second.get("uploaded"))


def _verified_cached_original(entry, info):
    """Verify an old cache version before deriving immutable migration fingerprints."""
    relative = entry.get("local_file")
    if not isinstance(relative, str):
        return None
    path = INPUT_DIR / relative
    try:
        if path.is_symlink() or not path.resolve().is_relative_to(INPUT_DIR.resolve()):
            return None
        if path.stat().st_size > MAX_FILE_BYTES:
            return None
        data = path.read_bytes()
        expected_hash = info.get("source_sha256")
        if expected_hash and hashlib.sha256(data).hexdigest() != expected_hash:
            return None
        if not expected_hash and not info.get("source_md5"):
            # Old cloud manifests use opaque ETags. Require the same trusted
            # ETag/size pairing already used by the cache reader, plus XLSX validation.
            if not (info.get("source_etag") and type(info.get("source_size")) is int
                    and entry.get("source_etag") == info["source_etag"]
                    and entry.get("source_size") == info["source_size"]):
                return None
        validate_workbook(data, {"md5": info.get("source_md5"), "size": info.get("source_size")}, read_xlsx_rows)
        return data
    except Exception:
        return None


def _backfill_source_metadata(state):
    metadata = _manifest_source_metadata()
    changed = False
    for entry in [*state.get("pending", []), *state.get("clean", {}).values()]:
        info = metadata.get(entry.get("local_file"))
        if not info:
            continue
        # Existing source/version evidence must agree. A newer cache must never
        # change the dates attached to an older approved workbook.
        if any(entry.get(field) and entry.get(field) != info.get(field)
               for field in ("source_identity", "source_path", "source_md5", "source_etag", "source_size", "source_sha256")):
            continue
        if not entry.get("source_identity") and entry.get("uploaded"):
            known = info.get("source_modified_at") or info.get("source_created_at")
            if known:
                try:
                    if abs(float(entry["uploaded"]) - datetime.fromisoformat(known).timestamp() * 1000) > 1:
                        continue
                except (ValueError, TypeError, OverflowError):
                    continue
        if not entry.get("source_sha256") or ("source_record_hash" not in entry and entry.get("rso_explicit") is not False):
            original = _verified_cached_original(entry, info)
            if original is not None:
                try:
                    parsed = parse_reestr(read_xlsx_rows(original), entry.get("fname") or "registry.xlsx")
                    # A curator may have corrected names/quantities in state. Only
                    # the verified original workbook supplies this baseline.
                    info = {**info, "source_sha256": hashlib.sha256(original).hexdigest(),
                            "rso_explicit": parsed.get("rso_explicit", False)}
                    fingerprint = _record_fingerprint(parsed)
                    if fingerprint:
                        info["source_record_hash"] = fingerprint
                except Exception:
                    pass
        for field, value in info.items():
            existing = entry.get(field)
            if existing == value:
                continue
            if not existing or (field in SOURCE_DATE_FIELDS and not _source_timestamp(existing)):
                entry[field] = value
                changed = True
    return changed


def _persist_published_source_metadata(state):
    """Backfill metadata without republishing rows or changing the publication time."""
    try:
        payload = json.loads(PUBLISHED_FILE.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            return
        published_keys = set(_clean_from_rows(payload.get("rows") or []))
        metadata = payload.setdefault("rso_metadata", {})
        if not isinstance(metadata, dict):
            return
        changed = False
        for key, entry in state.get("clean", {}).items():
            if key not in published_keys:
                continue
            fields = {field: entry[field] for field in SOURCE_FIELDS if field in entry}
            if metadata.get(key) != fields:
                metadata[key] = fields
                changed = True
        if changed:
            _atomic_write_json(PUBLISHED_FILE, payload)
    except (OSError, ValueError, TypeError, AttributeError):
        pass


def load_state():
    if STATE_FILE.exists():
        try:
            state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            if isinstance(state, dict):
                state.setdefault("pending", [])
                state.setdefault("clean", {})
                state.setdefault("excluded_rso", {})
                units_changed = _restore_pending_units(state)
                metadata_changed = _backfill_source_metadata(state)
                if units_changed or metadata_changed:
                    save_state(state)
                if metadata_changed:
                    _persist_published_source_metadata(state)
        except Exception:
            pass
        else:
            if isinstance(state, dict):
                return _with_approval_status(state)
    state = {"pending": [], "clean": _clean_from_published(), "excluded_rso": {}}
    metadata_changed = _backfill_source_metadata(state)
    if state["clean"]:
        save_state(state)
    if metadata_changed:
        _persist_published_source_metadata(state)
    return _with_approval_status(state)

def save_state(st):
    # Derived validation and revision are always recalculated, never trusted from disk.
    payload = {**st, "pending": [
        {key: value for key, value in record.items() if key not in {"revision", "approval_issues"}}
        for record in st.get("pending", [])
    ]}
    _atomic_write_json(STATE_FILE, payload)
    _with_approval_status(st)

def _enrich(p):
    for it in p["items"]:
        c = classify(it["name"]); u, unk = norm_unit(it.get("unitRaw"))
        it.update(nn=norm(it["name"]), cat=c["cat"], grp=c["grp"], water=c["water"], via=c["via"], unit=u, unitUnknown=unk)
    p["okrug"] = resolve_municipality(p["rso"])

@_curator_mutation_lock()
def ingest(list_of_parsed):
    st = load_state()
    added = 0
    for p in list_of_parsed:
        if not p["items"]: continue
        _enrich(p)
        previous = next((x for x in st["pending"] if norm(x["rso"]) == norm(p["rso"])), {})
        # Match renamed copies by the immutable source fingerprints, not the
        # mutable item rows that a curator may already have corrected.
        identical = next((x for x in st["pending"] if _matching_fingerprint(p, x)), None)
        if identical is not None:
            incoming_date, previous_date = _source_date_value(p), _source_date_value(identical)
            same_origin = (p.get("source_identity") and p.get("source_identity") == identical.get("source_identity")
                           and p.get("source_path") == identical.get("source_path"))
            newer = (incoming_date is not None and (previous_date is None or incoming_date > previous_date)
                     or (same_origin and incoming_date is None and previous_date is None))
            if _same_source_version(p, identical):
                identical.update({field: p[field] for field in (*SOURCE_INFO_FIELDS, *FINGERPRINT_FIELDS) if field in p})
            elif newer:
                # Equivalent inventory may arrive with a newer report date or
                # provenance. Advance freshness without losing curator edits.
                for field in (*SOURCE_INFO_FIELDS, *FINGERPRINT_FIELDS):
                    identical.pop(field, None)
                identical.update({field: p[field] for field in SOURCE_FIELDS if field in p and field != "loaded_at"})
                identical["date"] = p.get("date") or ""
            continue
        unchanged = _same_source_version(p, previous)
        if unchanged:
            previous.update({field: p[field] for field in (*SOURCE_INFO_FIELDS, *FINGERPRINT_FIELDS) if field in p})
            continue
        if previous and _older_source(p, previous):
            continue
        p["loaded_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
        st["pending"] = [x for x in st["pending"] if norm(x["rso"]) != norm(p["rso"])]
        st["pending"].append(p); added += 1
    save_state(st)
    return added


@_curator_mutation_lock()
def scan_folder(scan_dir=None, files=None, *, with_summary=False):
    scan_dir = Path(scan_dir or INPUT_DIR)
    scan_dir.mkdir(parents=True, exist_ok=True)
    state = load_state()
    approved = set(state.get("clean", {}))
    excluded = set(state.get("excluded_rso", {}))
    protected = [*state.get("clean", {}).values(), *state.get("excluded_rso", {}).values()]
    pending_before = len(state.get("pending", []))
    state["pending"] = [
        item for item in state.get("pending", [])
        if norm(item.get("rso")) not in (approved | excluded)
    ]
    if len(state["pending"]) != pending_before:
        save_state(state)
    files = [Path(p) for p in (scan_dir.rglob("*") if files is None else files)
             if Path(p).is_file() and not Path(p).is_symlink() and re.search(r"\.xlsx?$", Path(p).name, re.I)]
    source_metadata = _manifest_source_metadata()

    def local_name(path):
        try:
            return path.relative_to(INPUT_DIR).as_posix()
        except ValueError:
            return path.name

    def priority(path):
        info = source_metadata.get(local_name(path), {})
        date = _source_date_value(info)
        # Download time is never a ranking signal for remote workbooks.
        return (date is not None, date if date is not None else
                (0 if info.get("source_identity") else path.stat().st_mtime), local_name(path))

    files.sort(key=priority, reverse=True)
    taken, skipped, duplicates = {}, 0, 0
    byte_hashes, record_hashes = set(), set()
    for f in files:
        try:
            data = f.read_bytes()
            digest = hashlib.sha256(data).hexdigest()
            if digest in byte_hashes:
                skipped += 1; duplicates += 1
                continue
            byte_hashes.add(digest)
            rows = read_xlsx_rows(data)
        except Exception:
            continue
        try:
            p = parse_reestr(rows, f.name)
        except Exception:
            continue
        if not p["items"]: continue
        p["source_sha256"] = digest
        fingerprint = _record_fingerprint(p)
        if fingerprint:
            p["source_record_hash"] = fingerprint
            if fingerprint in record_hashes:
                skipped += 1; duplicates += 1
                continue
            record_hashes.add(fingerprint)
        key = norm(p["rso"])
        # Approved/excluded data remains protected even when a byte-identical
        # workbook gets a different filename in the second source.
        if key in approved or key in excluded or any(_matching_fingerprint(p, row) for row in protected):
            skipped += 1
            continue
        if key in taken:
            skipped += 1; duplicates += 1
            continue
        p["fname"] = f.name
        p["uploaded"] = int(f.stat().st_mtime * 1000)
        p["local_file"] = local_name(f)
        p.update(source_metadata.get(p["local_file"], {}))
        taken[key] = p
    added = ingest(list(taken.values()))
    return (added, skipped, duplicates) if with_summary else (added, skipped)


def _safe_remote_name(item, used):
    original = Path(str(item.get("name") or "registry.xlsx")).name
    cleaned = re.sub(r"[^0-9A-Za-zА-Яа-яЁё._() -]+", "_", original).strip(" .") or "registry.xlsx"
    candidate = cleaned
    remote_path = str(item.get("path") or original)
    if candidate.casefold() in used:
        stem, suffix = Path(cleaned).stem, Path(cleaned).suffix
        candidate = f"{stem}-{hashlib.sha1(remote_path.encode('utf-8')).hexdigest()[:8]}{suffix}"
        base = Path(candidate).stem
        counter = 2
        while candidate.casefold() in used:
            candidate = f"{base}-{counter}{suffix}"
            counter += 1
    used.add(candidate.casefold())
    return candidate


class FolderSyncBusy(DiskSyncError):
    pass


@contextmanager
def _folder_sync_lock():
    """One update across threads/processes, including Windows uvicorn workers."""
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with (STATE_FILE.parent / "folder_a.lock").open("a+b") as handle:
        handle.seek(0, os.SEEK_END)
        if not handle.tell():
            handle.write(b"0")
            handle.flush()
        handle.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            raise FolderSyncBusy("Папка А уже обновляется. Дождитесь окончания текущего обновления.") from None
        try:
            yield
        finally:
            handle.seek(0)
            if os.name == "nt":
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def _atomic_write_bytes(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".zip-sync-", suffix=".part", delete=False) as handle:
            tmp = Path(handle.name)
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp, path)
    finally:
        if tmp is not None:
            tmp.unlink(missing_ok=True)


def _cached_workbook(path, metadata):
    if not metadata or not path.is_file() or path.is_symlink():
        return False
    try:
        if path.stat().st_size > MAX_FILE_BYTES:
            return False
        validate_workbook(path.read_bytes(), metadata, read_xlsx_rows)
        return True
    except (OSError, DiskSyncError):
        return False


def _sync_folder_a(public_url, timeout=45, *, account=None, provider="yandex", folder_path=None):
    public_url = str(public_url or "").strip()
    if not public_url:
        raise DiskSyncError("Не указана ссылка на папку А.")
    if provider not in {"yandex", "mosreg"}:
        raise DiskSyncError("Неизвестный источник папки А.")
    if provider == "mosreg" and (not isinstance(account, dict) or not account.get("username") or not account.get("password")):
        raise DiskSyncError("Для Облака МО не настроены логин и пароль.")
    identity = _source_identity(provider, public_url, account, folder_path)
    with requests.Session() as session:
        if provider == "mosreg":
            from services.zip_curator.mosreg_cloud import CloudClient
            client = CloudClient(session, public_url, timeout, username=account["username"],
                                 password=account["password"], folder_path=folder_path)
        else:
            client = DiskClient(session, public_url, timeout)
        files = [item for item in client.list_files()
                 if item["type"] == "file" and re.search(r"\.xlsx?$", item["name"], re.I)]
        files.sort(key=lambda item: item["path"])
        FOLDER_A_DIR.mkdir(parents=True, exist_ok=True)
        # Attach OLD manifest dates before a newer file version replaces that manifest.
        load_state()
        sources = _manifest_sources()
        saved = sources.get(identity) or {}
        previous = {item["remote_path"]: item for item in saved.get("files", [])
                    if isinstance(item, dict) and item.get("remote_path")}

        # Reserve known names first so a new file cannot overwrite another file's cache.
        used, names = set(), {}
        for item in files:
            name = (previous.get(item["path"]) or {}).get("local_name")
            if (isinstance(name, str) and name not in {"", ".", ".."}
                    and "/" not in name and "\\" not in name and name.casefold() not in used):
                names[item["path"]] = name
                used.add(name.casefold())
        # Changing source/account must not overwrite another source's cached originals.
        owned_names = {item.get("local_name") for item in previous.values()}
        used.update(path.name.casefold() for path in FOLDER_A_DIR.iterdir() if path.name not in owned_names)
        for item in files:
            if item["path"] not in names:
                names[item["path"]] = _safe_remote_name(item, used)

        manifest, errors, ready_files, downloaded, reused = [], [], [], 0, 0
        for item in files:
            remote_path = item["path"]
            local_name = names[remote_path]
            local_path = FOLDER_A_DIR / local_name
            old = previous.get(remote_path) or {}
            cached = _cached_workbook(local_path, old)
            checksum = str(item.get("md5") or "").lower()
            etag = str(item.get("etag") or "")
            same_checksum = (checksum and checksum == str(old.get("md5") or "").lower()
                             and (item.get("size") is None or item["size"] == old.get("size")))
            same_etag = (provider == "mosreg" and etag and etag == str(old.get("etag") or "")
                         and type(item.get("size")) is int and item["size"] == old.get("size"))
            unchanged = cached and (same_checksum if checksum else same_etag)
            try:
                if unchanged:
                    reused += 1
                else:
                    data = client.download(item, read_xlsx_rows)
                    _atomic_write_bytes(local_path, data)
                    downloaded += 1
            except (DiskSyncError, OSError) as exc:
                message = _safe_sync_error(exc, account)
                errors.append({"name": item["name"], "message": message, "retained": bool(cached)})
                if cached:
                    # Keep the OLD checksum and timestamp: the next run must retry this file.
                    manifest.append(old)
                continue
            created = _source_timestamp(item.get("created"))
            modified = _source_timestamp(item.get("modified"))
            sorting_date = modified or created
            if sorting_date:
                try:
                    timestamp = datetime.fromisoformat(sorting_date).timestamp()
                    os.utime(local_path, (timestamp, timestamp))
                except (OSError, ValueError):
                    pass
            manifest.append({"name": item["name"], "local_name": local_name,
                             "remote_path": remote_path, "md5": checksum, "etag": etag,
                             "sha256": hashlib.sha256(local_path.read_bytes()).hexdigest(),
                             "created": created, "modified": modified, "size": item.get("size")})
            ready_files.append(local_path)

    if errors:
        # Retain ownership of caches absent from this listing until a complete run.
        listed_paths = {item["path"] for item in files}
        manifest.extend(old for remote_path, old in previous.items() if remote_path not in listed_paths)

    current = {
        "source": public_url, "provider": provider, "source_identity": identity,
        "synced_at": datetime.now().isoformat(timespec="seconds"),
        "partial": bool(errors), "errors": errors, "files": manifest,
    }
    sources[identity] = current
    _atomic_write_json(FOLDER_A_MANIFEST_FILE, {**current, "version": 2, "sources": sources})
    # A partial/invalid listing must never act as a request to delete the cache.
    active_names = {item.get("local_name") for source in sources.values() for item in source.get("files", [])
                    if isinstance(item, dict)}
    removed = 0
    if not errors:
        for old in previous.values():
            name = old.get("local_name")
            if isinstance(name, str) and Path(name).name == name and "\\" not in name and name not in active_names:
                path = FOLDER_A_DIR / name
                if path.is_file() and not path.is_symlink():
                    try:
                        path.unlink()
                        removed += 1
                    except OSError:
                        pass  # An unused cache file must not invalidate a completed update.
    result = {"ok": True, "partial": bool(errors), "files": len(files), "downloaded": downloaded,
              "reused": reused, "removed": removed, "failed": len(errors), "errors": errors}
    # A failed source must not re-import an old workbook over the curator's edits.
    return result, ready_files


def sync_folder_a(public_url, timeout=45, *, account=None, provider="yandex", folder_path=None):
    """Update the cache; leave publication and the review queue untouched."""
    with _folder_sync_lock(), _owned_curator_mutation():
        result, _ = _sync_folder_a(public_url, timeout, account=account, provider=provider, folder_path=folder_path)
        return result


def sync_and_scan_folder_a(public_url, *, account=None, provider="yandex", folder_path=None):
    with _folder_sync_lock(), _owned_curator_mutation():
        result, files = _sync_folder_a(public_url, account=account, provider=provider, folder_path=folder_path)
        added, skipped = scan_folder(FOLDER_A_DIR, files=files)
        result.update({"added": added, "skipped": skipped, "state": load_state()})
        return result


_SOURCE_LABELS = {"yandex": "Яндекс Диск", "mosreg": "Облако МО"}


def _safe_sync_error(error, account=None):
    if isinstance(error, DiskSyncError):
        message = str(error)
    elif isinstance(error, OSError) and not isinstance(error, requests.RequestException):
        message = "Не удалось сохранить файл на сервере. Проверьте свободное место и права на папку данных."
    else:
        message = "Не удалось обновить источник. Проверьте доступ и настройки подключения."
    # Exception strings must not expose signed download URLs or credentials.
    message = re.sub(r"https?://[^\s]+", "[адрес источника]", message)
    for field in ("username", "password"):
        secret = str((account or {}).get(field) or "")
        if secret:
            message = message.replace(secret, "[скрыто]")
    return message[:700]


def sync_and_scan_sources(sources):
    """Refresh independent sources under one lock, then ingest their fresh union."""
    with _folder_sync_lock(), _owned_curator_mutation():
        result = {"ok": False, "partial": False, "files": 0, "downloaded": 0,
                  "reused": 0, "removed": 0, "failed": 0, "errors": [], "sources": [],
                  "duplicates": 0, "added": 0, "skipped": 0, "failed_sources": 0}
        ready_files = []
        any_success = False
        for source in sources:
            provider = source.get("provider")
            label = _SOURCE_LABELS.get(provider, "Источник")
            summary = {"provider": provider, "label": label, "ok": False, "partial": False,
                       "files": 0, "downloaded": 0, "reused": 0, "failed": 0}
            try:
                if source.get("error"):
                    raise DiskSyncError(str(source["error"]))
                synced, files = _sync_folder_a(source.get("url"), provider=provider,
                    account=source.get("account"), folder_path=source.get("folder_path"))
                usable = bool(files) or not synced["failed"]
                summary.update({key: synced[key] for key in ("partial", "files", "downloaded", "reused", "failed")})
                summary["ok"] = usable
                if not usable:
                    summary["error"] = "Ни один файл источника не удалось обновить. Сохранённые данные оставлены."
                any_success = any_success or usable
                ready_files.extend(files)
                for field in ("files", "downloaded", "reused", "removed", "failed"):
                    result[field] += synced[field]
                result["partial"] = result["partial"] or synced["partial"]
                result["errors"].extend({**error, "provider": provider, "source": label} for error in synced["errors"])
            except Exception as exc:
                message = _safe_sync_error(exc, source.get("account"))
                summary.update(error=message, failed=1)
                result["failed_sources"] += 1
                result["partial"] = True
                result["errors"].append({"name": label, "provider": provider, "source": label,
                                         "message": message, "retained": True})
            result["sources"].append(summary)
        result["ok"] = any_success
        if any_success:
            added, skipped, duplicates = scan_folder(FOLDER_A_DIR, files=ready_files, with_summary=True)
            result.update(added=added, skipped=skipped, duplicates=duplicates)
        if not any_success:
            result["error"] = "Не удалось обновить ни один источник. Ранее загруженные данные сохранены."
            # Do not mutate legacy state during an unsuccessful refresh, even
            # through a metadata migration in load_state().
            try:
                state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
                result["state"] = state if isinstance(state, dict) else {"pending": [], "clean": {}, "excluded_rso": {}}
            except (OSError, ValueError, TypeError):
                result["state"] = {"pending": [], "clean": _clean_from_published(), "excluded_rso": {}}
        else:
            result["state"] = load_state()
        return result

def approve(idx, *, expected_revision=None):
    with _curator_mutation_lock():
        st = load_state()
        p = _pending_record(st, idx, expected_revision)
        if p is None:
            return False
        issues = approval_issues(p)
        if issues:
            raise ApprovalValidationError(issues)
        # Do not combine different units/classifications under the same name.
        grouped = {}
        for item in p["items"]:
            canonical_unit, _ = norm_unit(item["unit"])
            key = (norm(item["name"]), item["cat"], item.get("grp"), canonical_unit, bool(item.get("water")))
            if key not in grouped:
                grouped[key] = {**item, "nn": key[0], "qty": _quantity(item["qty"]),
                                "unit": canonical_unit, "unitUnknown": False}
            else:
                grouped[key]["qty"] += _quantity(item["qty"])
                if not math.isfinite(grouped[key]["qty"]):
                    raise ApprovalValidationError([{"field": "qty", "message": "Сумма количества одинаковых позиций слишком велика."}])
        key = norm(p["rso"])
        municipality = p.get("okrug")
        if _missing_label(municipality):
            municipality = resolve_municipality(p["rso"])
        st["clean"][key] = {"rso": p["rso"], "okrug": municipality, "date": p.get("date"), "items": list(grouped.values())}
        st["clean"][key].update({field: p[field] for field in SOURCE_FIELDS if field in p})
        st["pending"].pop(idx)
        st.setdefault("excluded_rso", {}).pop(key, None)
        # Publish first: a failed publication must leave the pending registry available.
        publish_clean(st)
        return True

@_curator_mutation_lock()
def reject(idx):
    st = load_state()
    if idx < 0 or idx >= len(st["pending"]): return False
    st["pending"].pop(idx); save_state(st); return True


@_curator_mutation_lock()
def delete_rso(rso):
    """Delete one RSO everywhere and keep its Folder A file from restoring it."""
    rso = str(rso or "").strip()
    key = norm(rso)
    if not key:
        raise ValueError("Не указано РСО")

    st = load_state()
    pending_before = len(st["pending"])
    deleted_source = next((item for item in st["pending"] if norm(item.get("rso")) == key), {})
    st["pending"] = [item for item in st["pending"] if norm(item.get("rso")) != key]
    pending_removed = pending_before - len(st["pending"])
    clean_item = st["clean"].pop(key, None)
    display_name = str((clean_item or {}).get("rso") or rso).strip()
    st.setdefault("excluded_rso", {})[key] = {
        "rso": display_name,
        **{field: (clean_item or deleted_source)[field] for field in FINGERPRINT_FIELDS
           if field in (clean_item or deleted_source)},
        "deleted_at": datetime.now().isoformat(timespec="seconds"),
    }

    # Состояние куратора является полным набором опубликованных РСО. Полная запись
    # здесь нужна специально: обычная публикация обновляет РСО по одному и сохраняет
    # остальные, поэтому не может выразить удаление.
    rows = export_rows_from_state(st)
    _atomic_write_json(PUBLISHED_FILE, {
        "rows": rows,
        "published_at": datetime.now().isoformat(timespec="seconds"),
        "rso_metadata": _source_metadata(st["clean"]),
    })
    save_state(st)
    return {
        "removed": bool(pending_removed or clean_item),
        "pending_removed": pending_removed,
        "published_removed": bool(clean_item),
        "rso": display_name,
    }

def edit_item(pi, ii, cat, grp, *, unit=_UNSET, qty=_UNSET, expected_revision=None):
    with _curator_mutation_lock():
        st = load_state()
        record = _pending_record(st, pi, expected_revision)
        if record is None or not isinstance(ii, int) or isinstance(ii, bool) or ii < 0:
            return False
        try:
            it = record["items"][ii]
        except (IndexError, KeyError):
            return False
        updates = {}
        # Validate every requested value before writing either state or dictionary.
        if unit is not _UNSET:
            if unit is not None and not isinstance(unit, str):
                raise ValueError("Единица измерения должна быть текстом")
            canonical, unknown = norm_unit(unit)
            if unit and str(unit).strip() and unknown:
                raise ValueError("Выберите единицу измерения из списка, например «пог. м»")
            updates.update(unit=canonical, unitRaw=canonical or "", unitUnknown=unknown)
        if qty is not _UNSET:
            updates["qty"] = _quantity(qty)
        if cat:
            cat, grp = dictionary.remember(norm(it["name"]), cat, grp, it.get("water"))
        elif grp:
            raise ValueError("Выберите категорию для группы")
        it.update(updates)
        it.update(cat=cat or None, grp=grp or None, via="override" if cat else None)
        save_state(st)
        return True

def export_rows():
    return export_rows_from_state(load_state())


def export_rows_from_state(st):
    rows = [["РСО","Округ","Наименование","Категория","Группа","Количество","Ед.","Водоподготовка","Дата"]]
    for c in st["clean"].values():
        for it in c["items"]:
            rows.append([c["rso"], c["okrug"], it["name"], it.get("cat") or "не указано", it.get("grp") or "", it.get("qty"), it.get("unit") or "", "да" if it.get("water") else "", c.get("date") or ""])
    return rows


def _published_rows():
    if not PUBLISHED_FILE.exists():
        return []
    try:
        payload = json.loads(PUBLISHED_FILE.read_text(encoding="utf-8"))
        rows = payload.get("rows") or []
        return rows if isinstance(rows, list) else []
    except (OSError, json.JSONDecodeError, AttributeError):
        return []


def merge_rows_by_rso(current_rows, incoming_rows):
    """Replace only RSO present in incoming rows and retain every other published RSO."""
    if len(incoming_rows) < 2:
        raise ValueError("Нет строк для публикации")
    incoming_header = incoming_rows[0]
    header_norm = [norm(value) for value in incoming_header]
    rso_index = next((i for i, value in enumerate(header_norm) if "рсо" in value), -1)
    if rso_index < 0:
        raise ValueError("В таблице отсутствует колонка РСО")

    def row_rso(row):
        return norm(row[rso_index]) if isinstance(row, list) and rso_index < len(row) else ""

    incoming_data = [row for row in incoming_rows[1:] if row_rso(row)]
    incoming_rso = {row_rso(row) for row in incoming_data}
    retained = []
    if len(current_rows) > 1:
        current_header = [norm(value) for value in current_rows[0]]
        current_rso_index = next((i for i, value in enumerate(current_header) if "рсо" in value), -1)
        if current_rso_index >= 0:
            for row in current_rows[1:]:
                key = norm(row[current_rso_index]) if isinstance(row, list) and current_rso_index < len(row) else ""
                if key and key not in incoming_rso:
                    retained.append(row)
    return [incoming_header, *retained, *incoming_data], incoming_rso


def _source_metadata(clean):
    return {key: {field: entry[field] for field in SOURCE_FIELDS if field in entry}
            for key, entry in clean.items()}


@_curator_mutation_lock()
def publish_rows(rows, st=None, *, validate=True):
    if validate:
        if not isinstance(rows, list) or len(rows) < 2 or not isinstance(rows[0], list):
            raise ValueError("Нет строк для публикации")
        incoming = _clean_from_rows(rows)
        parsed_count = sum(len(record["items"]) for record in incoming.values())
        if parsed_count != len(rows) - 1:
            raise ApprovalValidationError([{"field": "items", "message": "В каждой публикуемой позиции нужны название РСО и наименование."}])
        previous = _clean_from_published()
        issues = []
        classification = dictionary.load_dictionary()
        for key, record in incoming.items():
            old = previous.get(key, {})
            old_rows = export_rows_from_state({"clean": {key: old}}) if old else []
            new_rows = export_rows_from_state({"clean": {key: record}})
            # Do not invalidate untouched legacy data while updating another RSO.
            if old_rows != new_rows:
                issues.extend(approval_issues(record, classification))
            if _missing_label(record.get("okrug")):
                record["okrug"] = resolve_municipality(record["rso"])
        if issues:
            raise ApprovalValidationError(issues)
        rows = export_rows_from_state({"clean": incoming})
    merged_rows, incoming_rso = merge_rows_by_rso(_published_rows(), rows)
    state = st if st is not None else load_state()
    metadata = _source_metadata(state["clean"])
    clean = _clean_from_rows(merged_rows)
    for key, entry in clean.items():
        entry.update(metadata.get(key, {}))
    payload = {
        "rows": merged_rows,
        "published_at": datetime.now().isoformat(timespec="seconds"),
        "rso_metadata": _source_metadata(clean),
    }
    _atomic_write_json(PUBLISHED_FILE, payload)
    state["clean"] = clean
    save_state(state)
    payload["updated_rso"] = len(incoming_rso)
    payload["total_rso"] = len(state["clean"])
    return payload


def publish_clean(st=None):
    state = st or load_state()
    return publish_rows(export_rows_from_state(state), state, validate=False)


def resolve_source_file(file_path):
    candidate = (INPUT_DIR / str(file_path or "")).resolve()
    root = INPUT_DIR.resolve()
    if candidate == root or root not in candidate.parents or not candidate.is_file():
        return None
    return candidate

def write_xlsx(rows, path):
    def esc(s): return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    for r, row in enumerate(rows, 1):
        sheet.append(f'<row r="{r}">')
        for c, v in enumerate(row):
            ref = ""
            n = c
            while True:
                ref = chr(65 + n % 26) + ref
                n = n//26 - 1
                if n < 0:
                    break
            ref += str(r)
            if isinstance(v, (int, float)):
                sheet.append(f'<c r="{ref}"><v>{v}</v></c>')
            else:
                sheet.append(f'<c r="{ref}" t="inlineStr"><is><t>{esc(v)}</t></is></c>')
        sheet.append("</row>")
    sheet.append("</sheetData></worksheet>")
    sheet_xml = "".join(sheet)
    z = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    z.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Остатки" sheetId="1" r:id="rId1"/></sheets></workbook>')
    z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
    z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    z.close()
