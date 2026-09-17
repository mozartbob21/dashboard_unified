"""Куратор ЗиП: загрузка исходников, согласование и локальная публикация."""
import hashlib, io, json, os, re, zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import requests
from services.zip_curator import dictionary

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
SOURCE_FIELDS = ("loaded_at", "uploaded", "fname", "local_file")

D = json.loads(DICT_JSON.read_text(encoding="utf-8"))

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
        save_state(st)

def norm_unit(u):
    if u is None or not str(u).strip(): return None, True
    c = D["umap"].get(str(u).lower().strip())
    return c, not c

def classify(name):
    return dictionary.lookup(norm(name))

def parse_reestr(rows, fname):
    hi = -1
    for i, r in enumerate(rows[:15]):
        low = [str(x or "").lower() for x in (r or [])]
        if any("наимен" in c for c in low) and any("кол" in c for c in low): hi = i; break
    rso, date = "", ""
    for i in range(0, hi if hi > 0 else 3):
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
    return {"rso": rso, "date": date, "items": items}

def _atomic_write_json(path, payload):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


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


def load_state():
    if STATE_FILE.exists():
        try:
            state = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            if isinstance(state, dict):
                state.setdefault("pending", [])
                state.setdefault("clean", {})
                state.setdefault("excluded_rso", {})
                return state
        except Exception: pass
    state = {"pending": [], "clean": _clean_from_published(), "excluded_rso": {}}
    if state["clean"]:
        save_state(state)
    return state

def save_state(st):
    _atomic_write_json(STATE_FILE, st)

def _enrich(p):
    for it in p["items"]:
        c = classify(it["name"]); u, unk = norm_unit(it.get("unitRaw"))
        it.update(nn=norm(it["name"]), cat=c["cat"], grp=c["grp"], water=c["water"], via=c["via"], unit=u, unitUnknown=unk)
    p["okrug"] = resolve_municipality(p["rso"])

def ingest(list_of_parsed):
    st = load_state()
    added = 0
    for p in list_of_parsed:
        if not p["items"]: continue
        _enrich(p)
        previous = next((x for x in st["pending"] if norm(x["rso"]) == norm(p["rso"])), {})
        # Re-scanning an unchanged Folder A file is not a new upload.
        unchanged = (p.get("local_file") and p.get("uploaded")
                     and p.get("local_file") == previous.get("local_file")
                     and p.get("uploaded") == previous.get("uploaded"))
        p["loaded_at"] = (previous.get("loaded_at") if unchanged else None) or datetime.now().astimezone().isoformat(timespec="seconds")
        st["pending"] = [x for x in st["pending"] if norm(x["rso"]) != norm(p["rso"])]
        st["pending"].append(p); added += 1
    save_state(st)
    return added

def scan_folder(scan_dir=None):
    scan_dir = Path(scan_dir or INPUT_DIR)
    scan_dir.mkdir(parents=True, exist_ok=True)
    state = load_state()
    approved = set(state.get("clean", {}))
    excluded = set(state.get("excluded_rso", {}))
    pending_before = len(state.get("pending", []))
    state["pending"] = [
        item for item in state.get("pending", [])
        if norm(item.get("rso")) not in (approved | excluded)
    ]
    if len(state["pending"]) != pending_before:
        save_state(state)
    files = sorted(
        (p for p in scan_dir.rglob("*") if p.is_file() and re.search(r"\.xlsx?$", p.name, re.I)),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    taken, skipped = {}, 0
    for f in files:
        try: rows = read_xlsx_rows(f.read_bytes())
        except Exception: continue
        p = parse_reestr(rows, f.name)
        if not p["items"]: continue
        key = norm(p["rso"])
        # Согласованные РСО не возвращаем в очередь при каждом обновлении папки А.
        # Удалённые куратором РСО также не восстанавливаем из лежащего в папке файла.
        if key in approved or key in excluded:
            skipped += 1
            continue
        if key in taken: skipped += 1; continue
        p["fname"] = f.name
        p["uploaded"] = int(f.stat().st_mtime * 1000)
        try:
            p["local_file"] = f.relative_to(INPUT_DIR).as_posix()
        except ValueError:
            p["local_file"] = f.name
        taken[key] = p
    return ingest(list(taken.values())), skipped


def _safe_remote_name(item, used):
    original = Path(str(item.get("name") or "registry.xlsx")).name
    cleaned = re.sub(r"[^0-9A-Za-zА-Яа-яЁё._() -]+", "_", original).strip(" .") or "registry.xlsx"
    candidate = cleaned
    remote_path = str(item.get("path") or original)
    if candidate.casefold() in used:
        stem, suffix = Path(cleaned).stem, Path(cleaned).suffix
        candidate = f"{stem}-{hashlib.sha1(remote_path.encode('utf-8')).hexdigest()[:8]}{suffix}"
    used.add(candidate.casefold())
    return candidate


def sync_folder_a(public_url, timeout=45):
    """Download Excel files from public Yandex.Disk folder A to persistent local storage."""
    public_url = str(public_url or "").strip()
    if not public_url:
        raise ValueError("Не указана ссылка на папку А")

    api_url = "https://cloud-api.yandex.net/v1/disk/public/resources"
    files = []
    offset = 0
    with requests.Session() as session:
        while offset <= 2000:
            response = session.get(
                api_url,
                params={"public_key": public_url, "limit": 200, "offset": offset},
                timeout=timeout,
            )
            response.raise_for_status()
            payload = response.json()
            items = ((payload.get("_embedded") or {}).get("items") or [])
            files.extend(
                item for item in items
                if item.get("type") == "file" and re.search(r"\.xlsx?$", str(item.get("name") or ""), re.I)
            )
            if len(items) < 200:
                break
            offset += 200

        FOLDER_A_DIR.mkdir(parents=True, exist_ok=True)
        previous = {}
        if FOLDER_A_MANIFEST_FILE.exists():
            try:
                previous = {
                    str(item.get("remote_path")): item
                    for item in json.loads(FOLDER_A_MANIFEST_FILE.read_text(encoding="utf-8")).get("files", [])
                }
            except (OSError, json.JSONDecodeError, AttributeError):
                previous = {}

        used, manifest, downloaded, reused = set(), [], 0, 0
        for item in sorted(files, key=lambda value: str(value.get("path") or value.get("name") or "")):
            local_name = _safe_remote_name(item, used)
            local_path = FOLDER_A_DIR / local_name
            remote_path = str(item.get("path") or item.get("name") or local_name)
            checksum = str(item.get("md5") or "")
            old = previous.get(remote_path) or {}
            unchanged = local_path.exists() and checksum and checksum == str(old.get("md5") or "")
            if unchanged:
                reused += 1
            else:
                href = item.get("file")
                if not href:
                    link_response = session.get(
                        api_url + "/download",
                        params={"public_key": public_url, "path": remote_path},
                        timeout=timeout,
                    )
                    link_response.raise_for_status()
                    href = link_response.json().get("href")
                if not href:
                    raise RuntimeError(f"Нет ссылки для скачивания файла {item.get('name')}")
                file_response = session.get(href, timeout=timeout)
                file_response.raise_for_status()
                tmp_path = local_path.with_suffix(local_path.suffix + ".part")
                tmp_path.write_bytes(file_response.content)
                os.replace(tmp_path, local_path)
                downloaded += 1

            modified = item.get("modified") or item.get("created")
            if modified:
                try:
                    timestamp = datetime.fromisoformat(str(modified).replace("Z", "+00:00")).timestamp()
                    os.utime(local_path, (timestamp, timestamp))
                except (OSError, ValueError):
                    pass
            manifest.append({
                "name": item.get("name"),
                "local_name": local_name,
                "remote_path": remote_path,
                "md5": checksum,
                "modified": modified,
                "size": item.get("size"),
            })

    active_names = {item["local_name"] for item in manifest}
    removed = 0
    for path in FOLDER_A_DIR.iterdir():
        if path.is_file() and re.search(r"\.xlsx?$", path.name, re.I) and path.name not in active_names:
            path.unlink()
            removed += 1
    _atomic_write_json(FOLDER_A_MANIFEST_FILE, {
        "source": public_url,
        "synced_at": datetime.now().isoformat(timespec="seconds"),
        "files": manifest,
    })
    return {"files": len(manifest), "downloaded": downloaded, "reused": reused, "removed": removed}


def sync_and_scan_folder_a(public_url):
    result = sync_folder_a(public_url)
    added, skipped = scan_folder(FOLDER_A_DIR)
    result.update({"added": added, "skipped": skipped, "state": load_state()})
    return result

def approve(idx):
    st = load_state()
    if idx < 0 or idx >= len(st["pending"]): return False
    p = st["pending"].pop(idx)
    g = {}
    for it in p["items"]: g.setdefault(it["nn"], []).append(it)
    merged = []
    for arr in g.values():
        b = arr[0]
        if len(arr) > 1: b["qty"] = sum(x.get("qty") or 0 for x in arr)
        merged.append(b)
    key = norm(p["rso"])
    st["clean"][key] = {"rso": p["rso"], "okrug": p.get("okrug") or resolve_municipality(p["rso"]), "date": p.get("date"), "items": merged}
    st["clean"][key].update({field: p[field] for field in SOURCE_FIELDS if field in p})
    # Ручная повторная загрузка и согласование явно возвращают ранее удалённое РСО.
    st.setdefault("excluded_rso", {}).pop(key, None)
    save_state(st)
    publish_clean(st)
    return True

def reject(idx):
    st = load_state()
    if idx < 0 or idx >= len(st["pending"]): return False
    st["pending"].pop(idx); save_state(st); return True


def delete_rso(rso):
    """Delete one RSO everywhere and keep its Folder A file from restoring it."""
    rso = str(rso or "").strip()
    key = norm(rso)
    if not key:
        raise ValueError("Не указано РСО")

    st = load_state()
    pending_before = len(st["pending"])
    st["pending"] = [item for item in st["pending"] if norm(item.get("rso")) != key]
    pending_removed = pending_before - len(st["pending"])
    clean_item = st["clean"].pop(key, None)
    display_name = str((clean_item or {}).get("rso") or rso).strip()
    st.setdefault("excluded_rso", {})[key] = {
        "rso": display_name,
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

def edit_item(pi, ii, cat, grp):
    st = load_state()
    if pi < 0 or ii < 0:
        return False
    try:
        it = st["pending"][pi]["items"][ii]
    except (IndexError, KeyError):
        return False
    if cat:
        cat, grp = dictionary.remember(norm(it["name"]), cat, grp, it.get("water"))
    elif grp:
        raise ValueError("Выберите категорию для группы")
    it["cat"] = cat or None; it["grp"] = grp or None; it["via"] = "override" if cat else None
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


def publish_rows(rows, st=None):
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
    return publish_rows(export_rows_from_state(state), state)


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
