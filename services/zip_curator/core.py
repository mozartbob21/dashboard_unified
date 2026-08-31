"""Куратор ЗиП: локальные алгоритмы без облачных зависимостей."""
import io, json, re, zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

BASE = Path(__file__).resolve().parents[2]
DATA = BASE / "data" / "zip_curator"
INPUT_DIR = Path(__import__("os").getenv("ZIP_INPUT_DIR", str(DATA / "input")))
STATE_FILE = DATA / "state.json"
DICT_JSON = Path(__file__).with_name("zip_dict.json")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

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

def norm_unit(u):
    if u is None or not str(u).strip(): return None, True
    c = D["umap"].get(str(u).lower().strip())
    return c, not c

def classify(name):
    nn = norm(name)
    r = D["dict"].get(nn)
    if r: return {"cat": r[0], "grp": r[1], "water": bool(r[2]), "via": "match"}
    for kw in D["kws"]:
        if kw[0] in nn: return {"cat": kw[1], "grp": kw[2], "water": False, "via": "keyword"}
    return {"cat": None, "grp": None, "water": False, "via": None}

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

def load_state():
    if STATE_FILE.exists():
        try: return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception: pass
    return {"pending": [], "clean": {}}

def save_state(st):
    DATA.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(st, ensure_ascii=False, indent=1), encoding="utf-8")

def _enrich(p):
    for it in p["items"]:
        c = classify(it["name"]); u, unk = norm_unit(it.get("unitRaw"))
        it.update(nn=norm(it["name"]), cat=c["cat"], grp=c["grp"], water=c["water"], via=c["via"], unit=u, unitUnknown=unk)
    p["okrug"] = D["rso2omsu"].get(p["rso"], "")

def ingest(list_of_parsed):
    st = load_state()
    added = 0
    for p in list_of_parsed:
        if not p["items"]: continue
        _enrich(p)
        st["pending"] = [x for x in st["pending"] if norm(x["rso"]) != norm(p["rso"])]
        st["pending"].append(p); added += 1
    save_state(st)
    return added

def scan_folder():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    files = sorted(INPUT_DIR.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
    taken, skipped = {}, 0
    for f in files:
        try: rows = read_xlsx_rows(f.read_bytes())
        except Exception: continue
        p = parse_reestr(rows, f.name)
        if not p["items"]: continue
        key = norm(p["rso"])
        if key in taken: skipped += 1; continue
        p["fname"] = f.name; p["uploaded"] = int(f.stat().st_mtime)
        taken[key] = p
    return ingest(list(taken.values())), skipped

def approve(idx):
    st = load_state()
    if idx >= len(st["pending"]): return False
    p = st["pending"].pop(idx)
    g = {}
    for it in p["items"]: g.setdefault(it["nn"], []).append(it)
    merged = []
    for arr in g.values():
        b = arr[0]
        if len(arr) > 1: b["qty"] = sum(x.get("qty") or 0 for x in arr)
        merged.append(b)
    st["clean"][norm(p["rso"])] = {"rso": p["rso"], "okrug": p.get("okrug") or D["rso2omsu"].get(p["rso"], ""), "date": p.get("date"), "items": merged}
    save_state(st); return True

def reject(idx):
    st = load_state()
    if idx >= len(st["pending"]): return False
    st["pending"].pop(idx); save_state(st); return True

def edit_item(pi, ii, cat, grp):
    st = load_state()
    try:
        it = st["pending"][pi]["items"][ii]
        it["cat"] = cat or None; it["grp"] = grp or None; it["via"] = "override"
        if cat: D["dict"][it["nn"]] = [cat, grp or "", 1 if it.get("water") else 0]
        save_state(st); return True
    except Exception: return False

def export_rows():
    st = load_state()
    rows = [["РСО","Округ","Наименование","Категория","Группа","Количество","Ед.","Водоподготовка","Дата"]]
    for c in st["clean"].values():
        for it in c["items"]:
            rows.append([c["rso"], c["okrug"], it["name"], it.get("cat") or "не указано", it.get("grp") or "", it.get("qty"), it.get("unit") or "", "да" if it.get("water") else "", c.get("date") or ""])
    return rows

def write_xlsx(rows, path):
    def esc(s): return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    for r, row in enumerate(rows, 1):
        sheet.append(f'<row r="{r}">')
        for c, v in enumerate(row):
            ref = ""
            n = c
            while True: ref = chr(65 + n % 26) + ref; n = n//26 - 1
            if n < 0: pass
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
