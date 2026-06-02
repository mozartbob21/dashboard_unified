"""
Автогенерация municipality_registry.json из исходных данных
с защитой от мусора через whitelist эталонного списка.
"""
import json
import re
import sys
from pathlib import Path
from typing import Optional

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

REFERENCE_FILE = BASE_DIR / "data" / "mo_municipalities_reference.json"
REGISTRY_FILE = BASE_DIR / "data" / "municipality_registry.json"
SCAN_DIRS = [BASE_DIR / "data", BASE_DIR / "generated"]
ALLOWED_EXT = {".json", ".xlsx", ".xls", ".csv"}

MUNICIPALITY_FIELD_HINTS = (
    "муницип", "округ", "город", "район", "territory", "municipal", "area"
)

ref_data = json.loads(REFERENCE_FILE.read_text(encoding="utf-8"))
REFERENCE_LIST = ref_data["reference"]

def normalize(s):
    return s.lower().replace("ё", "е").strip()

REFERENCE_NORM = {normalize(name): name for name in REFERENCE_LIST}

PREFIX_PATTERNS = [
    r"^\s*городской\s+округ\s+",
    r"^\s*г\.?\s*о\.?\s+",
    r"^\s*муниципальн\w+\s+округ\s+",
    r"^\s*муниципальн\w+\s+район\s+",
    r"^\s*м\.?\s*о\.?\s+",
    r"^\s*мо\s+",
    r"^\s*город\s+",
    r"^\s*г\.\s*",
    r"^\s*район\s+",
    r"\s+район\s*$",
    r"\s+городской\s+округ\s*$",
    r"\s+г\.?\s*о\.?\s*$",
    r"\s+муниципальн\w+\s+округ\s*$",
]

def strip_prefixes(s):
    prev = None
    cur = s.strip()
    while prev != cur:
        prev = cur
        for pat in PREFIX_PATTERNS:
            cur = re.sub(pat, "", cur, flags=re.IGNORECASE).strip()
    return cur

def clean_candidate(raw):
    if not isinstance(raw, str):
        return None
    s = raw.strip()
    if not s:
        return None
    s = re.split(r"[,;:|/\\\n\r\t]", s)[0].strip()
    if not s:
        return None
    s = re.sub(r"\([^)]*\)", "", s).strip()
    if not re.fullmatch(r"[А-Яа-яЁё\-\s]+", s):
        return None
    if len(s) < 3 or len(s) > 40:
        return None
    if len(s.split()) > 4:
        return None
    s = strip_prefixes(s)
    if len(s) < 3:
        return None
    return s

def extract_values_from_obj(obj, depth=0):
    if depth > 6:
        return
    if isinstance(obj, dict):
        for k, v in obj.items():
            key_str = str(k).lower()
            if any(hint in key_str for hint in MUNICIPALITY_FIELD_HINTS):
                if isinstance(v, str):
                    yield v
                elif isinstance(v, list):
                    for item in v:
                        if isinstance(item, str):
                            yield item
            yield from extract_values_from_obj(v, depth + 1)
    elif isinstance(obj, list):
        for item in obj:
            yield from extract_values_from_obj(item, depth + 1)

def read_file_records(path):
    ext = path.suffix.lower()
    try:
        if ext == ".json":
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, list):
                yield from data
            elif isinstance(data, dict):
                yield data
        elif ext in {".xlsx", ".xls"}:
            try:
                import openpyxl
                wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                for ws in wb.worksheets:
                    headers = None
                    for row in ws.iter_rows(values_only=True):
                        if headers is None:
                            headers = [str(c) if c is not None else "" for c in row]
                            continue
                        rec = {h: v for h, v in zip(headers, row)}
                        yield rec
            except ImportError:
                pass
        elif ext == ".csv":
            import csv
            with path.open(encoding="utf-8") as f:
                reader = csv.DictReader(f)
                yield from reader
    except Exception as e:
        print(f"  ! Пропуск {path.name}: {e}", file=sys.stderr)

found_norm = {}

print("Сканирование исходных данных...")
files_scanned = 0
for scan_dir in SCAN_DIRS:
    if not scan_dir.exists():
        continue
    for path in scan_dir.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() not in ALLOWED_EXT:
            continue
        if path.name.startswith("~$") or path.name.startswith("."):
            continue
        low = path.name.lower()
        if "registry" in low or "reference" in low:
            continue
        files_scanned += 1
        for rec in read_file_records(path):
            for raw_val in extract_values_from_obj(rec):
                cand = clean_candidate(raw_val)
                if not cand:
                    continue
                norm = normalize(cand)
                if norm in REFERENCE_NORM:
                    found_norm[norm] = REFERENCE_NORM[norm]

print(f"Файлов просканировано: {files_scanned}")
print(f"Найдено муниципалитетов: {len(found_norm)}")

if REGISTRY_FILE.exists():
    registry = json.loads(REGISTRY_FILE.read_text(encoding="utf-8"))
else:
    registry = {}

municipalities = registry.get("municipalities", {})

for norm, display in found_norm.items():
    if display not in municipalities:
        municipalities[display] = {}

registry["municipalities"] = dict(sorted(municipalities.items(), key=lambda x: x[0].lower()))

if REGISTRY_FILE.exists():
    backup = REGISTRY_FILE.with_suffix(".json.bak")
    backup.write_text(REGISTRY_FILE.read_text(encoding="utf-8"), encoding="utf-8")
    print(f"Бэкап старого реестра: {backup}")

REGISTRY_FILE.write_text(
    json.dumps(registry, ensure_ascii=False, indent=2),
    encoding="utf-8"
)
print(f"Реестр обновлён: {REGISTRY_FILE}")
print("Финальный список:")
for name in registry["municipalities"]:
    print(f"  • {name}")
