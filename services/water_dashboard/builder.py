import json
import re
from datetime import datetime

from services.water_dashboard.config import SNAPSHOT_FILE


def to_int(v):
    try:
        s = str(v).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
        return int(float(s)) if s else 0
    except Exception:
        return 0


def norm_name(s):
    s = (s or "").strip()
    if not s or not s.isupper():
        return s
    fmt = ["-".join(p.capitalize() for p in w.split("-")) for w in s.split()]
    if len(fmt) > 1:
        return fmt[0] + " " + " ".join(w.lower() for w in fmt[1:])
    return fmt[0]


def build_table(extractions):
    merged = {}

    def row(name):
        return merged.setdefault(name, {
            "name": name, "resVS": 0, "sysVS": 0, "tasks": 0,
            "sysKR": 0, "resKR": 0, "att": 0,
        })

    for src_id, data in (extractions or {}).items():
        for t in data.get("tables", []):
            headers = t.get("headers", [])
            if not headers or not ("омсу" in headers[0] or "муниципал" in headers[0]):
                continue

            def col(*keys):
                for i, h in enumerate(headers):
                    if any(k in h for k in keys):
                        return i
                return None

            idx = {
                "resVS": col("резонанс"),
                "sysVS": col("систем"),
                "tasks": col("кол-во", "задач"),
                "att": col("явка"),
            }
            if src_id == "sys_kr":
                idx["resKR"], idx["sysKR"] = idx.pop("resVS"), idx.pop("sysVS")
            if src_id != "sys_vs":
                idx.pop("resVS", None); idx.pop("sysVS", None)
            if src_id != "tasks":
                idx.pop("tasks", None)
            if src_id != "meetings":
                idx.pop("att", None)

            for cells in t.get("rows", []):
                name = norm_name(cells[0])
                if not name or name.lower().startswith("итого"):
                    continue
                r = row(name)
                for field, i in idx.items():
                    if i is not None and i < len(cells):
                        r[field] = to_int(cells[i])

    return sorted(merged.values(), key=lambda r: -r["resVS"])


def derive_kpis(table):
    s = lambda k: sum(r[k] for r in table)
    att_rows = [r for r in table if r["att"] > 0]
    return {
        "tasks_total": s("tasks"),
        "sys_vs": s("sysVS"),
        "res_vs": s("resVS"),
        "sys_kr": s("sysKR"),
        "res_kr": s("resKR"),
        "att_avg": round(sum(r["att"] for r in att_rows) / len(att_rows)) if att_rows else 0,
    }


def top5(table, field):
    rows = sorted(table, key=lambda r: -r[field])[:5]
    return [{"name": r["name"], "value": r[field]} for r in rows if r[field] > 0]


def extract_refresh_info(text):
    """Достаёт строку вида «Дашборд автоматически обновляется каждые 30 мин.»"""
    if not text:
        return ""
    m = re.search(r"[^\n]*каждые\s+[\d\s–-]+мин[^\n]*", text, re.IGNORECASE)
    if m:
        return m.group(0).strip()
    m = re.search(r"[^\n]*обновляетс[^\n]*", text, re.IGNORECASE)
    return m.group(0).strip() if m else ""


def parse_nvos_kpis(text):
    """Вытаскивает живые показатели НВОС из innerText дашборда.
    Формат виджета: «Заголовок → Еще 0 → Значение → (подпись)»."""
    if not text:
        return {}
    out = {}

    # Собираемость: «Доля (%) / Еще 0 / 80,00 / собираемости...»
    m = re.search(
        r"Доля\s*\(%\)[^\n]*\s*(?:\n\s*Еще\s+\d+)?\s*\n\s*([\d.,\s]+?)\s*собираемости",
        text,
    )
    if m:
        out["sbor"] = m.group(1).strip()

    # % от НВВ: «% от НВВ / Еще 0 / 5,76»
    m = re.search(r"% от НВВ\s*(?:\n\s*Еще\s+\d+)?\s*\n\s*([\d.,]+)", text)
    if m:
        out["nvv_pct"] = m.group(1).strip()

    # Сумма НВВ: «Сумма НВВ / Еще 0 / 24 491M»
    m = re.search(r"Сумма НВВ\s*(?:\n\s*Еще\s+\d+)?\s*\n\s*([\d\s]+[MКМ])", text)
    if m:
        out["sum_nvv"] = m.group(1).strip()

    # План отборов проб (год)
    m = re.search(r"План отборов проб\s*(?:\n\s*Еще\s+\d+)?\s*\n\s*([\d\s]+?)\s*\(год\)", text)
    if m:
        out["plan_year"] = m.group(1).strip()

    # Факт отборов проб (год)
    m = re.search(r"Факт отборов проб\s*(?:\n\s*Еще\s+\d+)?\s*\n\s*([\d\s]+?)\s*\(год\)", text)
    if m:
        out["fact_year"] = m.group(1).strip()

    # План отборов проб (неделя)
    m = re.search(r"План отборов проб\s*(?:\n\s*Еще\s+\d+)?\s*\n\s*([\d\s]+?)\s*\(неделя\)", text)
    if m:
        out["plan_week"] = m.group(1).strip()

    # Факт отборов проб (неделя)
    m = re.search(r"Факт отборов проб\s*(?:\n\s*Еще\s+\d+)?\s*\n\s*([\d\s]+?)\s*\(неделя\)", text)
    if m:
        out["fact_week"] = m.group(1).strip()

    # Получено / начислено из графика «Плата за негативное воздействие, руб»
    # В тексте: «1 644 284,84K» (начислено) и «1 249 134,34K» (получено)
    m = re.search(
        r"Плата за негативное воздействие, руб(.*?)Организация",
        text, re.DOTALL,
    )
    if m:
        vals = re.findall(r"([\d][\d\s.,]*)K", m.group(1))
        if len(vals) >= 2:
            def to_mln(s):
                v = float(s.replace(" ", "").replace("\u00a0", "").replace(",", ".")) / 1000.0
                return "{:,.1f}".format(v).replace(",", " ").replace(".", ",")
            out["pay_str"] = f"{to_mln(vals[1])} / {to_mln(vals[0])} млн ₽"
                # неразрывные пробелы DataLens → обычные
    for k, v in list(out.items()):
        if isinstance(v, str):
            out[k] = v.replace("\u00a0", " ").strip()



    return out


def _merge_live(prev_live, new_live):
    """Новые значения побеждают; пустой парсинг НЕ затирает прошлые хорошие."""
    merged = dict(prev_live or {})
    for k, v in (new_live or {}).items():
        if v:
            merged[k] = v
    return merged


def build_snapshot(extractions):
    prev = {}
    if SNAPSHOT_FILE.exists():
        try:
            prev = json.loads(SNAPSHOT_FILE.read_text(encoding="utf-8"))
        except Exception:
            prev = {}

    table = build_table(extractions)
    kpis = derive_kpis(table)

    sources_refresh = {
        sid: extract_refresh_info((data or {}).get("text", ""))
        for sid, data in (extractions or {}).items()
    }

    snap = {
        "snapshot_date": datetime.now().strftime("%d.%m.%Y"),
        "sources_refresh": sources_refresh,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "table": table,
        "kpis": kpis,
        "tops": {
            "resVS": top5(table, "resVS"),
            "sysVS": top5(table, "sysVS"),
            "tasks": top5(table, "tasks"),
            "sysKR": top5(table, "sysKR"),
            "resKR": top5(table, "resKR"),
            "att_low": sorted([r for r in table if r["att"] > 0], key=lambda r: r["att"])[:5],
        },
        "sources_updated": {sid: bool(d.get("tables")) for sid, d in (extractions or {}).items()},
        "kpi_cards": prev.get("kpi_cards", {}),
        "kpi_live": {
            "nvos": _merge_live(
                (prev.get("kpi_live") or {}).get("nvos") or {},
                parse_nvos_kpis((extractions or {}).get("nvos", {}).get("text", "")),
            ),
        },
    }

    SNAPSHOT_FILE.parent.mkdir(parents=True, exist_ok=True)
    SNAPSHOT_FILE.write_text(json.dumps(snap, ensure_ascii=False, indent=2), encoding="utf-8")
    return snap