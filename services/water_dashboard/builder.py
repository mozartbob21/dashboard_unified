import json
import re
from datetime import datetime

from services.water_dashboard.config import SNAPSHOT_FILE


def to_int(v):
    try:
        s = str(v).strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
        return int(float(s.rstrip("%"))) if s and s not in ("—", "-") else None
    except Exception:
        return None


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
            "name": name, "resVS": None, "sysVS": None, "tasks": None,
            "sysKR": None, "resKR": None, "att": None,
        })

    for src_id, data in (extractions or {}).items():
        for t in data.get("tables", []):
            headers = [str(h).strip().lower() for h in t.get("headers", [])]
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
                "tasks": col("просроч", "кол-во", "задач"),
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
                if not cells: continue
                name = norm_name(cells[0])
                if not name or name.lower().startswith("итого"):
                    continue
                r = row(name)
                for field, i in idx.items():
                    if i is not None and i < len(cells):
                        r[field] = to_int(cells[i])

    return sorted(merged.values(), key=lambda r: -(r["resVS"] or 0))


def derive_kpis(table):
    s = lambda k: sum(r[k] for r in table if r[k] is not None) if any(r[k] is not None for r in table) else None
    att_rows = [r for r in table if r["att"] is not None]
    return {
        "tasks_total": s("tasks"),
        "sys_vs": s("sysVS"),
        "res_vs": s("resVS"),
        "sys_kr": s("sysKR"),
        "res_kr": s("resKR"),
        "att_avg": round(sum(r["att"] for r in att_rows) / len(att_rows)) if att_rows else None,
    }


def top5(table, field):
    rows = sorted((r for r in table if r[field] is not None), key=lambda r: -r[field])[:5]
    return [{"name": r["name"], "value": r[field]} for r in rows if r[field] > 0]


def bottom5(table, field):
    """Топ-5 лучших (минимальные значения = лучшие показатели)."""
    rows = sorted((r for r in table if r[field] is not None), key=lambda r: r[field])[:5]
    return [{"name": r["name"], "value": r[field]} for r in rows if r[field] >= 0]



def extract_refresh_info(text):
    """Достаёт строку вида «Дашборд автоматически обновляется каждые 30 мин.»
    Хвосты после времени/минут обрезаем (лишние пояснения про кликабельность)."""
    if not text:
        return ""
    m = re.search(r"[^\n]*каждые\s+[\d\s–-]+мин", text, re.IGNORECASE)
    if m:
        return m.group(0).strip().rstrip(".,;:")
    # «…обновляется ежедневно с 9:00 до 10:30» — стоп после времени
    m = re.search(r"[^\n]*обновляетс[^\n]*?\d{1,2}:\d{2}(?:\s*до\s*\d{1,2}:\d{2})?",
                  text, re.IGNORECASE)
    if m:
        return m.group(0).strip().rstrip(".,;:")
    m = re.search(r"[^\n]*обновляетс[^\n]*", text, re.IGNORECASE)
    return m.group(0).strip().rstrip(".,;:") if m else ""


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


def extract_widgets(text):
    """Read labelled numeric cards, never invent values from an old example."""
    lines = [line.strip() for line in (text or '').splitlines() if line.strip()]
    items = []
    for i, label in enumerate(lines[:-1]):
        if not re.search(r'[А-Яа-яA-Za-z]', label) or len(label) > 160:
            continue
        if re.match(r'^(Еще|Ещё|Показать|Скрыть|Дата|Период)\b', label):
            continue
        j = i + 1
        while j < len(lines) and re.match(r'^(Еще|Ещё)\s+\d+$', lines[j]):
            j += 1
        if j < len(lines) and re.fullmatch(r'[-+−]?\d[\d\s.,%/₽КМKMBкм]*', lines[j]):
            item = {'label': label, 'value': lines[j]}
            if item not in items:
                items.append(item)
    return items[:40]


def build_snapshot(extractions):
    from services.water_dashboard.config import SOURCES
    import os
    import tempfile
    prev = {}
    if SNAPSHOT_FILE.exists():
        try:
            prev = json.loads(SNAPSHOT_FILE.read_text(encoding='utf-8'))
        except (ValueError, OSError):
            pass
    now = datetime.now().isoformat(timespec='seconds')
    sources = {}
    for source in SOURCES:
        sid = source['id']
        data = (extractions or {}).get(sid, {})
        tables = data.get('tables') or []
        widgets = extract_widgets(data.get('text', ''))
        valid = not data.get('error') and bool(tables or widgets)
        previous = (prev.get('sources') or {}).get(sid, {})
        sources[sid] = dict(previous) if not valid else {
            'tables': tables, 'widgets': widgets,
            'text': data.get('text', ''), 'updated_at': now,
            'refresh': extract_refresh_info(data.get('text', '')),
        }
        sources[sid].update(source)
        sources[sid].update(checked_at=now, ok=valid,
            error='' if valid else data.get('error') or 'Источник не вернул распознаваемые показатели. Повторите обновление.')
    # Each source retains its own last successful data/time on a partial failure.
    table = build_table(sources)
    snap = {
        'schema_version': 2, 'checked_at': now,
        'updated_at': max((d.get('updated_at', '') for d in sources.values()), default=''),
        'snapshot_date': datetime.now().strftime('%d.%m.%Y') if any(d['ok'] for d in sources.values()) else prev.get('snapshot_date', '—'),
        'sources': sources, 'table': table, 'kpis': derive_kpis(table),
        'sources_refresh': {sid:d.get('refresh','') for sid,d in sources.items()},
        'sources_updated': {sid:d['ok'] for sid,d in sources.items()},
        'kpi_live': {'nvos': parse_nvos_kpis(sources['nvos'].get('text',''))},
        'tops': {k: top5(table, k) for k in ('resVS','sysVS','tasks','sysKR','resKR')},
        'bottoms': {k: bottom5(table, k) for k in ('resVS','sysVS','tasks','sysKR','resKR')},
    }
    SNAPSHOT_FILE.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=SNAPSHOT_FILE.parent, suffix='.tmp')
    try:
        with os.fdopen(fd,'w',encoding='utf-8') as out:
            json.dump(snap,out,ensure_ascii=False)
        os.replace(tmp,SNAPSHOT_FILE)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)
    return snap
