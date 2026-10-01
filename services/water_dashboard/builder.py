import json
import re
import math
from copy import deepcopy
from datetime import datetime

from services.water_dashboard.config import SNAPSHOT_FILE


def to_int(v):
    if '%' in str(v):
        return None
    number = to_number(v)
    return int(number) if number is not None and number.is_integer() else None


def to_number(v):
    try:
        s = re.sub(r'\s+', '', str(v)).replace(",", ".")
        value = float(s.rstrip("%")) if s and s not in ("—", "-") else None
        return value if value is not None and math.isfinite(value) else None
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
    seen_values = {}

    def row(name):
        return merged.setdefault(name.casefold().replace("ё", "е"), {
            "name": name, "resVS": None, "sysVS": None, "tasks": None,
            "sysKR": None, "resKR": None, "att": None,
        })

    for src_id, data in (extractions or {}).items():
        for t in data.get("tables", []):
            headers = [str(h).strip().lower() for h in t.get("headers", [])]
            def exact_col(*keys):
                for key in keys:
                    for i, header in enumerate(headers):
                        if ' '.join(header.split()) == key:
                            return i
                return None

            name_index = exact_col("омсу", "муниципалитет", "муниципальный округ", "городской округ")
            if name_index is None:
                continue

            idx = {
                "resVS": exact_col("резонансных", "резонансные адреса", "количество резонансных адресов", "кол-во резонансных адресов"),
                "sysVS": exact_col("системных", "системные адреса", "количество системных адресов", "кол-во системных адресов"),
                "tasks": exact_col("кол-во задач", "количество задач", "просроченные задачи", "количество просроченных задач", "кол-во просроченных задач", "количество"),
                "att": exact_col("присутствовали на перекличках, %", "явка, %", "явка (%)", "явка"),
            }
            if src_id == "sys_kr":
                idx["resKR"], idx["sysKR"] = idx.pop("resVS"), idx.pop("sysVS")
            if src_id != "sys_vs":
                idx.pop("resVS", None); idx.pop("sysVS", None)
            if src_id != "tasks":
                idx.pop("tasks", None)
            if src_id != "meetings":
                idx.pop("att", None)

            if not any(i is not None for i in idx.values()):
                continue
            for cells in t.get("rows", []):
                if len(cells) <= name_index: continue
                name = norm_name(cells[name_index])
                if not name or name.lower().startswith("итого"):
                    continue
                r = row(name)
                for field, i in idx.items():
                    if i is not None and i < len(cells):
                        value = to_number(cells[i]) if field == 'att' else to_int(cells[i])
                        if value is not None and (value < 0 or (field == 'att' and value > 100)):
                            value = None
                        key = (name.casefold().replace("ё", "е"), field)
                        observed = seen_values.setdefault(key, set())
                        observed.add(value)
                        # Duplicate tables repeat the same value. Conflicting
                        # municipality records are unknown, never last-wins.
                        r[field] = value if len(observed) == 1 else None

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
        "att_avg": round(sum(r["att"] for r in att_rows) / len(att_rows), 1) if att_rows else None,
    }


def verified_kpis(sources):
    mapping = {
        'tasks_total': ('tasks', 'overdue_tasks'),
        'sys_vs': ('sys_vs', 'system_addresses'),
        'res_vs': ('sys_vs', 'resonant_addresses'),
        'sys_kr': ('sys_kr', 'system_addresses'),
        'res_kr': ('sys_kr', 'resonant_addresses'),
        'att_avg': ('meetings', 'attendance_pct'),
    }
    return {key: next((item['value'] for item in sources.get(sid, {}).get('metrics', [])
                       if item['id'] == metric_id), None)
            for key, (sid, metric_id) in mapping.items()}


def top5(table, field):
    rows = sorted((r for r in table if r[field] is not None), key=lambda r: -r[field])[:5]
    return [{"name": r["name"], "value": r[field]} for r in rows if r[field] > 0]


def bottom5(table, field):
    """Топ-5 лучших (минимальные значения = лучшие показатели)."""
    rows = sorted((r for r in table if r[field] is not None), key=lambda r: r[field])[:5]
    return [{"name": r["name"], "value": r[field]} for r in rows if r[field] >= 0]



def extract_refresh_info(text):
    """Read the schedule published in the source, without inventing an interval."""
    if not isinstance(text, str):
        return ""
    for line in text.splitlines():
        line = ' '.join(line.split()).strip()
        if not re.search(r'обновля(?:ется|ются)|актуализиру(?:ется|ются)', line, re.I):
            continue
        # Exclude adjacent how-to text; preserve weekdays, times and event-based
        # schedules verbatim, including «после переклички на совещании».
        line = re.split(r',\s*(?:названия|название|нажмите|для открытия)\b', line, maxsplit=1, flags=re.I)[0]
        return line[:350].rstrip('.,;:')
    return ""


def parse_nvos_kpis(text):
    """Вытаскивает живые показатели НВОС из innerText дашборда.
    Формат виджета: «Заголовок → Еще 0 → Значение → (подпись)»."""
    if not text:
        return {}
    # DataLens uses both spellings in the widget menu label.
    text = re.sub(r"\bЕщё\b", "Еще", text)
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
    # Error/challenge pages may have numeric status or verification codes. Those
    # are not dashboard values and must not advance the source's success date.
    error_screen = re.compile(
        r'^(?:ошибка\b|error\b|код ошибки\b|страница не найдена\b|'
        r'нет доступа\b|доступ запрещ[её]н\b|access denied\b|forbidden\b|'
        r'войдите в аккаунт\b|авторизуйтесь\b|captcha\b|капча\b|'
        r'я не робот\b|подтвердите.{0,60}(?:не робот|человек)|'
        r'проверка.{0,30}(?:робот|безопасност))', re.IGNORECASE,
    )
    if any(error_screen.search(line) for line in lines):
        return []
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


def normalize_snapshot(snapshot):
    """Project any saved snapshot onto the current seven-source contract.

    Used on reads as well as writes: removing a source takes effect immediately,
    even before the next successful collection. Does not mutate the saved object
    or advance any source's successful observation date.
    """
    from services.water_dashboard.config import SOURCES
    from services.water_dashboard.details import source_details
    from services.water_dashboard.metrics import source_metrics

    original = snapshot if isinstance(snapshot, dict) else {}
    saved_sources = original.get('sources') or {}
    sources = {}
    for meta in SOURCES:
        sid = meta['id']
        data = deepcopy(saved_sources.get(sid) or {})
        if data.get('metric_schema') != 1:
            # Unverified legacy examples and adjacent-page-text numbers are not
            # eligible for either summaries or rankings.
            data = {'metric_schema': 1, 'metrics': source_metrics(sid, {}, []), 'ok': False}
        data.update(meta)
        if not data.get('refresh_checked_at'):
            saved_frequency = extract_refresh_info(data.get('text', ''))
            if saved_frequency:
                data['refresh'] = saved_frequency
                data['refresh_checked_at'] = data.get('updated_at', '')
                data['refresh_available'] = True
        data['details'] = source_details(sid, data)
        sources[sid] = data
    table = build_table(sources)
    result = {
        'schema_version': 3, 'metric_schema': 1, 'details_schema': 1,
        'checked_at': original.get('checked_at', ''),
        'last_checked_sources': [sid for sid in original.get('last_checked_sources', []) if sid in sources],
        'last_updated_sources': [sid for sid in original.get('last_updated_sources', []) if sid in sources],
        'updated_at': max((item.get('updated_at', '') for item in sources.values()), default=''),
        'snapshot_date': original.get('snapshot_date', '—'),
        'sources': sources, 'table': table, 'kpis': verified_kpis(sources),
        'sources_refresh': {sid: item.get('refresh', '') for sid, item in sources.items()},
        'sources_updated': {sid: bool(item.get('ok')) for sid, item in sources.items()},
        'kpi_live': {'nvos': parse_nvos_kpis(sources['nvos'].get('text', ''))},
        'tops': {key: top5(table, key) for key in ('resVS', 'sysVS', 'tasks', 'sysKR', 'resKR')},
        'bottoms': {key: bottom5(table, key) for key in ('resVS', 'sysVS', 'tasks', 'sysKR', 'resKR')},
    }
    return result


def build_snapshot(extractions, source_ids=None):
    from services.water_dashboard.config import SOURCES
    from services.water_dashboard.metrics import source_metrics, primary_available
    import os
    import tempfile
    prev = {}
    if SNAPSHOT_FILE.exists():
        try:
            prev = json.loads(SNAPSHOT_FILE.read_text(encoding='utf-8'))
        except (ValueError, OSError):
            pass
    known_ids = {source['id'] for source in SOURCES}
    selected = known_ids if source_ids is None else set(source_ids)
    if not selected or not selected <= known_ids:
        raise ValueError('Unknown or empty water dashboard source selection')
    now = datetime.now().isoformat(timespec='seconds')
    sources = {}
    for source in SOURCES:
        sid = source['id']
        previous = (prev.get('sources') or {}).get(sid, {})
        if sid not in selected:
            sources[sid] = deepcopy(previous)
            sources[sid].update(source)
            continue
        data = (extractions or {}).get(sid, {})
        tables = data.get('tables') or []
        widgets = data.get('widgets') or []
        profile = dict(data, widgets=widgets)
        if sid == 'nvos':
            profile['nvos'] = parse_nvos_kpis(data.get('text', ''))
        metrics = source_metrics(sid, profile, build_table({sid: profile}))
        valid = not data.get('error') and primary_available(metrics)
        previous = (prev.get('sources') or {}).get(sid, {})
        if previous.get('metric_schema') != 1:
            previous = {}
        sources[sid] = dict(previous) if not valid else {
            'tables': tables, 'widgets': widgets, 'metrics': metrics,
            'metric_schema': 1,
            'text': data.get('text', ''), 'updated_at': now, 'data_date': data.get('data_date'),
            'refresh': extract_refresh_info(data.get('text', '')),
        }
        if valid and sid == 'edo_rso':
            sources[sid].update(ranking_tables=data.get('ranking_tables') or [],
                                ranking_url=data.get('ranking_url', ''),
                                ranking_error=data.get('ranking_error', ''))
        sources[sid].update(source)
        if sources[sid].get('metric_schema') != 1:
            # Old snapshots used adjacent body text as indicators. Do not
            # present those arbitrary values as verified summary metrics.
            sources[sid]['metrics'] = source_metrics(sid, {}, [])
            sources[sid]['metric_schema'] = 1
        frequency_text = data.get('refresh_text', data.get('text', ''))
        frequency = extract_refresh_info(frequency_text)
        if isinstance(frequency_text, str) and frequency_text.strip() and (frequency or not data.get('error')):
            sources[sid]['refresh'] = frequency
            sources[sid]['refresh_checked_at'] = now
            sources[sid]['refresh_available'] = bool(sources[sid]['refresh'])
        sources[sid].update(checked_at=now, ok=valid,
            error='' if valid else data.get('error') or 'В источнике не найден нужный итоговый показатель. Прежние подтверждённые данные сохранены, если были получены ранее.')
    # Each source retains its own last successful data/time on a partial failure.
    table = build_table(sources)
    snap = {
        'schema_version': 2, 'metric_schema': 1, 'checked_at': now,
        'updated_at': max((d.get('updated_at', '') for d in sources.values()), default=''),
        'snapshot_date': datetime.now().strftime('%d.%m.%Y') if any(sources[sid].get('ok') for sid in selected) else prev.get('snapshot_date', '—'),
        'last_checked_sources': [source['id'] for source in SOURCES if source['id'] in selected],
        'last_updated_sources': [source['id'] for source in SOURCES if source['id'] in selected and sources[source['id']].get('ok')],
        'sources': sources, 'table': table, 'kpis': verified_kpis(sources),
        'sources_refresh': {sid:d.get('refresh','') for sid,d in sources.items()},
        'sources_updated': {sid:bool(d.get('ok')) for sid,d in sources.items()},
        'kpi_live': {'nvos': parse_nvos_kpis(sources['nvos'].get('text',''))},
        'tops': {k: top5(table, k) for k in ('resVS','sysVS','tasks','sysKR','resKR')},
        'bottoms': {k: bottom5(table, k) for k in ('resVS','sysVS','tasks','sysKR','resKR')},
    }
    snap = normalize_snapshot(snap)
    SNAPSHOT_FILE.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=SNAPSHOT_FILE.parent, suffix='.tmp')
    try:
        with os.fdopen(fd,'w',encoding='utf-8') as out:
            json.dump(snap,out,ensure_ascii=False)
        os.replace(tmp,SNAPSHOT_FILE)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)
    return snap
