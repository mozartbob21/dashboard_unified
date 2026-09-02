# -*- coding: utf-8 -*-
"""Детализированный контекст платформы для ИИ: кто, где и что нарушил."""
import json
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = BASE_DIR / "data"

F_EDO = DATA_DIR / "edo" / "result.json"
F_OVERDUE = DATA_DIR / "overdue" / "final_result.json"
F_WATER = DATA_DIR / "watercontrol" / "result.json"
F_UTNKR = DATA_DIR / "utnkr" / "violators.json"
F_MGKH = DATA_DIR / "mgkh_rm" / "result.json"
F_CAM = DATA_DIR / "cameras" / "state" / "dashboard_state.json"

BAD = {"critical", "red", "критично", "красный", "не работает", "not_working", "offline", "error", "broken"}
RISK = {"risk", "warning", "yellow", "риск", "желтый", "жёлтый", "high", "высокий", "not_connected"}


def _load(path):
    try:
        if not path.exists():
            return None
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def _rows(data):
    if not data:
        return []
    if isinstance(data, list):
        return [r for r in data if isinstance(r, dict)]
    if isinstance(data, dict):
        for key in ("rows", "items", "violators", "cameras", "results", "data", "records"):
            v = data.get(key)
            if isinstance(v, list):
                return [r for r in v if isinstance(r, dict)]
    return []


def _s(row, *keys):
    for k in keys:
        v = row.get(k)
        if v not in (None, ""):
            return str(v).strip()
    return ""


def _num(row, *keys):
    for k in keys:
        v = row.get(k)
        if v in (None, ""):
            continue
        try:
            return int(float(str(v).replace(" ", "").replace(",", ".")))
        except Exception:
            continue
    return 0


def _status(row):
    return _s(row, "status", "traffic_light", "color", "level", "camera_status").lower()


def _detail(row, module):
    muni = _s(row, "municipality", "city", "округ")
    org = _s(row, "organization", "owner", "company")
    obj = _s(row, "object", "name", "address", "subject")
    resp = _s(row, "responsible", "responsible_name")
    overdue = _num(row, "overdue_days", "days_overdue", "days", "delay_days", "overdue_count")
    st = _status(row)
    parts = []
    if muni:
        parts.append(muni)
    if org and org != muni:
        parts.append(org)
    if obj:
        parts.append(obj)
    if overdue:
        parts.append(f"просрочка {overdue} дн.")
    if st:
        parts.append(f"статус: {st}")
    if resp:
        parts.append(f"отв.: {resp}")
    if not parts:
        return ""
    return "  • [" + module + "] " + " | ".join(parts)


def _module_block(title, rows, module, limit=8):
    bad = [r for r in rows if _status(r) in BAD]
    risk = [r for r in rows if _status(r) in RISK]
    ok = len(rows) - len(bad) - len(risk)
    out = [f"{title}: всего {len(rows)}, критичных {len(bad)}, риск {len(risk)}, в норме {ok}."]
    shown = 0
    for r in bad + risk:
        if shown >= limit:
            left = len(bad) + len(risk) - shown
            if left > 0:
                out.append(f"  … и ещё {left} проблемных.")
            break
        line = _detail(r, module)
        if line:
            out.append(line)
            shown += 1
    return out


def _mgkh_rows(data):
    out = []
    if not isinstance(data, dict):
        return out
    b = data.get("buckets")
    if isinstance(b, dict):
        for label, key in (("закрыть", "close"), ("продлить", "extend"), ("переделать", "rework")):
            for r in b.get(key) or []:
                if isinstance(r, dict):
                    r2 = dict(r)
                    r2["_action"] = label
                    out.append(r2)
    return out


def build_detailed_context(question: str = "") -> str:
    lines = ["АКТУАЛЬНЫЕ ДАННЫЕ ПЛАТФОРМЫ НА " + datetime.now().strftime("%d.%m.%Y %H:%M") +
             ". Отвечай СТРОГО по ним; ниже — детализация по проблемным объектам."]

    edo = _rows(_load(F_EDO))
    lines += _module_block("ЭДО", edo, "ЭДО")

    ov = _rows(_load(F_OVERDUE))
    for r in ov:
        c = _num(r, "overdue_count")
        r["status"] = "critical" if c >= 20 else ("risk" if c > 0 else "ok")
    lines += _module_block("Просроченные задачи", ov, "просрочка")

    lines += _module_block("WaterControl", _rows(_load(F_WATER)), "вода")

    ut = _rows(_load(F_UTNKR))
    for r in ut:
        if not _status(r):
            d = _num(r, "overdue_days", "days_overdue", "days", "delay_days")
            r["status"] = "critical" if d >= 20 else ("risk" if d > 0 else "ok")
    lines += _module_block("Технадзор УТНКР", ut, "технадзор")

    lines += _module_block("Камеры", _rows(_load(F_CAM)), "камеры")

    mgkh = _load(F_MGKH)
    mrows = _mgkh_rows(mgkh)
    if mrows:
        lines.append(f"МКХ Redmine: задач {len(mrows)}.")
        for r in mrows[:8]:
            muni = _s(r, "municipality", "city")
            subj = _s(r, "subject", "name", "object")
            line = " | ".join(x for x in [muni, subj, "действие: " + r.get("_action", "")] if x)
            if line:
                lines.append("  • [мгх] " + line)
    else:
        m = mgkh.get("metrics") if isinstance(mgkh, dict) else None
        if isinstance(m, dict):
            lines.append(f"МКХ Redmine: всего {m.get('total', 0)}, закрыть {m.get('close', 0)}, "
                         f"продлить {m.get('extend', 0)}, переделать {m.get('rework', 0)}.")

    return "\n".join(lines)
