# -*- coding: utf-8 -*-
"""Контекст сводного дашборда по качеству водоснабжения для ИИ-чата."""
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
WD_DIR = BASE_DIR / "data" / "water_dashboard"
SNAP = WD_DIR / "snapshot.json"
BEST5 = WD_DIR / "best5.json"


def _load(p):
    try:
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        pass
    return None


def _name_of(e):
    if isinstance(e, dict):
        for k in ("municipality", "name", "org", "rso", "object", "label"):
            v = e.get(k)
            if v:
                return str(v)
        v = next((x for x in e.values() if isinstance(x, str) and x.strip()), None)
        return str(v) if v else str(e)
    return str(e)


def build_water_context(limit_rows: int = 15) -> str:
    snap = _load(SNAP)
    best5 = _load(BEST5)
    if not snap and not best5:
        return ""
    lines = ["🌊 СВОДНЫЙ ДАШБОРД ПО КАЧЕСТВУ ВОДОСНАБЖЕНИЯ (данные платформы, отвечай строго по ним):"]

    if snap:
        lines.append(f"Снимок от {snap.get('snapshot_date', '—')}, обновлено {snap.get('updated_at', '—')}.")

        kpi = snap.get("kpi_live") or {}
        if isinstance(kpi, dict):
            for k, v in list(kpi.items())[:12]:
                lines.append(f"  KPI: {k} = {v}")

        srcs = snap.get("sources_refresh") or {}
        if isinstance(srcs, dict):
            for k, v in list(srcs.items())[:12]:
                lines.append(f"  Источник «{k}»: {v}")

        bottoms = snap.get("bottoms") or {}
        if isinstance(bottoms, dict):
            for k, v in bottoms.items():
                if isinstance(v, list):
                    lines.append(f"  ХУДШИЕ по «{k}»: {', '.join(_name_of(x) for x in v[:5])}")
                else:
                    lines.append(f"  ХУДШИЕ по «{k}»: {v}")

        table = snap.get("table") or []
        if isinstance(table, list) and table and isinstance(table[0], dict):
            hdr = list(table[0].keys())
            lines.append(f"  Таблица показателей: строк {len(table)}; колонки: {', '.join(hdr[:8])}.")
            for row in table[:limit_rows]:
                name = _name_of(row)
                nums = [f"{k}={row.get(k)}" for k in hdr[1:7]
                        if row.get(k) not in (None, "", "-")]
                if nums:
                    lines.append("  • " + name + " | " + " | ".join(nums))

    if best5:
        if isinstance(best5, dict):
            for key, val in best5.items():
                if isinstance(val, list):
                    lines.append(f"  ЛУЧШИЕ по «{key}»: {', '.join(_name_of(x) for x in val[:5])}")
                else:
                    lines.append(f"  ЛУЧШИЕ по «{key}»: {val}")
        elif isinstance(best5, list):
            lines.append("  ЛУЧШИЕ: " + ", ".join(_name_of(x) for x in best5[:5]))

    return "\n".join(lines)
