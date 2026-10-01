"""Persist only aggregate ARM indicators for later chat reports.

Selection mirrors dashboard.html build(): valid claim ID and creation date,
then the five water/wastewater object flags. Raw claims never reach the snapshot.
"""
import json
import math
import os
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

TARGET = Path(__file__).resolve().parents[2] / 'data/edds/arm-summary.json'
WATER_FIELDS = ('obj_vzu', 'obj_vs', 'obj_kns', 'obj_kos', 'obj_vo')


def _text(value):
    return '' if value is None else str(value).strip()


def _date(value):
    match = re.match(r'^(\d{2})\.(\d{2})\.(\d{4}|\d{2})\s+(\d{1,2}):(\d{2})', _text(value))
    if not match:
        return None
    day, month, year, hour, minute = (int(v) for v in match.groups())
    try:
        return datetime(year + 2000 if len(match[3]) == 2 else year, month, day, hour, minute)
    except ValueError:
        return None


def _count(value):
    try:
        number = float(_text(value).replace(',', '.'))
        return number if math.isfinite(number) and number >= 0 else None
    except ValueError:
        return None


def _empty():
    return {'incidents': 0, 'closed': 0, 'active': 0, 'people_known_sum': 0, 'people_missing': 0}


def project(grid, start, end):
    if not isinstance(grid, list):
        raise ValueError('Invalid ARM grid')
    header = next((i for i, row in enumerate(grid[:40]) if isinstance(row, list) and row and _text(row[0]) == 'id_cds_claim'), None)
    if header is None:
        raise ValueError('ARM machine header missing')
    columns = {}
    for i, field in enumerate(grid[header]):
        columns.setdefault(_text(field), i)
    if not {'name_mr', 'ispolnitel', 'd_create', 'd_doklad', 'type_otkl'} <= columns.keys() or not any(f in columns for f in WATER_FIELDS):
        raise ValueError('ARM required fields missing')
    def get(row, key):
        pos = columns.get(key)
        return row[pos] if pos is not None and pos < len(row) else None
    municipalities, total = {}, _empty()
    broken, source_rows, unknown_municipality = 0, 0, 0
    for row in grid[header + 1:]:
        if not isinstance(row, list) or not row or not _text(row[0]):
            continue
        if not re.fullmatch(r'\d+', _text(row[0])) or not _date(get(row, 'd_create')):
            broken += 1
            continue
        source_rows += 1
        if not any(_text(get(row, field)) == '1' for field in WATER_FIELDS):
            continue
        name = _text(get(row, 'name_mr'))[:240]
        buckets = [total]
        if name:
            buckets.append(municipalities.setdefault(name, _empty()))
        else:
            unknown_municipality += 1
        closed = bool(_date(get(row, 'd_close')))
        people = _count(get(row, 'cnt_people'))
        for bucket in buckets:
            bucket['incidents'] += 1
            bucket['closed' if closed else 'active'] += 1
            if people is None:
                bucket['people_missing'] += 1
            else:
                bucket['people_known_sum'] += people
    # An entirely broken export must not erase the last good data with zeroes.
    if broken and not total['incidents']:
        raise ValueError('ARM rows damaged')
    return {'schema_version': 1, 'collected_at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
            'period': {'from': start.isoformat(), 'to': end.isoformat()},
            'totals': total, 'municipalities': municipalities, 'broken_rows': broken,
            'source_rows': source_rows, 'unknown_municipality': unknown_municipality}


def persist(grid, start, end):
    """Snapshot failure must not break the existing browser report route."""
    temporary = None
    try:
        payload = project(grid, start, end)
        TARGET.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix='.arm-summary-', dir=TARGET.parent)
        with os.fdopen(fd, 'w', encoding='utf-8') as out:
            json.dump(payload, out, ensure_ascii=False, allow_nan=False)
        os.replace(temporary, TARGET)
        return True
    except (OSError, ValueError, TypeError):
        return False
    finally:
        if temporary:
            try:
                Path(temporary).unlink(missing_ok=True)
            except OSError:
                pass
