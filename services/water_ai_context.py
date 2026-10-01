"""Read-only, typed reporting view of the latest water dashboard snapshot.

No browser profiles, raw portal text, credentials or debug files are read here.
The UI and AI consume the same saved snapshot; every source keeps its own date.
"""
import json
import math
import re
from pathlib import Path
from services.report_municipalities import key as municipality_key, display as municipality_display

BASE_DIR = Path(__file__).resolve().parents[1]
SNAP = BASE_DIR / 'data/water_dashboard/snapshot.json'
MAX_SNAPSHOT_BYTES = 20 * 1024 * 1024


def load_snapshot(path=None):
    path = Path(path or SNAP)
    try:
        if path.stat().st_size > MAX_SNAPSHOT_BYTES:
            return {}
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        if not isinstance(data, dict) or data.get('schema_version') not in (2, 3):
            return {}
        from services.water_dashboard.builder import normalize_snapshot
        return normalize_snapshot(data)
    except (OSError, ValueError, TypeError):
        return {}


def normalized(value):
    return re.sub(r'\s+', ' ', str(value or '').casefold().replace('ё', 'е')).strip()


def text(value, limit=240):
    return str(value or '').strip()[:limit] if isinstance(value, (str, int, float)) else ''


def finite_number(value):
    return value if type(value) in (int, float) and math.isfinite(value) else None


def metric(item):
    if not isinstance(item, dict):
        return None
    label = text(item.get('label'))
    if not label:
        return None
    value = item.get('value')
    # Some reviewed NVOS metrics are ratios displayed as "171 / 171".
    if not (type(value) in (int, float) and math.isfinite(value)):
        value = text(value, 80) if isinstance(value, str) and re.fullmatch(r'[\d\s.,%/—–+−-]+', value) else None
    return {'id': text(item.get('id'), 80), 'label': label, 'value': value, 'unit': text(item.get('unit'), 50)}


def source_catalog():
    from services.water_dashboard.config import SOURCES
    # Exclude the retired source even while an old snapshot remains on disk.
    return [s for s in SOURCES if s['id'] != 'flush']


def municipality_names(snapshot):
    names = set()
    for row in snapshot.get('table', []):
        if isinstance(row, dict) and text(row.get('name')):
            names.add(municipality_display(row['name']))
    for sid, source in snapshot.get('sources', {}).items():
        if not isinstance(source, dict):
            continue
        for entity in (source.get('details') or {}).get('entities', []):
            if isinstance(entity, dict):
                name = text(entity.get('municipality')) or (text(entity.get('name')) if sid != 'edo_rso' else '')
                if name:
                    names.add(municipality_display(name))
        for table in source.get('tables', []):
            if not isinstance(table, dict):
                continue
            headers = [normalized(h) for h in table.get('headers', [])]
            name_index = next((i for i, h in enumerate(headers) if h in {'омсу', 'муниципалитет', 'муниципальный округ', 'городской округ'}), None)
            if name_index is None:
                continue
            for row in table.get('rows', []):
                if isinstance(row, list) and len(row) > name_index:
                    name = text(row[name_index])
                    if name and not normalized(name).startswith(('итого', 'всего')):
                        names.add(municipality_display(name))
    return sorted(names, key=normalized)


TABLE_FIELDS = {
    'tasks': {'tasks': 'Просроченные задачи'},
    'sys_vs': {'sysVS': 'Системные адреса ВС', 'resVS': 'Резонансные адреса ВС'},
    'sys_kr': {'sysKR': 'Системные адреса капремонта', 'resKR': 'Резонансные адреса капремонта'},
    'meetings': {'att': 'Явка, %'},
}
# Labels, not arbitrary keys or a raw JSON dump, are the reporting boundary.
SAFE_COLUMN = re.compile(r'омсу|муниципал|городской округ|наименование|организац|\bрсо\b|задвиж|внесено|должно быть|план|факт|корректн|процент|доля|\bэдо\b|\bэцп\b|право подписи|подписант|должностн|документ|присутств|явка|системн|резонанс|кол.во задач|количество задач|просроч', re.I)
UNSAFE_COLUMN = re.compile(r'парол|логин|токен|cookie|секрет|телефон|почт|email|password|token|secret|ключ доступа', re.I)


def selected_tables(source, municipality='', limit=12):
    """Small safe table slices for a requested municipality, never global totals."""
    tables = []
    for table in source.get('tables', []):
        if not isinstance(table, dict):
            continue
        headers = table.get('headers', [])
        indices = [i for i, h in enumerate(headers) if SAFE_COLUMN.search(str(h)) and not UNSAFE_COLUMN.search(str(h))][:10]
        if not indices:
            continue
        muni_index = next((i for i, h in enumerate(headers) if normalized(h) in {'омсу', 'муниципалитет', 'муниципальный округ', 'городской округ'}), None)
        if municipality and muni_index is None:
            continue
        rows = []
        for row in table.get('rows', []):
            if not isinstance(row, list):
                continue
            if municipality and (len(row) <= muni_index or municipality_key(row[muni_index]) != municipality_key(municipality)):
                continue
            rows.append([text(row[i], 180) if i < len(row) else '' for i in indices])
        if rows:
            tables.append({'columns': [text(headers[i]) for i in indices], 'rows': rows[:limit],
                           'matching_rows': len(rows), 'omitted_rows': max(0, len(rows) - limit)})
        if len(tables) >= 3:
            break
    return tables


def safe_entity(entity):
    value = finite_number(entity.get('value'))
    if value is None and isinstance(entity.get('value'), str):
        value = text(entity['value'], 100)
    return {'name': text(entity.get('name')), 'municipality': text(entity.get('municipality')), 'value': value, 'unit': text(entity.get('unit'), 50),
            'secondary': [m for item in entity.get('secondary', [])[:8] if (m := metric(item))]}


def source_reports(snapshot, source_id='', municipality='', limit_rows=12):
    out = []
    stored = snapshot.get('sources', {}) if isinstance(snapshot.get('sources'), dict) else {}
    for spec in source_catalog():
        sid = spec['id']
        if source_id and sid != source_id:
            continue
        data = stored.get(sid) or {}
        valid_metrics = [m for item in data.get('metrics', [])[:20] if (m := metric(item))] if data.get('metric_schema') == 1 else []
        has_values = any(m['value'] is not None for m in valid_metrics)
        report = {'id': sid, 'module': 'water-dashboard', 'title': spec['name'], 'source_url': spec['url'],
                  'status': 'current' if has_values and data.get('ok') else ('stale' if has_values else 'missing'),
                  'collected_at': text(data.get('updated_at'), 50), 'data_date': text(data.get('data_date'), 120),
                  'checked_at': text(data.get('checked_at') or snapshot.get('checked_at'), 50),
                  'warning': text(data.get('error'), 500), 'scope': 'municipality' if municipality else 'region',
                  'metrics': valid_metrics if not municipality else [], 'rows': []}
        if municipality:
            for row in snapshot.get('table', []):
                if not isinstance(row, dict) or municipality_key(row.get('name')) != municipality_key(municipality):
                    continue
                values = [{'label': label, 'value': finite_number(row.get(key)), 'unit': '%' if key == 'att' else ''}
                          for key, label in TABLE_FIELDS.get(sid, {}).items()]
                if any(v['value'] is not None for v in values):
                    report['rows'].append({'municipality': text(row['name']), 'metrics': values})
            report['tables'] = selected_tables(data, municipality, limit_rows)
            if not report['rows'] and not report['tables']:
                report['warning'] = (report['warning'] + ' Нет отдельного показателя по выбранному муниципалитету; областной итог не подставлен.').strip()
        else:
            report['tables'] = selected_tables(data, limit=limit_rows)
        details = data.get('details') or {}
        if isinstance(details, dict) and details.get('schema_version') == 1:
            entities = [safe_entity(e) for e in details.get('entities', []) if isinstance(e, dict)]
            if municipality:
                report['entities'] = [e for e in entities if municipality_key(e.get('municipality') or e['name']) == municipality_key(municipality)]
                if report['entities']:
                    report['warning'] = text(data.get('error'), 500)
            else:
                report['details'] = {
                    'basis': text(details.get('basis'), 500),
                    'coverage_note': text((details.get('coverage') or {}).get('note'), 500),
                    'groups': [{'title': text(g.get('title')), 'items': [safe_entity(e) for e in g.get('items', [])[:5] if isinstance(e, dict)]}
                               for g in details.get('groups', [])[:3] if isinstance(g, dict)],
                }
        out.append(report)
    return out


def build_water_context(limit_rows=12, *, municipality='', source_id='', allowed_modules=None):
    """Compatibility entry point. Permission must be supplied explicitly."""
    if 'water-dashboard' not in set(allowed_modules or []):
        return ''
    return json.dumps({'schema_version': 1, 'sources': source_reports(load_snapshot(), source_id, municipality, limit_rows)}, ensure_ascii=False)
