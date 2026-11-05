"""Import a KOS dashboard snapshot as data. Embedded HTML and scripts are never served."""
from __future__ import annotations

from datetime import date, datetime
import gzip
from html.parser import HTMLParser
import io
import json
import math
import os
from pathlib import Path
import re
import stat
import tempfile
import zipfile
import zlib

STORE = Path(__file__).resolve().parents[2] / 'data/mingkh/treatment/snapshot.json.gz'
MAX_UPLOAD = 24 * 1024 * 1024
MAX_HTML = 64 * 1024 * 1024
TOPICS = {'izliv': 'Излив / сброс стоков', 'smell': 'Запах от очистных', 'kos': 'Работа очистных'}
RADII = {'izliv': 0.5, 'smell': 1.0, 'kos': 0.5}
FIELDS = {
    'kos': 'n omsu okrug name place owner operator state power lat lon check okey',
    'complaints': 'id cur topic why created omsu address subcat fact org status sg deadline changed text lat lon mk',
    'edds': 'id omsu rso type typeName start close dur text objs people mkd n how kind why okey lat lon',
    'records': 'id t prog event start end status sum ready contractor rp stage forecast note n omsu obj people places houses mkd social',
    'payments': 'org recv accr asof',
}


class SnapshotError(ValueError):
    pass


class _Payload(HTMLParser):
    """Read only the two known JS declarations, and only literal JSON operands."""
    def __init__(self):
        super().__init__(convert_charrefs=False)
        self.in_script = False
        self.snapshot_script = False
        self.data = None
        self.state = None

    def handle_starttag(self, tag, attrs):
        if tag == 'script':
            attrs = dict(attrs)
            self.in_script = not attrs.get('src')
            self.snapshot_script = attrs.get('id') == 'snap'

    def handle_endtag(self, tag):
        if tag == 'script':
            self.in_script = False
            self.snapshot_script = False

    def handle_data(self, body):
        if not self.in_script:
            return
        name = 'state' if self.snapshot_script else 'data'
        expression = r'\s*window\.KOS_SNAPSHOT\s*=\s*' if self.snapshot_script else r'\s*const\s+DATA\s*=\s*'
        match = re.match(expression, body)
        if not match:
            return
        if getattr(self, name) is not None:
            raise SnapshotError('В снимке несколько наборов данных.')
        value, end = json.JSONDecoder().raw_decode(body, match.end())
        if not body[end:].lstrip().startswith(';') or not isinstance(value, dict):
            raise SnapshotError('В снимке нет ожидаемого JSON-объекта.')
        setattr(self, name, value)


def _rows(values, group):
    if not isinstance(values, list) or len(values) > 100_000:
        raise SnapshotError('Неверный список записей в снимке.')
    clean = []
    for item in values:
        if not isinstance(item, dict):
            raise SnapshotError('Неверная структура записи в снимке.')
        row = {}
        for key in FIELDS[group].split():
            value = item.get(key)
            if isinstance(value, list) and key == 'objs':
                if len(value) > 1_000 or any(not isinstance(v, str) or len(v) > 2_048 for v in value):
                    raise SnapshotError('Неверный список объектов ЕДДС.')
            elif isinstance(value, str):
                if len(value) > 32_000:
                    raise SnapshotError('Слишком длинное текстовое поле.')
            elif value is not None and (type(value) not in (bool, int, float) or not math.isfinite(value)):
                raise SnapshotError('Некорректное значение в записи снимка.')
            row[key] = value
        if group in ('kos', 'complaints', 'edds'):
            for axis, limit in [('lat', 90), ('lon', 180)]:
                coord = row.get(axis)
                if coord is not None and (type(coord) not in (float, int) or abs(coord) > limit):
                    raise SnapshotError('Недопустимые координаты в снимке.')
        clean.append(row)
    return clean


def _unique(rows, key, expected):
    values = [r[key] for r in rows]
    if any(type(v) is not expected or not v for v in values) or len(set(values)) != len(values):
        raise SnapshotError('В снимке отсутствуют идентификаторы или есть их дубликаты.')


def _nearest(complaints, objects):
    # Match the source's 0.3° / 0.5° candidate window without scanning every
    # object for every complaint; reject contrived unbounded workloads.
    buckets = {}
    for k in objects:
        if k['lat'] is not None and k['lon'] is not None:
            key = (math.floor(k['lat']), math.floor(k['lon']))
            buckets.setdefault(key, []).append((k['n'], k['lat'], k['lon']))
    comparisons = 0
    for row in complaints:
        lat, lon = row['lat'], row['lon']
        row.update(near=None, distance=None)
        if lat is None or lon is None:
            continue
        candidates = (item for a in range(math.floor(lat-.3), math.floor(lat+.3)+1)
                      for b in range(math.floor(lon-.5), math.floor(lon+.5)+1)
                      for item in buckets.get((a, b), ()))
        for number, x, y in candidates:
            comparisons += 1
            if comparisons > 8_000_000:
                raise SnapshotError('Превышен предел обработки координат одного снимка.')
            if abs(x - lat) > 0.3 or abs(y - lon) > 0.5:
                continue
            a = math.sin(math.radians(x - lat) / 2) ** 2
            a += math.cos(math.radians(lat)) * math.cos(math.radians(x)) * math.sin(math.radians(y - lon) / 2) ** 2
            distance = 12742 * math.asin(math.sqrt(min(a, 1)))
            if row['distance'] is None or distance < row['distance']:
                row['near'], row['distance'] = number, distance


def _normalize(data, state):
    try:
        built = datetime.strptime(data['built'], '%d.%m.%Y %H:%M')
        start = date.fromisoformat(data['from'])
        objects = _rows(data['kos'], 'kos')
        complaints = _rows(data['complaints'], 'complaints')
        events = _rows(data.get('edds', []), 'edds')
        _unique(objects, 'n', int)
        _unique(complaints, 'id', int)
        _unique(events, 'id', str)
        if any(r['topic'] not in TOPICS or r['sg'] not in ('closed', 'work') for r in complaints):
            raise SnapshotError('В снимке есть неизвестный тип жалобы или статуса.')
        for row in complaints:
            date.fromisoformat(row['created'][:10])
        overrides = state.get('overrides', {})
        for obj in objects:
            pair = overrides.get(str(obj['n']))
            if pair is not None:
                if (not isinstance(pair, list) or len(pair) != 2 or
                    any(type(v) not in (int, float) or not math.isfinite(v) or abs(v) > limit
                        for v, limit in zip(pair, (90, 180)))):
                    raise SnapshotError('Некорректная поправка координат объекта.')
                obj['lat'], obj['lon'] = pair
        _nearest(complaints, objects)
        extra = data.get('extra') or {}
        snapshot_extra = state.get('extra') or {}
        records = {**extra.get('recs', {}), **snapshot_extra.get('recs', {})}
        removed = {**extra.get('del', {}), **snapshot_extra.get('del', {})}
        records = _rows([v for k, v in records.items() if not removed.get(k)], 'records')
        pay = data.get('pay', {})
        payments = _rows(list(pay.values()), 'payments')
        for key, row in zip(pay, payments):
            row['key'] = key[:256]
        inputs = state.get('inputs') or {}
        radii = {k: float(inputs.get('t:r:' + k, default)) for k, default in RADII.items()}
        if any(not (0 < v <= 100) for v in radii.values()):
            raise SnapshotError('Некорректные радиусы привязки жалоб.')
        return {
            'built': built.strftime('%d.%m.%Y %H:%M'), 'from': start.isoformat(),
            'as_of': built.date().isoformat(), 'topics': TOPICS, 'radii': radii,
            'kos': objects, 'complaints': complaints, 'edds': events,
            'records': records, 'payments': payments,
            'pay_as_of': str(data.get('payAsof', ''))[:40], 'edds_built': str(data.get('eddsBuilt', ''))[:40],
        }
    except SnapshotError:
        raise
    except (KeyError, IndexError, TypeError, AttributeError, ValueError, OverflowError) as exc:
        raise SnapshotError('Архив не соответствует структуре снимка «Жалобы по очистным».') from exc


def parse_archive(body):
    if len(body) > MAX_UPLOAD:
        raise SnapshotError('Архив больше 24 МБ.')
    try:
        with zipfile.ZipFile(io.BytesIO(body)) as archive:
            members = archive.infolist()
            if len(members) != 1:
                raise SnapshotError('В ZIP должен быть один HTML-файл снимка в корне.')
            item = members[0]
            if (item.is_dir() or not re.fullmatch(r'[^/\\:\x00]+\.html?', item.filename, re.I)
                or item.file_size > MAX_HTML or item.flag_bits & 1
                or stat.S_ISLNK(item.external_attr >> 16)):
                raise SnapshotError('Недопустимый путь, тип или размер файла в архиве.')
            with archive.open(item) as source:
                raw = source.read(MAX_HTML + 1)
        if len(raw) > MAX_HTML:
            raise SnapshotError('HTML-файл больше 64 МБ.')
        parser = _Payload()
        parser.feed(raw.decode('utf-8-sig'))
        parser.close()
        if parser.data is None:
            raise SnapshotError('В HTML отсутствует набор данных снимка КОС.')
        return _normalize(parser.data, parser.state or {})
    except SnapshotError:
        raise
    except (zipfile.BadZipFile, UnicodeError, ValueError, RecursionError, NotImplementedError, zlib.error, EOFError) as exc:
        raise SnapshotError('Не удалось прочитать ZIP со снимком КОС.') from exc


def import_archive(body):
    snapshot = parse_archive(body)
    STORE.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=STORE.parent, prefix='.treatment-', delete=False) as tmp:
            temporary = Path(tmp.name)
            with gzip.GzipFile(fileobj=tmp, mode='wb', compresslevel=5) as stream:
                stream.write(json.dumps(snapshot, ensure_ascii=False, separators=(',', ':'), allow_nan=False).encode())
            tmp.flush()
            os.fsync(tmp.fileno())
        os.replace(temporary, STORE)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return {'built': snapshot['built'], 'complaints': len(snapshot['complaints']), 'objects': len(snapshot['kos']),
            'incidents': len(snapshot['edds'])}


def load(include_edds):
    try:
        with gzip.open(STORE, 'rt', encoding='utf-8') as source:
            result = json.load(source)
    except FileNotFoundError:
        return None
    if not include_edds:
        result.pop('edds', None)
        result.pop('edds_built', None)
    return result


def overview(include_edds):
    data = load(include_edds)
    if data is None:
        return None
    # Texts of complaints, street addresses, and incident details are fetched
    # only when a permitted user deliberately opens that particular record.
    for row in data['complaints']:
        for key in ('address', 'text', 'why', 'deadline'):
            row.pop(key, None)
    for row in data.get('edds', []):
        for key in ('text', 'objs', 'why'):
            row.pop(key, None)
    return data


def detail(group, ident):
    data = load(group == 'edds')
    if data:
        return next((row for row in data[group] if str(row['id']) == ident), None)
    return None
