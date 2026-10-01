"""Private water complaint map storage. Only reviewed JSON fields cross the API."""
import gzip
import io
import json
import math
import os
import re
import tempfile
import threading
import time
import zipfile
from datetime import date, datetime, timedelta, timezone
from pathlib import Path, PurePosixPath


DATA_FILE = Path(__file__).resolve().parents[2] / 'data/mingkh/water-map.json.gz'
SEED_FILE = Path(__file__).with_name('water_map_seed.json.gz')
MAX_BYTES = 100 * 1024 * 1024
MAX_RECORDS = 200_000
MOSCOW = timezone(timedelta(hours=3))
MAX_REFRESH_SECONDS = 20 * 60
MAX_COORDINATE_REQUESTS = 2000
COLS = ['id', 'created', 'omsu', 'kind', 'address', 'fact', 'status', 'lat', 'lon', 'geo', 'cat', 'subcat', 'org']
KINDS = {'hvs': 0, 'vo': 1, 'gvs': 2, 'kr': 3, 'etc': 4}
SKIP_STATUS = {'Опубликовано', 'Сбор подписантов', 'На уточнении', 'Закрыто. Отправлено в ЕДС', 'Модерация ЕДС', 'Закрыто пользователем'}
STATUSES = '9,30,32,34,35,37,38,50,51,53,56,60,54,57,111,310,330,511,512,513'
CURATOR = 'Министерство жилищно-коммунального хозяйства Московской области'
SOURCE = 'ДоброДел · ЕЦУР · МинЖКХ · все жалобы · координаты из карточек'
LOCK = threading.Lock()
STATE_LOCK = threading.Lock()
STATE = {'busy': False, 'message': '', 'error': None, 'completed_at': None}
RE_COORDS = re.compile(r'var\s+coords\s*=\s*\[\s*(-?[\d.]+)\s*,\s*(-?[\d.]+)\s*\]')


class StoreBusy(ValueError):
    pass


def _text(value, limit=1000):
    if value is None:
        return ''
    if not isinstance(value, str) or len(value) > limit:
        raise ValueError('В архиве слишком длинное или неверное текстовое поле.')
    return value.strip()


def _created(value):
    value = _text(value, 30)
    try:
        stamp = datetime.fromisoformat(value)
    except ValueError:
        try:
            stamp = datetime.strptime(value, '%d.%m.%Y %H:%M:%S')
        except ValueError:
            raise ValueError('В архиве некорректная дата подачи жалобы.') from None
    return stamp.strftime('%Y-%m-%d %H:%M')


def _point(point):
    if not isinstance(point, dict):
        raise ValueError('В архиве некорректная запись жалобы.')
    cid = point.get('id')
    if isinstance(cid, str) and cid.isascii() and cid.isdecimal():
        cid = int(cid)
    if type(cid) is not int or not 0 < cid < 2**53:
        raise ValueError('В архиве некорректный номер жалобы.')
    kind = point.get('kind')
    if not isinstance(kind, str) or kind not in KINDS:
        raise ValueError('В архиве неизвестный вид водоснабжения.')
    lat, lon = point.get('lat'), point.get('lon')
    if lat is None or lon is None:
        lat = lon = None
        geo = 2 if point.get('geo') == 'none' else 3
    else:
        if any(type(v) not in (int, float) or not math.isfinite(v) for v in (lat, lon)):
            raise ValueError('В архиве некорректные координаты.')
        if not -90 <= lat <= 90 or not -180 <= lon <= 180:
            raise ValueError('Координаты в архиве за допустимыми пределами.')
        lat, lon = round(lat, 6), round(lon, 6)
        geo = 0 if 54 <= lat <= 57.5 and 35 <= lon <= 40.5 else 1
        if lat == 0 and lon == 0:
            lat = lon = None
            geo = 2
    return [cid, _created(point.get('created')), _text(point.get('omsu'), 250) or '—',
            KINDS[kind], _text(point.get('address'), 3000), _text(point.get('fact'), 1500),
            _text(point.get('status'), 250), lat, lon, geo,
            _text(point.get('cat'), 1500), _text(point.get('subcat'), 1500), _text(point.get('org'), 1500)]


def compact(points, meta=None):
    """Drop raw text, names, organisation and all unexpected source fields."""
    if not isinstance(points, dict) or len(points) > MAX_RECORDS:
        raise ValueError('Ожидается словарь points (не более 200 000 жалоб).')
    rows = []
    seen = set()
    omitted = 0
    for point in points.values():
        if isinstance(point, dict) and (_text(point.get('status'), 250) in SKIP_STATUS or point.get('heatMap') or 'не в компетенции' in str(point.get('cat') or '').lower()):
            omitted += 1
            continue
        row = _point(point)
        if row[0] in seen:
            raise ValueError('В архиве повторяется номер жалобы.')
        seen.add(row[0])
        rows.append(row)
    if not rows:
        raise ValueError('В архиве нет подходящих жалоб по воде.')
    rows.sort(key=lambda r: (r[1], r[0]))
    dictionaries = {'omsu': [], 'fact': [], 'status': [], 'cat': [], 'subcat': [], 'org': []}
    indices = {name: {} for name in dictionaries}
    for row in rows:
        for name, col in [('omsu', 2), ('fact', 5), ('status', 6), ('cat', 10), ('subcat', 11), ('org', 12)]:
            value = row[col]
            if value not in indices[name]:
                indices[name][value] = len(dictionaries[name])
                dictionaries[name].append(value)
            row[col] = indices[name][value]
    raw_meta = meta or {}
    metadata = {name: raw_meta[name] for name in ('archive_updated', 'live_updated', 'refreshed_from', 'refreshed_to', 'seed_version') if name in raw_meta}
    metadata.update({'updated': _text(raw_meta.get('updated'), 100), 'from': rows[0][1][:10],
                     'to': rows[-1][1][:10], 'count': len(rows), 'source': SOURCE, 'omitted': omitted,
                     'archive': not bool(raw_meta.get('live_updated'))})
    return {'meta': metadata, 'cols': COLS, **dictionaries, 'rows': rows}


def parse_upload(content):
    """Read a single known JSON member; never extract or execute archive files."""
    if not content or len(content) > MAX_BYTES:
        raise ValueError('Пустой файл или размер больше 100 МБ.')
    if content.startswith(b'PK'):
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                if len(archive.infolist()) > 200:
                    raise ValueError('В архиве слишком много файлов.')
                members = [info for info in archive.infolist()
                           if PurePosixPath(info.filename.replace('\\', '/')).name == 'water_points.json']
                if len(members) != 1:
                    raise ValueError('ZIP должен содержать ровно один файл water_points.json.')
                member = members[0]
                path = PurePosixPath(member.filename.replace('\\', '/'))
                if path.is_absolute() or '..' in path.parts or member.is_dir() or member.flag_bits & 1:
                    raise ValueError('Недопустимый путь или зашифрованный файл в архиве.')
                if member.file_size > MAX_BYTES:
                    raise ValueError('water_points.json внутри архива больше 100 МБ.')
                with archive.open(member) as stream:
                    content = stream.read(MAX_BYTES + 1)
                if len(content) > MAX_BYTES:
                    raise ValueError('Распакованный файл больше 100 МБ.')
        except (zipfile.BadZipFile, RuntimeError, NotImplementedError, EOFError, OSError):
            raise ValueError('Не удалось прочитать ZIP. Загрузите исходный архив или water_points.json.') from None
    try:
        source = json.loads(content.decode('utf-8-sig'))
    except (ValueError, UnicodeError, RecursionError):
        raise ValueError('Файл должен содержать корректный JSON.') from None
    if not isinstance(source, dict):
        raise ValueError('Ожидается хранилище water_points.json.')
    result = compact(source.get('points'), {'updated': source.get('updated')})
    result['meta']['archive_updated'] = result['meta']['updated']
    return result


def save(dataset):
    """One atomic private gzip file; an interrupted write leaves the old map intact."""
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix='.water-map-', dir=DATA_FILE.parent)
    try:
        with os.fdopen(fd, 'wb') as stream:
            with gzip.GzipFile(fileobj=stream, mode='wb', mtime=0) as compressed:
                compressed.write(json.dumps(dataset, ensure_ascii=False, separators=(',', ':'), allow_nan=False).encode('utf-8'))
        os.replace(temporary, DATA_FILE)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


SEED_VERSION = '2026-10-01-all'
_seed_checked = None


def ensure_seed():
    """Merge the new reviewed archive once, keeping newer live server values."""
    global _seed_checked
    if not SEED_FILE.exists():
        return False
    signature = (str(DATA_FILE), DATA_FILE.stat().st_mtime_ns if DATA_FILE.exists() else None)
    if _seed_checked == signature:
        return False
    with LOCK:
        current = None
        if DATA_FILE.exists():
            with gzip.open(DATA_FILE, 'rt', encoding='utf-8') as stream:
                current = json.load(stream)
            if current.get('meta', {}).get('seed_version') == SEED_VERSION:
                _seed_checked = signature
                return False
        with gzip.open(SEED_FILE, 'rt', encoding='utf-8') as stream:
            seed = json.load(stream)
        if seed.get('cols') != COLS or not isinstance(seed.get('rows'), list):
            raise ValueError('Начальный архив карты повреждён.')
        points = _expand(seed)
        meta = dict(seed['meta'])
        if current:
            previous = _expand(current)
            # A live refresh later than the supplied archive is authoritative.
            def stamp(value):
                try:
                    return datetime.strptime(value, '%d.%m.%Y %H:%M')
                except (TypeError, ValueError):
                    return datetime.min
            if stamp(current['meta'].get('live_updated')) > stamp(meta.get('updated')):
                points.update(previous)
                meta.update(current['meta'])
            else:
                previous.update(points)
                points = previous
        meta['seed_version'] = SEED_VERSION
        save(compact(points, meta))
        _seed_checked = (str(DATA_FILE), DATA_FILE.stat().st_mtime_ns)
    return True


def load():
    ensure_seed()
    try:
        with gzip.open(DATA_FILE, 'rt', encoding='utf-8') as stream:
            return json.load(stream)
    except FileNotFoundError:
        raise ValueError('Архив карты ещё не загружен. Администратор может загрузить ZIP или water_points.json на этой странице.') from None


def import_upload(content):
    if not LOCK.acquire(blocking=False):
        raise StoreBusy('Карта уже обновляется. Дождитесь завершения.')
    try:
        dataset = parse_upload(content)
        dataset['meta']['seed_version'] = SEED_VERSION
        save(dataset)
        return dataset['meta']
    finally:
        LOCK.release()


def _expand(dataset):
    kinds = list(KINDS)
    geos = ['ok', 'out', 'none', None]
    return {str(r[0]): {'id': r[0], 'created': r[1], 'omsu': dataset['omsu'][r[2]],
            'kind': kinds[r[3]], 'address': r[4], 'fact': dataset['fact'][r[5]],
            'status': dataset['status'][r[6]], 'lat': r[7], 'lon': r[8], 'geo': geos[r[9]],
            **{key: dataset[key][r[col]] if key in dataset and len(r) > col else ''
               for key, col in [('cat', 10), ('subcat', 11), ('org', 12)]}}
            for r in dataset['rows']}


def water_kind(record):
    value = ' '.join(str(record.get(k) or '') for k in ('subcategory', 'ecurFact', 'ecurCategory'))
    for kind, pattern in [('gvs', r'горяч\w*\s*вод|\bгвс\b'),
                          ('vo', r'водоотвед|канализ|\bстоки\b|очистн|\bкос\b|\bкнс\b|септик|выгреб'),
                          ('hvs', r'холодн\w*\s*вод|\bхвс\b|водоснабж|водопровод|водозабор|водоразбор|колонк|скважин|подвоз\w*\s*вод|качеств\w*\s*вод|ржав')]:
        if re.search(pattern, value, re.I):
            return kind
    if re.search(r'капитальн\w*\s+ремонт', value, re.I):
        return 'kr'
    return 'etc'


def _state(**values):
    with STATE_LOCK:
        STATE.update(values)


def status():
    with STATE_LOCK:
        return dict(STATE)


def refresh(client, today=None):
    """Fetch seven-day overlap and missing tail; preserve older archive and coordinates."""
    started = time.monotonic()
    client.deadline = min(getattr(client, 'deadline', None) or float('inf'), started + MAX_REFRESH_SECONDS)
    def check_deadline():
        if time.monotonic() - started > MAX_REFRESH_SECONDS:
            raise ValueError('Обновление превысило 20 минут. Прежний архив сохранён; повторите позже или загрузите свежий архив.')
    dataset = load()
    points = _expand(dataset)
    today = today or datetime.now(MOSCOW).date()
    meta = dataset['meta']
    last = date.fromisoformat(meta.get('refreshed_to') or meta['to'])
    first = max(date.fromisoformat(meta['from']), last - timedelta(days=7))
    if first > today:
        raise ValueError('Даты архива находятся в будущем. Проверьте дату сервера.')
    if (today - first).days > 120:
        raise ValueError('Архив старше 120 дней. Сначала загрузите свежий архив, затем дозапишите новые жалобы.')
    start = first
    while start <= today:
        check_deadline()
        end = min(today, start + timedelta(days=6))
        _state(message=f'Получаю жалобы {start:%d.%m.%Y} — {end:%d.%m.%Y}…')
        records = client.fetch_all({'filters.curators': CURATOR, 'filters.statuses': STATUSES,
                                    'filters.createdAfter': start.isoformat(), 'filters.createdBefore': (end + timedelta(days=1)).isoformat()})
        seen = set()
        for record in records:
            cid = str(record.get('cardId') or '')
            if not cid.isascii() or not cid.isdecimal():
                raise ValueError('Портал вернул некорректный номер жалобы; сохранён прежний архив.')
            created = _created(record.get('created'))
            if not start.isoformat() <= created[:10] <= end.isoformat():
                continue
            seen.add(cid)
            kind = water_kind(record)
            if not kind or record.get('heatMap') or record.get('status') in SKIP_STATUS or 'не в компетенции' in str(record.get('ecurCategory') or '').lower():
                points.pop(cid, None)
                continue
            point = points.get(cid, {'id': int(cid), 'lat': None, 'lon': None, 'geo': None})
            point.update({'created': created, 'omsu': record.get('district') or '—', 'kind': kind,
                          'address': record.get('address') or '', 'fact': record.get('ecurFact') or '',
                          'status': record.get('status') or '', 'cat': record.get('ecurCategory') or '',
                          'subcat': record.get('subcategory') or '', 'org': record.get('org') or ''})
            _point(point)  # Fail closed on unexpected schema before replacing the file.
            points[cid] = point
        # A large disappearance is more likely an incomplete portal response than a real change.
        inside = [cid for cid, p in points.items() if start.isoformat() <= p['created'][:10] <= end.isoformat()]
        gone = [cid for cid in inside if cid not in seen]
        if len(gone) > 20 and len(gone) > len(inside) * .2:
            raise ValueError('Портал вернул неполный период. Прежний архив сохранён; повторите обновление позже.')
        for cid in gone:
            del points[cid]
        start = end + timedelta(days=1)
    pending = [point for point in points.values() if point.get('geo') is None]
    if len(pending) > MAX_COORDINATE_REQUESTS:
        raise ValueError('Требуются координаты более 2000 жалоб. Загрузите свежий архив; прежние данные сохранены.')
    for index, point in enumerate(pending, 1):
        check_deadline()
        _state(message=f'Получаю координаты: {index} из {len(pending)}…')
        response = client.request('GET', '/Topic', params={'id': point['id']})
        match = RE_COORDS.search(response.text)
        if match:
            point['lat'], point['lon'] = map(float, match.groups())
            point['geo'] = 'ok'
            _point(point)
        elif 'complaintId' in response.text:
            point['geo'] = 'none'
        else:
            raise ValueError('Портал вернул карточку без данных жалобы. Прежний архив сохранён.')
    stamp = datetime.now(MOSCOW).strftime('%d.%m.%Y %H:%M')
    meta.update({'updated': stamp, 'live_updated': stamp, 'refreshed_from': first.isoformat(), 'refreshed_to': today.isoformat()})
    result = compact(points, meta)
    save(result)
    return result['meta']


def start_refresh(credentials):
    ensure_seed()
    if not DATA_FILE.exists():
        raise ValueError('Сначала загрузите архив карты.')
    if not LOCK.acquire(blocking=False):
        raise StoreBusy('Карта уже обновляется. Дождитесь завершения.')
    _state(busy=True, message='Вхожу в ДоброДел…', error=None, completed_at=None)

    def worker():
        try:
            from services.edds.dobrodel import DobrodelClient, DobrodelError
            with DobrodelClient(**credentials) as client:
                client.deadline = time.monotonic() + MAX_REFRESH_SECONDS
                client.login()
                result = refresh(client)
            _state(message=f"Обновлено: {result['count']} жалоб.", completed_at=result['updated'])
        except (ValueError, DobrodelError) as exc:
            _state(error=str(exc), message='Обновление не завершено. Прежний архив сохранён.')
        except Exception:
            _state(error='Не удалось обновить карту. Проверьте соединение и настройки ДоброДела.', message='Прежний архив сохранён.')
        finally:
            _state(busy=False)
            LOCK.release()

    try:
        threading.Thread(target=worker, name='mingkh-water-map', daemon=True).start()
    except Exception:
        _state(busy=False)
        LOCK.release()
        raise
    return status()
