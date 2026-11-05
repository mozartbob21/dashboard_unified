"""House complaint ratings. Input and output are in-memory XLSX workbooks only."""
from __future__ import annotations

import re
import statistics
import zipfile
import zlib
from collections import Counter, defaultdict
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from xml.etree import ElementTree as ET

import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

SHEET_EDS = 'ЕДС'
SHEET_ECUR = 'Тепловая карта ЕЦУР'
GEO_COLUMNS = ('mkd_code', 'address_eds', 'point', 'address_yandex', 'address_yandex_components')
EDS_COLUMNS = ('id заявки', 'Источник', 'Адрес', 'КОД АИС', 'Дата создания')
ECUR_COLUMNS = ('Дубль заявки ЕДС', 'Факт ГВС', 'Подтема', 'Источник', 'Адрес', 'Внутренний Id', 'created')
DEFAULT_BBOX = (37.1, 37.9, 54.95, 55.35)
MAX_FILE_BYTES = 8 * 1024 * 1024
MAX_EXPANDED = 96 * 1024 * 1024
MAX_PART = 48 * 1024 * 1024
MAX_ROWS = 60_000
MAX_GEO_ROWS = 160_000
MAX_COLUMNS = 128
MAX_RATED_HOUSES = 2_000
MAX_UNPLACED = 20_000
MAX_NEIGHBOR_TEXT = 12_000_000
MAX_CELL_TEXT = 32767
THEMES = ('Отопление', 'ГВС', 'Тепловые сети и прочее')
STREET_WORDS = set('улица ул бульвар б-р площадь пл проезд пр-д переулок пер шоссе ш проспект пр-кт тупик туп аллея линия набережная наб'.split())
LOCAL_SUFFIX = re.compile(r'\s+(г|рп|п|д|с|пгт|дп|х|тер|мкр|снт|днп|кв-л)$')
LOCAL_PREFIX = re.compile(r'^(?:поселок городского типа|поселок дома отдыха|дачный поселок|рабочий поселок|поселок|деревня|село|хутор)\s+')
AREA = re.compile(r'^(?:микрорайон|территория|мкр|квартал)\s+|\s+(?:мкр|тер|кв-л)$')
HOUSE = re.compile(r'^(?:д\.?\s*|дом\s*)?\d+[а-яa-z]?(?:(?:[/к]|\s*корп\.?\s*|\s*корпус\s*)\d+[а-я]?)?(?:\s*стр\.?\s*\d+)?$')
DETAIL = re.compile(r'(?<!\w)(?:кв(?:артира|артиры)?|подъезд|этаж|помещение|пом|офис)\b\.?\s*(?:№|номер)?\s*(?=\d|$)', re.I)
PRIVATE_EMAIL = re.compile(r'[\w.+-]+@[\w.-]+', re.I)
PRIVATE_PHONE = re.compile(r'(?<!\d)(?:\+7|8)[\s(-]*(?:\d[\s().-]*){9}\d(?!\d)')
COMPONENT = re.compile(r'kind":"(\w+)","name":"([^"]+)"')
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


class RatingError(ValueError):
    """Safe, source-free user message."""


def _fold(value):
    return ' '.join(str(value).casefold().replace('ё', 'е').split())


def _text(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _redact(value):
    return PRIVATE_PHONE.sub('[телефон скрыт]', PRIVATE_EMAIL.sub('[почта скрыта]', _text(value)))


def _address(value):
    # Only the building-level address may leave the calculator (including the unplaced sheet).
    text = _redact(value)
    match = DETAIL.search(text)
    return (text[:match.start()] if match else text).strip(' ,;').replace('"', '')


def _validate_package(data):
    if not isinstance(data, bytes) or not data or len(data) > MAX_FILE_BYTES:
        raise RatingError('Размер каждой книги XLSX должен быть от 1 байта до 8 МБ.')
    with zipfile.ZipFile(BytesIO(data)) as archive:
        parts = archive.infolist()
        names = {part.filename for part in parts}
        if len(parts) > 1000 or len(names) != len(parts) or sum(p.file_size for p in parts) > MAX_EXPANDED:
            raise RatingError('Книга превышает допустимый объём распакованных данных.')
        if not {'[Content_Types].xml', 'xl/workbook.xml'}.issubset(names):
            raise RatingError('Ожидается книга в формате XLSX.')
        for part in parts:
            name = part.filename.casefold()
            if (part.flag_bits & 1 or part.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED)
                    or part.file_size > MAX_PART or name.startswith('/') or '..' in name.split('/') or '\\' in name):
                raise RatingError('Книга содержит неподдерживаемый или слишком большой компонент.')
            if name.endswith('.bin') or name.startswith(('xl/externallinks/', 'xl/embeddings/')):
                raise RatingError('Нужна обычная книга XLSX без макросов, вложенных объектов и внешних книг.')
            # Relationship targets can have nonstandard extensions. Scan every
            # part before openpyxl can pass any workbook data to an XML parser.
            with archive.open(part) as raw:
                previous = b''
                while chunk := raw.read(131072):
                    scan = previous + chunk.replace(b'\x00', b'').upper()
                    if b'<!DOCTYPE' in scan or b'<!ENTITY' in scan:
                        raise RatingError('Объявления сущностей и DTD в книгах не поддерживаются.')
                    previous = scan[-16:]
            if name == 'xl/sharedstrings.xml':
                # openpyxl loads the entire shared string table even in read-only mode.
                if part.file_size > 24 * 1024 * 1024:
                    raise RatingError('Справочник строк книги превышает допустимый размер.')
                count = 0
                with archive.open(part) as raw:
                    for _, node in ET.iterparse(raw, events=('end',)):
                        if node.tag == NS + 'si':
                            count += 1
                            if count > 200_000:
                                raise RatingError('Слишком много строк в общем справочнике книги.')
                        node.clear()
            elif name.endswith(('.xml', '.rels')) and not name.startswith('xl/worksheets/') and part.file_size > 2 * 1024 * 1024:
                raise RatingError('Служебные таблицы книги превышают допустимый размер.')


def _open_book(data):
    _validate_package(data)
    return load_workbook(BytesIO(data), read_only=True, data_only=False, keep_links=False)


def _records(sheet, required, limit):
    # Some exports set the declared dimensions to A1 even when data follow.
    sheet.reset_dimensions()
    rows = sheet.iter_rows(max_col=MAX_COLUMNS)
    header = next(rows, ())
    names = [_text(cell.value) for cell in header]
    absent = [column for column in required if column not in names]
    if absent:
        raise RatingError('В таблице отсутствуют обязательные колонки: ' + ', '.join(absent) + '.')
    if any(names.count(col) != 1 for col in required):
        raise RatingError('Названия обязательных колонок должны быть уникальными.')
    indexes = [names.index(column) for column in required]
    for number, row in enumerate(rows, start=1):
        if number > limit:
            raise RatingError('В книге слишком много строк. Сократите исходную выгрузку.')
        cells = [row[index] for index in indexes]
        if any(cell.data_type in ('f', 'e') for cell in cells):
            raise RatingError('В расчётных полях входной книги нужны значения без формул и ошибок.')
        record = {col: _text(cell.value) for col, cell in zip(required, cells)}
        if any(len(value) > 4096 for value in record.values()):
            raise RatingError('В обязательном поле встретилось слишком длинное значение.')
        if any(record.values()) or any(cell.value is not None and cell.value != '' for cell in row):
            # A complaint with only its description (outside the projection)
            # still counts as an unassigned complaint. Never export descriptions.
            yield record


def _street(value):
    return ' '.join(sorted(w for w in _fold(value).replace('.', ' ').split() if w not in STREET_WORDS))


def _house(value):
    compact = _fold(value).replace(' ', '').replace('дом', '').replace('д.', '')
    return re.sub(r'корпус|корп\.?', 'к', compact).replace('/', 'к')


def _eds_address(address):
    tokens = [_fold(t) for t in address.split(',') if t.strip()]
    locality = LOCAL_SUFFIX.sub('', tokens[0]) if tokens else ''
    street = house = ''
    for token in tokens[1:]:
        if HOUSE.fullmatch(token):
            house = _house(token)
        elif re.search(r'\s(мкр|тер|кв-л|снт)$', token):
            continue
        elif re.search(r'\s(рп|п|д|с|пгт|дп|х)$', token):
            locality = LOCAL_SUFFIX.sub('', token)
        else:
            street = _street(token)
    return locality, street, house


def _ecur_address(address, city):
    locality = area = street = house = ''
    country = {'россия', 'московская область', 'муниципальный округ ' + city, 'городской округ ' + city}
    for token in (_fold(v) for v in _address(address).split(',') if v.strip()):
        if token in country or re.fullmatch(r'\d{6}', token):
            continue
        if token == city:
            locality = locality or city
        elif LOCAL_PREFIX.match(token):
            locality = LOCAL_PREFIX.sub('', token)
        elif AREA.search(token):
            area = AREA.sub('', token)
        elif HOUSE.fullmatch(token):
            house = _house(token)
        else:
            street = _street(token)
    return locality, area, street, house


def _yandex_address(components):
    fields = dict(COMPONENT.findall(components))
    return (LOCAL_PREFIX.sub('', _fold(fields.get('locality', ''))),
            _street(fields.get('street', '')), _house(fields.get('house', '')))


@dataclass
class House:
    address: str
    coords: tuple | None
    note: str


def _index_add(index, key, code):
    if key[0] and key[2]:
        if key in index and index[key] != code:
            index[key] = None
        else:
            index.setdefault(key, code)


def _directory(sheet, city, bbox):
    houses, eds_index, ya_index = {}, {}, {}
    for record in _records(sheet, GEO_COLUMNS, MAX_GEO_ROWS):
        address = _address(record['address_eds'])
        code = record['mkd_code']
        if not _fold(address).startswith(city) or not code or code in houses:
            continue
        place = _address(record['address_yandex'])
        try:
            lon, lat = (float(t) for t in record['point'].split())
            valid = (bbox[0] <= lon <= bbox[1] and bbox[2] <= lat <= bbox[3] and city in _fold(place))
        except ValueError:
            valid = False
        note = '' if valid else 'Точка отсутствует или не подтверждена в границах округа. Расстояние не считалось.'
        houses[code] = House(address,(lon,lat) if valid else None,note)
        _index_add(eds_index, _eds_address(address), code)
        _index_add(ya_index, _yandex_address(record['address_yandex_components']), code)
    if not houses:
        raise RatingError('В геосправочнике нет домов выбранного округа.')
    return houses, eds_index, ya_index


def _match(parts, eds, ya, city):
    locality, area, street, number = parts
    if not number:
        return None, 'нет номера дома'
    locations = [locality] if locality else [city] + ([area] if area else [])
    for index in (eds, ya):
        for place in locations:
            key = (place, street, number)
            if key in index:
                code = index[key]
                return (code,'точно') if code else (None,'неоднозначно')
    return None, 'нет в справочнике'


def _day(value):
    s = value[:10]
    for fmt in ('%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _period(days):
    if not days:
        return 'даты не указаны'
    return f'{min(days):%d.%m.%Y}–{max(days):%d.%m.%Y}'


def _aggregate(complaints, houses, eds_index, ya_index, city):
    if not {SHEET_EDS, SHEET_ECUR} <= set(complaints.sheetnames):
        raise RatingError('В книге жалоб нужны оба листа: «ЕДС» и «Тепловая карта ЕЦУР».')
    counts = defaultdict(Counter)
    incidents, seen_ids = set(), set()
    dates = [set(), set()]
    stats = dict(rows_eds=0,rows_ecur=0,dup_flagged=0,dup_hidden=0,unplaced=0,blank_address=0)
    for row in _records(complaints[SHEET_EDS], EDS_COLUMNS, MAX_ROWS):
        stats['rows_eds'] += 1
        code = row['КОД АИС']
        if code not in houses:
            raise RatingError('Код дома из ЕДС отсутствует в справочнике выбранного округа. Проверьте обе выгрузки.')
        counts[code][THEMES[0]] += 1
        if _fold(row['Источник']) == 'инцидент' and row['id заявки'].lstrip('-'):
            incidents.add(row['id заявки'].lstrip('-'))
        if day := _day(row['Дата создания']):
            dates[0].add(day)
    unplaced = []
    unplaced_text = 0
    for row in _records(complaints[SHEET_ECUR], ECUR_COLUMNS, MAX_ROWS):
        stats['rows_ecur'] += 1
        ident = row['Внутренний Id']
        if day := _day(row['created']):
            dates[1].add(day)
        if row['Дубль заявки ЕДС']:
            stats['dup_flagged'] += 1
            continue
        if ident and ident in incidents:
            stats['dup_hidden'] += 1
            continue
        if ident and ident in seen_ids:
            raise RatingError('В ЕЦУР повторяется внутренний идентификатор без пометки дубля. Проверьте выгрузку.')
        seen_ids.add(ident)
        code, reason = _match(_ecur_address(row['Адрес'],city), eds_index, ya_index, city)
        if code:
            subtopic = row['Подтема']
            theme = THEMES[1] if ('горячее водоснабжение (гвс) в мкд' == _fold(subtopic) or row['Факт ГВС']) else (
                THEMES[0] if _fold(subtopic).startswith('отопление') else THEMES[2])
            counts[code][theme] += 1
        else:
            if len(unplaced) >= MAX_UNPLACED:
                raise RatingError('Слишком много непривязанных обращений. Сократите период или уточните адреса.')
            if not row['Адрес']:
                reason = 'адрес пуст'
                stats['blank_address'] += 1
            creation_date = _day(row['created'])
            record = (reason.capitalize(), _address(row['Адрес']), _address(row['Подтема']),
                      _address(row['Источник']), creation_date.isoformat() if creation_date else '',
                      _redact(ident))
            unplaced_text += sum(map(len, record))
            if unplaced_text > 12_000_000:
                raise RatingError('Список непривязанных обращений слишком велик. Сократите период.')
            unplaced.append(record)
    stats['unplaced'] = len(unplaced)
    stats['unique_complaints'] = stats['rows_eds'] + stats['rows_ecur'] - stats['dup_flagged'] - stats['dup_hidden']
    stats['period_eds'], stats['period_ecur'] = map(_period,dates)
    return counts, unplaced, stats


def _rank(counts,houses,heat,radius,wide,city):
    totals = {k:(v[THEMES[0]] if heat else sum(v.values())) for k, v in counts.items()}
    codes = sorted((k for k,n in totals.items() if n), key=lambda k: (-totals[k], houses[k].address))
    if len(codes) > MAX_RATED_HOUSES:
        raise RatingError('В одной выборке допускается до 2 000 домов с жалобами. Сократите период.')
    if len({houses[k].address for k in codes}) != len(codes):
        raise RatingError('В справочнике один адрес соответствует нескольким кодам домов с жалобами.')
    good = sorted(k for k in codes if houses[k].coords is not None)
    longitude = np.radians([houses[k].coords[0] for k in good])
    latitude = np.radians([houses[k].coords[1] for k in good])
    position = {code: i for i,code in enumerate(good)}
    rows = []
    previous = None
    used_text = 0
    for i, code in enumerate(codes,1):
        n = totals[code]
        if n != previous:
            rank = i
            previous = n
        h = houses[code]
        row = dict(rank=rank,code=code,address=h.address,count=n,distance=None,nearest=None,nearest_count=None,
                   neighbors=None,neighbor_complaints=None,neighbors_wide=None,neighbor_list='',note=h.note)
        if code in position:
            p = position[code]
            a = np.sin((latitude-latitude[p])/2)**2 + np.cos(latitude)*np.cos(latitude[p])*np.sin((longitude-longitude[p])/2)**2
            distances = 2*6371000*np.arcsin(np.sqrt(np.clip(a,0,1)))
            distances[p] = np.inf
            near = np.flatnonzero(distances <= radius)
            near = near[np.argsort(distances[near],kind='stable')]
            near_list = []
            for j in near:
                address = re.sub(r'^'+re.escape(city)+r' г,\s*', '', houses[good[j]].address, flags=re.I)
                near_list.append(f'{address} ({round(distances[j])} м)')
            detail = '; '.join(near_list)
            used_text += len(detail)
            if len(detail)>MAX_CELL_TEXT or used_text>MAX_NEIGHBOR_TEXT:
                raise RatingError('Список соседей слишком велик для книги. Уменьшите радиус или сократите период.')
            row.update(neighbors=len(near),neighbor_complaints=sum(totals[good[j]] for j in near),
                       neighbors_wide=int(np.count_nonzero(distances<=wide)),neighbor_list=detail)
            if len(good)>1:
                j = int(np.argmin(distances))
                row.update(distance=round(distances[j]),nearest=houses[good[j]].address,nearest_count=totals[good[j]])
        rows.append(row)
    return rows


def _self_check(rows, houses, radius, wide):
    """Independently compute spherical angles using unit vectors, one house at a time."""
    candidates = [r for r in rows if houses[r['code']].coords]
    if not candidates:
        return
    lon, lat = np.radians(np.array([houses[r['code']].coords for r in candidates])).T
    vectors = np.column_stack((np.cos(lat)*np.cos(lon),np.cos(lat)*np.sin(lon),np.sin(lat)))
    for i,row in enumerate(candidates):
        # atan2(cross,dot) remains well-conditioned for two close coordinates.
        distances = 6371000*np.arctan2(np.linalg.norm(np.cross(vectors,vectors[i]),axis=1),vectors @ vectors[i])
        distances[i] = np.inf
        nearest = None if len(vectors)==1 else float(min(distances))
        if nearest is None:
            if row['distance'] is not None or row['nearest'] is not None:
                raise RatingError('Проверка ближайшего дома не пройдена.')
        elif (row['distance'] is None or abs(row['distance']-nearest)>0.501
              or not any(candidate['address']==row['nearest'] and abs(d-nearest)<0.001
                         for candidate,d in zip(candidates,distances))):
            raise RatingError('Проверка расстояния до ближайшего дома не пройдена.')
        for limit,key in ((radius,'neighbors'),(wide,'neighbors_wide')):
            if not np.sum(distances<=limit-1e-5)<=row[key]<=np.sum(distances<=limit+1e-5):
                raise RatingError('Проверка числа домов в радиусе не пройдена.')


_HEAD_FILL = PatternFill('solid', fgColor='1F4E78')
_WARN_FILL = PatternFill('solid', fgColor='FFF2CC')
_THIN = Side(style='thin', color='D9D9D9')
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_FONT = Font(name='Arial', size=10)
_BOLD = Font(name='Arial', size=10, bold=True)
_ALIGN = Alignment(vertical='top')
_WRAP = Alignment(vertical='top', wrap_text=True)


def _header(ws, headers, widths, freeze):
    for j,(name,width) in enumerate(zip(headers,widths),1):
        cell = ws.cell(1,j,name)
        cell.font = Font(name='Arial',size=10,bold=True,color='FFFFFF')
        cell.fill = _HEAD_FILL
        cell.alignment = Alignment(horizontal='center',vertical='center',wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = width
    ws.row_dimensions[1].height = 62
    ws.freeze_panes = freeze
    ws.sheet_view.zoomScale = 90


def _cell(ws, i, j, value, *, warn=False, wrap=False):
    if value == '':
        value=None
    cell=ws.cell(i,j,value)
    if isinstance(value,str):
        # openpyxl treats strings beginning with "=" and exact Excel errors specially.
        cell.data_type='s'
    cell.font = _BOLD if j == 3 and ws.title != 'Не привязаны' else _FONT
    cell.border = _BORDER
    cell.alignment = _WRAP if wrap else _ALIGN
    if warn:
        cell.fill = _WARN_FILL
    return cell


def _formula(cell, expression, computed, cache):
    cell.value=expression
    cache[cell.coordinate]=int(computed)


def _cache_formulas(data, caches):
    """OOXML formula results for preview clients, without starting Office or executing formulas."""
    output = BytesIO()
    with zipfile.ZipFile(BytesIO(data)) as src,zipfile.ZipFile(output,'w',compression=zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            raw=src.read(item)
            if item.filename in caches:
                cache=caches[item.filename]
                # Only XML created by the writer is processed here. Set <v> on known numeric formula cells.
                found=set()
                def replace(match):
                    coordinate=match[2].decode('ascii')
                    if coordinate not in cache:
                        raise RatingError('Не удалось подготовить значения формул.')
                    found.add(coordinate)
                    return match[1]+str(cache[coordinate]).encode('ascii')+b'</v>'
                raw=re.sub(rb'(<c r="([A-Z]+\d+)"[^>]*><f>[^<]+</f><v>)</v>',replace,raw)
                if found!=set(cache):
                    raise RatingError('Не удалось сохранить все значения формул.')
            dst.writestr(item,raw)
    return output.getvalue()


def _write(main, heating, counts, unplaced, stats, radius, wide):
    wb=Workbook()
    ws=wb.active
    ws.title='Рейтинг'
    wh=wb.create_sheet('Только отопление')
    wu=wb.create_sheet('Не привязаны')
    wm=wb.create_sheet('Как считали')
    cache={name:{} for name in (ws.title,wh.title,wm.title)}
    neighbors=[f'Домов с жалобами рядом (до {radius} м)',f'Жалоб в этих домах (до {radius} м)',
               f'Домов с жалобами до {wide} м',f'Какие дома рядом (до {radius} м)','Примечание']
    _header(ws,['Место','Дом (адрес)','Жалоб всего',*THEMES,'До ближайшего дома с жалобами, м','Ближайший дом с жалобами','Жалоб у ближайшего',*neighbors],
            [7,40,8,10,7,10,12,40,10,11,10,10,70,45],'C2')
    _header(wh,['Место','Дом (адрес)','Жалоб на отопление','До ближайшего дома с жалобами на отопление, м','Ближайший дом','Жалоб у ближайшего',*neighbors],
            [7,40,10,14,40,10,11,10,10,70,45],'C2')
    for sheet, rows in ((ws,main),(wh,heating)):
        for i,row in enumerate(rows,2):
            values=[None,row['address'],row['count']]
            if sheet is ws:
                values.extend(counts[row['code']][topic] for topic in THEMES)
            values.extend(row[key] for key in ('distance','nearest','nearest_count','neighbors','neighbor_complaints','neighbors_wide','neighbor_list','note'))
            wrap_cols=(2,8,13,14) if sheet is ws else (2,5,10,11)
            for j,val in enumerate(values,1):
                _cell(sheet,i,j,val,warn=bool(row['note']),wrap=j in wrap_cols)
            _formula(sheet.cell(i,1),f'=RANK(C{i},$C$2:$C${len(rows)+1})',row['rank'],cache[sheet.title])
            if sheet is ws:
                _formula(sheet.cell(i,3),f'=SUM(D{i}:F{i})',row['count'],cache[sheet.title])
        sheet.auto_filter.ref=sheet.dimensions
    _header(wu,['Причина','Адрес в ЕЦУР','Подтема','Источник','Создана','Внутренний Id'],[20,70,34,18,12,13],'A2')
    for i,record in enumerate(sorted(unplaced,key=lambda r:(r[0],r[1])),2):
        for j,value in enumerate(record,1):
            _cell(wu,i,j,value,wrap=j in (1,2))
    wu.auto_filter.ref=wu.dimensions
    wm.column_dimensions['A'].width=64
    wm.column_dimensions['B'].width=16
    wm.column_dimensions['C'].width=80
    last=max(2,len(main)+1)
    methodology=[
        ('Сверка количества жалоб',None,None),
        ('Строк на листе ЕДС', stats['rows_eds'],stats['period_eds']),
        ('Строк на листе «Тепловая карта ЕЦУР»',stats['rows_ecur'],stats['period_ecur']),
        ('Минус: помечены «Дубль заявки ЕДС»',-stats['dup_flagged'],'Пометка из исходного файла.'),
        ('Минус: скрытые дубли инцидентов',-stats['dup_hidden'],'Внутренний Id ЕЦУР совпадает с id инцидента ЕДС без минуса.'),
        ('Уникальных жалоб',stats['unique_complaints'],None),
        ('Минус: нельзя привязать к дому',-stats['unplaced'],'Подробности на листе «Не привязаны».'),
        ('Жалоб в рейтинге (расчёт)',stats['in_rating'],None),
        ('Жалоб в рейтинге (сумма по листу «Рейтинг»)',stats['in_rating'],'Должно совпасть с предыдущей строкой.'),
        ('Домов в рейтинге',stats['houses'],None),
        ('Из них с неподтверждённой точкой',stats['bad_point_houses'],'Выделены жёлтым.'),
        (None,None,None),('Правила расчёта',None,None),
        ('Расстояние',None,'По прямой от центра до центра дома, гаверсинус, радиус Земли 6 371 000 м.'),
        ('Рядом',None,f'Основной радиус {radius} м, дополнительный {wide} м. Медиана расстояния до ближайшего другого дома: '
         +(f"{stats['median_nearest_m']} м." if stats['median_nearest_m'] is not None else 'недостаточно домов.')),
        ('Сосед',None,'Другой дом с жалобами и подтверждённой точкой. На листе «Только отопление» круг соседей пересчитан.'),
        ('Привязка',None,'ЕДС — по «КОД АИС». ЕЦУР — по населённому пункту, улице и дому. Адрес ЕДС имеет приоритет перед Яндексом.'),
        ('Темы',None,'Все строки ЕДС — отопление. ЕЦУР распределён по подтеме и признаку «Факт ГВС».'),
        ('Тепловые сети',None,'Часть постов из соцсетей связана с домом по упоминанию.'),
        ('Пустой адрес',None,f"{stats['blank_address']} обращений без адреса. Тексты обращений для ручной привязки не использованы.")
    ]
    for i,row in enumerate(methodology,1):
        for j,val in enumerate(row,1):
            cell=_cell(wm,i,j,val,wrap=True)
            if i in (1,13):
                cell.fill=_HEAD_FILL
                cell.font=Font(name='Arial',size=10,bold=True,color='FFFFFF')
        if i>=13:
            wm.row_dimensions[i].height=32
    for i,expr in ((6,'SUM(B2:B5)'),(8,'SUM(B6:B7)'),(9,f'SUM(Рейтинг!C2:C{last})'),(10,f'COUNTA(Рейтинг!B2:B{last})'),(11,f'COUNTA(Рейтинг!N2:N{last})')):
        _formula(wm.cell(i,2),'='+expr,methodology[i-1][1],cache[wm.title])
    output=BytesIO()
    wb.save(output)
    wb.close()
    return _cache_formulas(output.getvalue(),{'xl/worksheets/sheet'+str(i)+'.xml':cache[sh.title]
                           for i,sh in enumerate((ws,wh,wu,wm),1) if sh.title in cache})


def _parameters(city, bbox, radius, radius_wide):
    if not isinstance(city,str) or not re.fullmatch(r'[а-яёА-ЯЁa-zA-Z][а-яёА-ЯЁa-zA-Z0-9 .-]{0,79}',city.strip()):
        raise RatingError('Укажите корректное название города или округа.')
    city=city.strip()
    if bbox is None or bbox=='':
        if _fold(city)!='чехов':
            raise RatingError('Для другого округа укажите границы: минимум и максимум долготы, затем широты.')
        bbox=DEFAULT_BBOX
    if isinstance(bbox,str):
        bbox=bbox.split(',')
    try:
        x1,x2,y1,y2=map(float,bbox)
        if not -180<=x1<x2<=180 or not -90<=y1<y2<=90:
            raise ValueError
    except (TypeError,ValueError,OverflowError):
        raise RatingError('Некорректные границы. Нужны lon_min, lon_max, lat_min, lat_max.') from None
    if type(radius) is not int or type(radius_wide) is not int or not 1<=radius<radius_wide<=50_000:
        raise RatingError('Радиусы должны быть целыми числами от 1 до 50 000 м; дополнительный радиус должен быть больше основного.')
    return city,(x1,x2,y1,y2)


def build_rating(geo: bytes, complaints: bytes, *, city='Чехов', bbox=None, radius=100, radius_wide=200):
    """Return (verified XLSX bytes, bounded summary) or raise RatingError."""
    city,box=_parameters(city,bbox,radius,radius_wide)
    try:
        with _book_context(geo) as geobook:
            if not geobook.worksheets:
                raise RatingError('В геосправочнике нет листов.')
            houses,eds_idx,ya_idx=_directory(geobook.worksheets[0],_fold(city),box)
        with _book_context(complaints) as book:
            counts,unplaced,stats=_aggregate(book,houses,eds_idx,ya_idx,_fold(city))
    except RatingError:
        raise
    except (ValueError,KeyError,TypeError,IndexError,AttributeError,ET.ParseError,zipfile.BadZipFile,
            InvalidFileException,EOFError,zlib.error,OverflowError,OSError,UnicodeError,NotImplementedError):
        raise RatingError('Не удалось прочитать книги XLSX. Проверьте их формат и структуру.') from None
    main=_rank(counts,houses,False,radius,radius_wide,_fold(city))
    heat=_rank(counts,houses,True,radius,radius_wide,_fold(city))
    distances=[r['distance'] for r in main if r['distance'] is not None]
    stats.update(in_rating=sum(r['count'] for r in main),houses=len(main),houses_heat_only=len(heat),
                 bad_point_houses=sum(bool(r['note']) for r in main),
                 median_nearest_m=int(statistics.median(distances)) if distances else None)
    if stats['in_rating'] + stats['unplaced'] != stats['unique_complaints']:
        raise RatingError('Сверка жалоб не пройдена. Файл не сформирован.')
    _self_check(main,houses,radius,radius_wide)
    _self_check(heat,houses,radius,radius_wide)
    output=_write(main,heat,counts,unplaced,stats,radius,radius_wide)
    with _book_context(output, output=True) as verify:
        # Re-read numeric formula caches, row counts and the independent reconciliation.
        summary=verify['Как считали']
        values=[r[0] for r in summary.iter_rows(min_row=6,max_row=11,min_col=2,max_col=2,values_only=True)]
        if (values!=[stats['unique_complaints'],-stats['unplaced'],stats['in_rating'],stats['in_rating'],
                     stats['houses'],stats['bad_point_houses']] or verify['Не привязаны'].max_row!=1+stats['unplaced']):
            raise RatingError('Сверка готового файла не пройдена.')
        for title,rows in (('Рейтинг',main),('Только отопление',heat)):
            result=verify[title]
            if result.max_row!=len(rows)+1 or any(v[0:3]!=(r['rank'],r['address'],r['count'])
                    for v,r in zip(result.iter_rows(min_row=2,values_only=True),rows)):
                raise RatingError('Сверка строк готового рейтинга не пройдена.')
    return output,{'city':city,'radius':radius,'radius_wide':radius_wide,'stats':stats,
                   'top10':[{k:r[k] for k in ('rank','address','count','distance','nearest','neighbors')} for r in main[:10]],
                   'bad_points':[{'address':r['address'],'count':r['count']} for r in main if r['note']][:100],
                   'checks':{'reconciliation':True,'distances':True,'workbook':True}}


@contextmanager
def _book_context(data, output=False):
    book=load_workbook(BytesIO(data),read_only=True,data_only=True,keep_links=False) if output else _open_book(data)
    try:
        yield book
    finally:
        book.close()
