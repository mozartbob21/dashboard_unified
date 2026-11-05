"""Deterministic row lookup in the same saved results used by platform reports."""
from __future__ import annotations

from datetime import datetime
from functools import lru_cache
import re
import unicodedata

from fastapi import HTTPException

from core.roles import check_module_access
from services.aichat import operational_lookup as lookup, report_context, report_periods
from services.max_bot import stock

SUPPORTED = frozenset(module for module, (_, path) in report_context.MODULES.items() if path)
# Free-form subjects, explanations and attachments from reports never enter
# this more broadly visible search output. Keep only named operational fields.
_LABELS = (
    'Номер задачи', 'Номер', 'Муниципалитет', 'Организация', 'Наименование', 'Объект',
    'Адрес', 'Действие', 'Статус камеры', 'Статус', 'Категория', 'Тип обращения',
    'Тип заявки', 'Дата', 'Срок исполнения', 'Проверено', 'Дней просрочки',
    'Просроченные задачи', 'Незаполненные поля', 'Отсутствующие даты',
)
_ALIASES = {
    'edo':r'заполненн\w*\s+данн\w*|эдо',
    'watercontrol':r'контрол\w*\s+вод\w*|watercontrol',
    'overdue':r'просроченн\w*\s+задач\w*|просрочк\w*|overdue',
    'utnkr':r'(?:технадзор\w*\s+)?утнкр|технадзор\w*',
    'cameras':r'(?:проверк\w*\s+)?камер\w*',
    'cds':r'обращени\w*\s+(?:1с|1c)|цдс|1с|1c',
    'mgkh_rm':r'(?:жкх|минжкх)\s+redmine|redmine|р[еэ]дмайн',
}
_CAMERA_LABELS = {'working':'работает','not_working':'не работает','not_connected':'не подключена','unknown':'статус неизвестен'}
_CAMERA_FILTERS = (
    (r'не\s*подключ\w*|отключенн\w*', 'not_connected'),
    (r'не\s*работа\w*|неисправн\w*|без\s+видеопотока', 'not_working'),
    (r'работающ\w*|исправн\w*|работают|работает', 'working'),
    (r'неизвестн\w*\s+статус\w*|статус\w*\s+неизвест\w*','unknown'),
)
_NUMBER = re.compile(r'\b(?:номер(?:ом|у|а)?|id|задач[ауи]|обращение)\s*[#:№]?\s*([a-zа-я\d-]*\d[\w./-]*)',re.I)
_ROW_KEYS = ('rows','items','violators','cameras','results','data','records')


def _answer(module, status, *, note, as_of=None, items=None, total=None):
    return {'scope':module, 'title':report_context.MODULES[module][0], 'url':'/'+module.replace('_','-'),
            'status':status, 'note':note, 'as_of':as_of, 'items':items or [], 'total':total}


def _shape(data):
    if isinstance(data,list):
        return data
    if not isinstance(data,dict):
        return None
    if 'buckets' in data:
        buckets=data['buckets']
        if not isinstance(buckets,dict) or not all(isinstance(buckets.get(key),list) for key in ('close','extend','rework')):
            return None
        return [row for key in ('close','extend','rework') for row in buckets[key]]
    return next((data[key] for key in _ROW_KEYS if isinstance(data.get(key),list)), None)


def _project(module, raw):
    # Map the current normalized source keys onto the report's existing
    # whitelist before any free-word search. Never stringify a raw row.
    data = {key:raw[key] for key in report_context.FIELDS if key in raw}
    if module in ('cameras','mgkh_rm') and 'id' in raw:
        data['Номер'] = raw['id']
    if module == 'cds':
        for source,label in (('number','Номер'),('date','Дата'),('type','Тип обращения'),
                             ('request_type','Тип заявки'),('deadline','Срок исполнения')):
            if source in raw:
                data[label] = raw[source]
    if module == 'edo' and isinstance(raw.get('empty_fields'),list):
        data['missing_fields'] = raw['empty_fields']
    if module == 'cameras':
        data.pop('owner',None)  # The owner field need not name an organization.
    projected = report_context._safe_row(data)
    safe = {}
    for field in _LABELS:
        value = projected.get(field)
        # Residential addresses attached to citizen requests stay in the CDS UI.
        if (module == 'cds' and field == 'Адрес') or value is None or value == '':
            continue
        if field in ('Номер','Номер задачи') and not re.fullmatch(r'[\w./-]{1,70}',str(value)):
            continue
        cleaned = lookup._plain(str(value),150)
        if cleaned:
            safe[field] = cleaned
    return safe


def _date(value):
    if not isinstance(value,str) or not value:
        return None
    try:
        datetime.fromisoformat(value.replace('Z','+00:00'))
        return value
    except ValueError:
        for fmt in ('%d.%m.%Y %H:%M:%S','%d.%m.%Y %H:%M','%d.%m.%Y'):
            try:
                return datetime.strptime(value,fmt).isoformat()
            except ValueError:
                continue
    return None


@lru_cache(maxsize=4096)
def _word_forms(word):
    if len(word) >= 4 and len(word) <= 60 and re.fullmatch(r'[а-я]+', word):
        return frozenset(report_context._word_variants(word))
    return frozenset((word,))


def _matches(words, row):
    tokens = set(stock._tokens(' '.join(row.values())))
    missing = set(words) - tokens
    # Russian case endings for words such as «школу» / «школа». Identifiers
    # and dimensions always require exact tokens.
    return not missing or all(term.isalpha() and any(term in _word_forms(token) for token in tokens)
                              for term in missing)


def search(user, module, query, limit=5):
    if not user:
        raise HTTPException(401,'Требуется авторизация.')
    if module not in SUPPORTED or not check_module_access(user,module):
        raise HTTPException(403,'Нет доступа к данным этого блока.')
    if not isinstance(query,str) or not query.strip() or len(query)>1200 or type(limit) is not int or not 1<=limit<=20:
        raise HTTPException(400,'Уточните поисковый запрос.')
    question = unicodedata.normalize('NFKC',query.replace('№',' номер ')).casefold().replace('ё','е')
    if re.search(r'\b(?:парол\w*|токен\w*|почт\w*|телефон\w*|контакт\w*|заявител\w*)\b',question):
        return _answer(module,'clarify',note='Поиск в чате показывает операционные поля блока. Контакты и личные сведения в выдачу не входят.')
    try:
        data, error = report_context._load_result(module)
    except (RecursionError, OverflowError):
        data, error = None, 'Не удалось прочитать сохранённый результат.'
    raw_rows = _shape(data)
    if raw_rows is None or len(raw_rows)>300_000 or any(not isinstance(row,dict) for row in raw_rows):
        return _answer(module,'unavailable',note=error or 'Таблица сохранённого результата неполна или имеет неизвестный формат.')
    try:
        report = report_context._module_report(module,'',loaded=(data,error))
    except (ValueError, TypeError, OverflowError, RecursionError):
        return _answer(module,'unavailable',note='Сохранённые показатели имеют некорректный формат.')
    stamp = _date(report.get('collected_at'))
    if report['status']!='saved' or (module=='utnkr' and isinstance(data,dict) and data.get('source_mode')=='local_rows'):
        return _answer(module,'unavailable',as_of=stamp,note='Источник не подтвердил результат проверки. Отсутствие совпадений подтвердить нельзя.')
    if module in ('edo','watercontrol','overdue') and isinstance(data,dict) and data.get('extraction_note') and not raw_rows:
        return _answer(module,'unavailable',as_of=stamp,note='Пустая таблица сопровождается замечанием сбора. Отсутствие записей не подтверждено.')
    period, problem = report_periods.extract(question)
    if problem:
        return _answer(module,'clarify',as_of=stamp,note=problem)
    if period and not report_periods.matches(report.get('period'),period):
        return _answer(module,'unavailable',as_of=stamp,note='У сохранённого результата нет подтверждённого совпадения с запрошенным периодом.')
    rows = report_context._rows(data)
    names = {lookup._plain(raw[key],120) for raw in rows
             for key in ('municipality','city','округ','Муниципалитет','ОМСУ') if isinstance(raw.get(key),str)}
    try:
        selected, rest = lookup._city_filter(question,names)
    except ValueError:
        return _answer(module,'unavailable',as_of=stamp,note='Не удалось проверить справочник муниципалитетов результата.')
    rest = re.sub(r'\b(?:'+_ALIASES[module]+r')\b',' ',rest)
    if period:
        rest = lookup._DATE.sub(' ',rest)
        rest = re.sub(r'\b(?:сегодня|вчера|позавчера|за\s+(?:последн\w*\s+)?(?:\d+\s+)?(?:дней|неделю)|(?:текущ\w*|эт\w*|прошл\w*|предыдущ\w*)\s+(?:месяц\w*|недел\w*))\b',' ',rest)
    num = None
    marked = list(_NUMBER.finditer(rest))
    if len(marked)>1:
        return _answer(module,'clarify',as_of=stamp,note='Укажите один номер для поиска.')
    if marked:
        number = marked[0]
        num = number[1].strip('.,/').casefold()
        rest = rest[:number.start()]+' '+rest[number.end():]
    camera_filter = None
    if module=='cameras':
        for pattern, status in _CAMERA_FILTERS:
            if re.search(r'\b(?:'+pattern+r')\b',rest):
                camera_filter = status
                rest = re.sub(r'\b(?:'+pattern+r')\b',' ',rest)
                break
    rest = re.sub(r'\b(?:провер\w*|задач\w*|строк\w*|запис\w*|карточк\w*|показател\w*|статус\w*|до|результат\w*|организаци\w*)\b',' ',rest)
    words = stock._tokens(lookup._FILLER.sub(' ',rest),query=True)
    if not num and module in ('utnkr','cds','mgkh_rm','watercontrol','cameras') and len(words)==1 and words[0].isdigit():
        num, words = words[0], ()
    items, total, unknown = [], 0, False
    city_keys = {}
    for raw in rows:
        try:
            row = _project(module,raw)
        except (ValueError, TypeError, OverflowError, RecursionError):
            unknown = True
            continue
        if not row:
            unknown = True
            continue
        if selected:
            city = row.get('Муниципалитет')
            if not city:
                unknown = True
                continue
            if city not in city_keys:
                city_keys[city] = lookup.cities.key(city)
            if city_keys[city] not in selected:
                continue
        if num:
            numbers = [row[k].casefold() for k in ('Номер','Номер задачи') if k in row]
            if not numbers:
                unknown = True
                continue
            if num not in numbers:
                continue
        if camera_filter and row.get('Статус камеры') != camera_filter:
            if not row.get('Статус камеры'):
                unknown = True
            continue
        if row.get('Статус камеры') in _CAMERA_LABELS:
            row['Статус камеры'] = _CAMERA_LABELS[row['Статус камеры']]
        if words and not _matches(words, row):
            continue
        total += 1
        if len(items)<limit:
            items.append({'text':' · '.join(label+' '+value for label,value in row.items())})
    note = 'Поиск по операционным полям последней сохранённой таблицы.'
    if report.get('period'):
        note += ' Период таблицы: '+str(report['period']['from'])+' — '+str(report['period']['to'])+'.'
    if unknown:
        note += ' Не у всех строк есть искомые поля.'
        if total == 0:
            return _answer(module,'unavailable',as_of=stamp,note=note+' Отсутствие совпадений не подтверждено.')
    return _answer(module,'ok' if total else 'empty',items=items,total=total,as_of=stamp,note=note)
