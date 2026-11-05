"""Dispatch read-only questions to registered projections, never to a file path, SQL or a collector."""
import re
from core.roles import MODULE_NAMES, MODULE_ALIASES, check_module_access
from services.aichat.report_context import _detected_blocks
from services.aichat import catalog_lookup, operational_lookup, treatment_lookup

def _notice(text):
    return {'answer': text, 'kind':'clarification', 'scope':{}}

def resolve(text,user,selection=None,previous=None):
    selection,previous=selection or {},previous or {}
    q=' '.join(text.casefold().replace('ё','е').split())
    if not q or re.search(r'как (?:мне |можно )?(?:найти|добавить|загрузить|открыть)|что такое|пример запроса',q):
        return None
    modules,sources=_detected_blocks(q)
    treatment=bool(re.search(r'очистн|\bкос\b',q))
    if treatment:
        modules=list(dict.fromkeys(['mingkh',*[m for m in modules if m!='mingkh']]))
    for name in MODULE_NAMES:
        if re.search(r'(?<![а-яa-z0-9_-])'+re.escape(name)+r'(?![а-яa-z0-9_-])',q) and name not in modules:
            modules.append(name)
    find=bool(re.search(r'\b(?:найди|найдите|поищи|отыщи|проверь наличие|есть ли|укажи номер|найдется)\b|(?:поиск|найти) (?:в |по )?(?:блок|систем|нейрон)',q))
    detail=bool(re.search(r'\b(?:каки[ех]|сколько|покажи|выведи|перечисли|что (?:есть|имеется|на складе)|где (?:есть|найти)|номер|статус|остаток|наличие)\b|[№#]\s*\d',q))
    if not find and '№' not in q and re.search(r'отчет|сводк|статистик|объясни|возможност|кнопк|как пользоваться|как работает',q):
        return None
    if not (find or detail or treatment):
        return None
    if re.search(r'\b(?:отчет|сводк|презентацию|сравни|динамик|инструкци)',q) and not find and not treatment:
        return None   # keep the existing report, document and history capabilities
    supported={*catalog_lookup.SUPPORTED,'zips','collective'}
    module=modules[0] if len(modules)==1 else None
    if len(modules)>1:
        return _notice('Уточните один блок для поиска записи.') if find or treatment else None
    if not module:
        selected=selection.get('module')
        module=selected if selected in MODULE_NAMES else previous.get('lookup_module')
        if not module and find and re.search(r'\b(?:муфт|рубероид|труб[ауые]|материал|остатки)\w*',q):
            module='zips'
        if not module and find and re.search(r'в (?:нейроне|системе|блоке|платформе)|по (?:системе|блокам)',q):
            return _notice('В каком доступном блоке искать? Укажите название блока и номер, материал или ОМСУ.')
        if not module:
            return None
    treatment=treatment or bool(not modules and previous.get('lookup_treatment') and module=='mingkh')
    if module not in supported and not treatment:
        if module in ('edds','mingkh','water-dashboard','water_rm','ecur'):
            return None # dedicated operational reports, filtered by municipality/period on the server
        if find and module in MODULE_NAMES:
            return _notice('Для этого блока не подключён поиск по отдельным записям в общем чате.')
        return None
    if not check_module_access(user,module):
        return _notice('Нет доступа к данным указанного блока.')
    if len(text)>1200:
        return _notice('Для точного поиска сократите вопрос до 1 200 символов и укажите номер, материал или муниципалитет.')
    city=selection.get('municipality')
    if city and city not in ('all','*') and city.casefold() not in q:
        text+=' '+city
    if len(text)>1200:
        return _notice('Уточните один короткий запрос и муниципалитет.')
    item=(treatment_lookup.search(user,text) if treatment else
          operational_lookup.search(user,module,text) if module in ('zips','collective') else catalog_lookup.search(user,module,text))
    # All URLs come from registered adapters. Rows are quoted as display data; no model interprets them as instructions.
    date=item.get('as_of')
    lines=['Источник: '+item['title']+' ('+item['url']+'). '+('Дата источника: '+str(date)+'.' if date else 'Дата источника не указана.')]
    total=item.get('total')
    if total is not None:
        lines.append('Найдено записей: '+str(total)+'.')
    for row in item['items']:
        if row.get('text'):
            lines.append('• '+row['text'])
    if total is not None and total>len(item['items']):
        lines.append('Показано: '+str(len(item['items']))+' из '+str(total)+'.')
    if item['note']:
        lines.append(item['note'])
    return {'answer':'\n'.join(lines),'kind':'answer' if total is not None else 'clarification','scope':{
        'module':module,'modules':[module],'lookup_module':module,'lookup_treatment':treatment}}
