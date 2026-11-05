"""Only the current user's question requests history; result text is never executable."""
import re
from core.roles import MODULE_ALIASES, MODULE_NAMES, check_module_access
from services import check_analysis, check_history, check_metrics
from services.aichat.report_context import _detected_blocks

_SIGNAL = re.compile(r'динамик|сравни|сравнен|сопостав|изменил|изменен|стало (?:лучше|хуже)|улучшил|ухудшил')
_HISTORY = re.compile(r'проверк|запуск|прогон|блок|ситуац|динамик|месяц|показател')
_EXACT = re.compile(r'(?<![a-z0-9_-])('+'|'.join(re.escape(x) for x in sorted({*MODULE_NAMES,*MODULE_ALIASES},key=len,reverse=True))+r')(?![a-z0-9_-])')
_EXTRA = ((r'очистн|\bкос\b','mingkh','treatment'),(r'карта жалоб|карт[ауые] (?:жалоб )?минжкх','mingkh-water-map',''),
          (r'предписани\w* по камер','camera_prescriptions',''),(r'рейтинг\w* омсу','omsu-rating',''),
          (r'презентаци','presentations',''),(r'\bсмс\b','mingkh','sms'))
_OPERATION_HINTS = {'zips_scan':'сканирование остатков', 'treatment':'жалобы по очистным сооружениям','sms':'СМС'}
def _clarify(text):
    return {'answer':text,'scope':{},'kind':'clarification'}

def resolve(text,user,selection=None,previous=None):
    q=' '.join(text.lower().replace('ё','е').split())
    selection,previous=selection or {},previous or {}
    last=bool(re.search(r'последн\w* (?:дв[аеу]\w* |2 )?(?:проверк|запуск)',q))
    followup=bool(previous.get('history_module') and re.fullmatch(r'(?:а |тогда )?(?:за )?(?:последн\w* )?месяц[?.!]*',q))
    if not (last or followup or (_SIGNAL.search(q) and (_HISTORY.search(q) or _EXACT.search(q)))):
        return None
    if any(w in q for w in ('проверку орфографии','проверки орфографии','проверки текста')):
        return None
    exact=list(dict.fromkeys(_EXACT.findall(q)))
    found,sources=_detected_blocks(q)
    special=[(m,op) for pat,m,op in _EXTRA if re.search(pat,q)]
    if len(special)>1:
        return _clarify('Уточните один раздел для сравнения проверок.')
    prev=previous.get('history_module') or previous.get('module') or ''
    if not prev and len(previous.get('modules') or [])==1:
        prev=previous['modules'][0]
    selected=selection.get('module')
    op=''
    if special:
        module,op=special[0]
        # A parent title in the same question describes a specific submodule.
        found=[m for m in found if m not in (module,MODULE_ALIASES.get(module),'mingkh-sms','mingkh' if op=='treatment' else None)]
        candidates={module,*found,*[m for m in exact if MODULE_ALIASES.get(module)!=m]}
    else:
        candidates=set(exact+found)
    if len(candidates)>1:
        return _clarify('Уточните один блок: история и показатели разных блоков не складываются.')
    module=next(iter(candidates),None) or (selected if selected not in (None,'','*','all') else prev)
    if not module or module not in {*MODULE_NAMES,*MODULE_ALIASES}:
        return _clarify('Для какого блока сравнить проверки? Например: «Сравни последние проверки ЗиП».')
    canonical=MODULE_ALIASES.get(module,module)
    if not check_module_access(user,canonical):
        return _clarify('Нет доступа к истории указанного блока.')
    if re.search(r'сканирован\w* остатков',q) and canonical=='zips':
        op='zips_scan'
    if not candidates and module==prev:
        op=op or previous.get('history_operation') or ''
    source=selection.get('source') or (sources[0] if len(sources)==1 else '')
    if not candidates and module==prev and not source:
        source=previous.get('history_source') or ''
    if len(sources)>1 or source and (canonical!='water-dashboard' or source not in check_metrics.SOURCE_IDS):
        return _clarify('Выберите один источник сводного дашборда для сравнения его проверок.')
    if (selection.get('municipality') not in (None,'','*','all') or re.search(r'муниципалитет|\bокруг|\bвласих|\b(?:по|для|в)\s+(?!блок|эт|последн|предыдущ|целом|общем|систем|жалоб|камер|очистн|качеств|водоснабж|рсо|омсу|проверк|источник|сравнен|данным|зип|еддс|минжкх|цдс|у?тнкр|ецур|сводн|контрол|коллективн|остатк|просроч|заполн|предписан|технадзор|рейтинг)[а-я]{3,}',q)):
        return _clarify('История содержит итоги блока без разбивки по муниципалитетам. Уточните, нужно ли сравнить блок в целом.')
    if re.search(r'\d+[./-]\d+|\b(?:недел|год|квартал|прошл\w* месяц|предыдущ\w* месяц)|\b(?:два|три|\d+)\s+месяц|\d+\s+дн|пяти? провер|три провер',q):
        return _clarify('Доступно сравнение двух последних проверок или первой и последней за скользящий месяц. Какой период взять?')
    window='month' if 'месяц' in q else 'last_two'
    params={k:v for k,v in [('operation',op),('source',source)] if v}
    result=check_history.compare(user,module,window=window,params=params,public_only=True)
    return {'answer':check_analysis.answer(result),'scope':{'module':canonical,'modules':[canonical],'history_module':module,
            'history_operation':op,'history_source':source},'kind':'answer'}
