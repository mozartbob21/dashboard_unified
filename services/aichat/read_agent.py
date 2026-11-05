"""Execute server-authorized, read-only chat tools against actual saved block data."""
import re

from fastapi import HTTPException

from core.roles import MODULE_NAMES, check_module_access, effective_modules
from services import check_analysis, check_history
from services.aichat import (
    agent_intent, catalog_lookup, check_dynamics, complaint_causes, data_lookup,
    operational_lookup, other_complaint_causes, platform_guide, read_planner,
    report_context, treatment_lookup,
)
from services.aichat.skills.preferences import enabled_skills

MAX_READS = 2


def _reply(answer, scope=None, sources=(), kind='answer'):
    return {'answer': answer, 'kind': kind, 'scope': scope or {}, 'sources': list(sources)}


def _source(module, title, url, stamp, status):
    return {'module': module, 'id': module, 'title': title, 'source_url': url,
            'collected_at': stamp, 'status': status}


def _date_label(value):
    if value and re.fullmatch(r'\d{4}-\d{2}-\d{2}', value):
        year, month, day = value.split('-')
        return f'{day}.{month}.{year}'
    return value or 'не указана'


def _causes_text(result):
    period = result.get('period')
    place = result.get('municipality') or 'Московская область'
    total = result['total']
    lines = [place + (' · ' + _date_label(period['from']) + ' — ' + _date_label(period['to']) if period else '')]
    if total is not None:
        lines.append(f'Количество обращений: {total}.')
    groups = result['groups']
    if groups:
        first = groups[0]['count']
        leaders = [g['label'] for g in groups if g['count'] == first]
        share = round(100 * first / total) if total else 0
        if len(leaders) == 1:
            lines.append(f'Самая частая тема — «{leaders[0]}»: {first} из {total} ({share} %).')
        else:
            lines.append('Чаще всего жители обращались по темам ' + ', '.join('«'+label+'»' for label in leaders[:4])
                         + f'. На каждую приходится {first} из {total} обращений ({share} %).')
        lines.extend(f"• {group['label']} — {group['count']}" for group in groups)
    if result.get('other'):
        lines.append(f"Остальные группы — {result['other']}.")
    if result.get('unknown'):
        lines.append(f"Без распознанной причины — {result['unknown']}.")
    if result['note']:
        lines.append(result['note'])
    lines.append(f"Источник: {result['title']} ({result['url']}). Дата источника: {_date_label(result['as_of'])}.")
    return '\n'.join(lines)


def _causes(user, plan, read, original_text):
    modules = list(plan['modules'])
    # Prefer the ministry's general register. An unscoped request can use one
    # different authorized complaints register if that source is unavailable.
    named, _ = report_context._detected_blocks(original_text.casefold())
    if modules == ['mingkh'] and not named and not plan['source'] and not plan.get('pinned'):
        modules += [module for module in ('collective','ecur','edds') if check_module_access(user,module)]
    texts, sources = [], []
    for module in modules[:MAX_READS]:
        def fetch():
            if module == 'mingkh' and plan['source'] != 'treatment':
                return complaint_causes.analyze(user, plan['municipality'], plan['period'], source=plan['source'] or 'auto')
            key = 'treatment' if module == 'mingkh' else module
            return other_complaint_causes.analyze(user,key,plan['municipality'],plan['period'])
        result = read(module, fetch)
        texts.append(_causes_text(result))
        sources.append(_source(module,result['title'],result['url'],result['as_of'],result['status']))
        if modules != plan['modules'] and result['status'] in ('ok','limited'):
            break
    ids = [s['module'] for s in sources]
    scope = {**plan['scope'],'modules':ids,'module':ids[0] if len(ids)==1 else '', 'agent_action':'complaint_causes'}
    if len(sources)>1:
        texts.append('Каждый источник приведён отдельно. Периоды и наборы обращений могут различаться.')
    return _reply('\n\n'.join(texts),scope,sources)


def _search(user, module, text):
    if module == 'mingkh':
        return treatment_lookup.search(user,text)
    if module in ('zips','collective'):
        return operational_lookup.search(user,module,text)
    return catalog_lookup.search(user,module,text)


def _search_text(result):
    lines = [result['title'] + ' (' + result['url'] + '). Дата источника: ' + _date_label(result.get('as_of')) + '.']
    if result['total'] is not None:
        lines.append(f"Найдено записей: {result['total']}.")
    lines.extend('• '+item['text'] for item in result['items'])
    if result.get('note'):
        lines.append(result['note'])
    return '\n'.join(lines)


def _summary(bundle):
    facts = []
    for source in bundle['sources']:
        metrics = list(source.get('metrics') or [])
        for row in source.get('rows') or []:
            if isinstance(row,dict):
                metrics.extend(row.get('metrics') or [])
        numbers = [f"{item['label']} — {item['value']} {item.get('unit','')}".strip()
                   for item in metrics if item.get('value') is not None]
        note = source.get('warning') or ''
        if source.get('status') == 'stale':
            note = 'Показан ранее сохранённый результат. ' + note
        facts.append(f"{source['title']} ({source.get('source_url','')}) · дата: "
                     f"{_date_label(source.get('data_date') or source.get('collected_at'))}. "
                     + ('; '.join(numbers[:5])+'. ' if numbers else ' ') + note)
    return '\n'.join(dict.fromkeys(facts))[:6500] or 'В выбранных источниках нет подтверждённых показателей за этот период.'


def _report(user, text, plan, read, narrate):
    allowed = plan['modules'][:MAX_READS]
    combined = {'schema_version':2,'selection':plan['scope'],'sources':[],'limitations':[]}
    for module in allowed:
        # prepare() discovers municipal names only inside the assigned scope.
        bundle = read(module, lambda: report_context.prepare(text,[module],{
            'module':module,'municipality':plan['municipality']},period_constraint=plan['period']))
        if bundle.get('clarification'):
            combined['limitations'].append(bundle['clarification'])
        combined['sources'].extend(s for s in bundle['sources'] if s.get('module') == module)
        combined['limitations'].extend(bundle['limitations'])
    summary = _summary(combined)
    if not combined['sources']:
        reason = next(iter(combined['limitations']), summary)
        return _reply(reason,plan['scope'],kind='clarification')
    useful = any(s.get('status') in ('saved','current','stale') for s in combined['sources'])
    narrative = ''
    if useful and narrate:
        try:
            narrative = narrate(report_context.to_prompt(combined), allowed)
        except (HTTPException, PermissionError):
            raise
        except Exception:
            narrative = ''
    # Even if model transport fails, the actual source values and dates stay visible.
    lines = [narrative.strip(),summary] if isinstance(narrative,str) and narrative.strip() else [summary]
    sources = [_source(s['module'],s['title'],s.get('source_url'),s.get('collected_at'),s.get('status')) for s in combined['sources']]
    return _reply('\n\n'.join(lines),plan['scope'],sources)


def resolve(text,user,selection,previous,*,automatic_skills=False,explicit='',refresh=None,emit=None,narrate=None):
    """No files, writes, raw SQL, credentials or arbitrary URLs are available here."""
    if explicit == 'false':
        return None
    emit = emit or (lambda event: None)
    def step(ident, label, state):
        emit({'type':'step','id':ident,'kind':'tool','label':label,'status':state})
    def read(module, operation):
        current = refresh() if refresh else user
        if not current or not check_module_access(current,module):
            raise HTTPException(403,'Права на источник изменились во время анализа.')
        ident = 'agent-read:'+module
        step(ident,'Читаю сохранённые данные: '+MODULE_NAMES[module],'running')
        try:
            result = operation()
        except Exception:
            step(ident,'Не удалось прочитать источник: '+MODULE_NAMES[module],'error')
            raise
        failed = result.get('status') in ('unavailable','missing') if isinstance(result,dict) else False
        step(ident,('Источник недоступен: ' if failed else 'Проверен источник: ')+MODULE_NAMES[module], 'error' if failed else 'done')
        return result

    # Preserve the direct history and exact-record tools. Identify complaint
    # analysis before a generic record search can turn it into a list of examples.
    decision = agent_intent.decide(text,user,selection,previous)
    previous_task = previous.get('agent_action')
    def permits_automatic_read():
        if not automatic_skills or explicit=='true':
            return True
        return any(skill['id']=='platform-report' for skill in enabled_skills(user))
    if decision and decision['action'] in ('clarify','denied'):
        return _reply(decision['message'],kind='clarification')
    if decision and decision['action']=='complaint_causes':
        if not permits_automatic_read():
            return _reply('Автоматическое чтение данных отключено. Включите «Отчёт по данным платформы» в навыках чата.',kind='clarification')
        decision['pinned'] = bool(selection.get('module') or previous_task)
        return _causes(user,decision,read,text)
    direct = check_dynamics.resolve(text,user,selection,previous)
    if direct is None:
        direct = data_lookup.resolve(text,user,selection,previous)
    if direct is not None:
        return direct
    if platform_guide.help_only(text) or explicit=='true':
        return None
    # Preserve existing report contracts (including all blocks and multiple cities).
    # The planner handles informal questions which those report routes don't take.
    if not decision:
        if not read_planner.would_read(text,selection,previous):
            return None
        if report_context.wants_context(text,effective_modules(user),previous):
            return None
    if not permits_automatic_read():
        return _reply('Автоматическое чтение данных отключено в навыках чата.',kind='clarification')
    if decision and len(decision['modules'])<=MAX_READS:
        plan = {**decision,'tool':'report'}
    else:
        step('agent-route','Выбираю доступный источник','running')
        plan = read_planner.plan(text,user,selection,previous)
        step('agent-route','Выбор источника завершён','done')
    if plan is None:
        return None
    tool = plan['tool']
    if tool=='clarify':
        return _reply(plan['message'],kind='clarification')
    if tool=='complaint_causes':
        return _causes(user,plan,read,text)
    if tool=='report':
        return _report(user,text,plan,read,narrate)
    if tool=='history':
        module = plan['modules'][0]
        window = 'month' if 'месяц' in text.casefold() else 'last_two'
        result = read(module,lambda:check_history.compare(user,module,window=window,public_only=True))
        return _reply(check_analysis.answer(result),plan['scope'])
    if tool=='search':
        results=[read(m,lambda:_search(user,m,text)) for m in plan['modules']]
        sources=[_source(r['scope'],r['title'],r['url'],r.get('as_of'),r.get('status')) for r in results]
        return _reply('\n\n'.join(map(_search_text,results)),plan['scope'],sources)
    return None
