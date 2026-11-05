"""Plan read-only operational questions using account grants, geographic taxonomy and dialog scope."""
from __future__ import annotations

import re
from datetime import timedelta

from core.roles import effective_modules
from services import report_municipalities as cities
from services.aichat import report_context, report_periods
from services.water_ai_context import normalized

COMPLAINT_MODULES = ('mingkh', 'collective', 'ecur', 'edds', 'cds')
_COMPLAINT = re.compile(r'\b(?:жалоб\w*|жалова\w*|жалу\w*|обращени\w*|обращен\w*|недоволь\w*)\b')
_CAUSE = re.compile(r'\b(?:почему|причин\w*|повод\w*|проблем\w*|тематик\w*|темам|темы|категори\w*|на что|из-за чего|о чем)\b')
_SITUATION = re.compile(r'\b(?:что (?:там|сейчас|у нас|происходит)|как (?:там|обстоят дела|ситуация)|обстановк\w*|ситуаци\w*|положени\w* дел|чем недовольны)\b')
_QUERY = re.compile(r'\b(?:покажи\w*|расскажи\w*|дай\w*|найди\w*|сколько|каки\w*|проанализируй\w*|разбери\w*|что|сводк\w*|отчет\w*)\b')
_CONTINUATION = re.compile(
    r'^(?:(?:а|и|теперь)\s+){0,2}(?:(?:по|в|для|за|с|на)\s+|(?:сегодня|вчера|подробнее|там|почему так)\b)'
)
_GLOBAL = re.compile(r'\b(?:по (?:всей )?области|вс[яю] область|вся м[.]?о[.]?|в целом|все муниципалитеты)\b')
# Prose tasks and interface directions stay in their existing routes even if they quote a real town.
_OUTSIDE = re.compile(
    r'\b(?:сказк\w*|анекдот\w*|стих\w*|рассказ\w* чехова|творчеств\w*|роман\w*|пьес\w*|'
    r'пример запроса|переведи\w*|перевести|перефразируй\w*|что такое|как (?:подать|написать|составить|обжаловать|пользоваться|работает|открыть|загрузить)|'
    r'(?:напиши\w*|составь\w*|составить|напишите) (?:мне |образец |пример )?(?:жалоб\w*|письмо|текст)|'
    r'помоги (?:написать|составить)|врач\w*|пациент\w*|на боль|на здоровье)\b')
_MUTATION = re.compile(
    r'\b(?:запусти(?:ть|те)?|обнови(?:ть|те)?|перезапусти\w*|удали(?:ть|те)?|'
    r'создай\w*|создать|назначь\w*|назначить|подтверди\w*|согласу[йе]\w*|'
    r'выполни(?:ть|те)?|отправь\w*|отправить|опубликуй\w*|опубликовать|сохрани(?:ть|те)?)\b')
_REGION_OR_PLACE = re.compile(r'\b(?:в|по|для|из)\s+(?:(?:город(?:е|а)?|округ(?:е|у|а)?|омсу)\s+)?([А-ЯЁ][а-яё]+(?:-[А-ЯЁа-яё]+)?)\b')
_TOPIC_HINT = re.compile(r'\b(?:очистн\w*|кос)\b')


def _result(action, modules=(), municipality='', period=None, message='', source=''):
    modules = list(modules)
    scope = {
        'agent_action': action, 'modules': modules,
        'module': modules[0] if len(modules) == 1 else '',
        'municipality': municipality, 'municipalities': [municipality] if municipality else [],
    }
    if period:
        scope['requested_period'] = period
    if source:
        scope['agent_source'] = source
    return {
        'action': action, 'modules': modules, 'municipality': municipality, 'period': period,
        'source': source, 'scope': scope, 'message': message,
    }


def _names(text, modules):
    # The registry is public taxonomy; avoid opening unrelated blocks simply to resolve a city.
    _, titles = cities._maps()
    known = list(titles.values())
    if not known and 'mingkh' in modules:
        # A clean installation can use the dashboard's bundled taxonomy without an imported registry.
        from services.aichat.complaint_causes import _known_cities
        known = list(_known_cities().values())
    if not known:
        known = report_context._known_names(modules)
    matches = report_context._named_municipalities(text, known)
    # Official district adjectives are meaningful when they modify «округ» or «район».
    if not matches and re.search(r'\b(?:округ|район)\w*', text):
        for name in known:
            base = cities.key(name)
            if not base.endswith(('ов', 'ев', 'ин', 'ск')):
                continue
            stem = base if base.endswith('ск') else base + 'ск'
            if re.search(r'(?<!\w)' + re.escape(stem) + r'(?:ий|ого|ому|ом)\s+(?:округ|район)\w*', text):
                matches.append(name)
    return list(dict.fromkeys(cities.display(name) for name in matches))


def _period(text, previous, follow):
    if re.search(r'\bза (?:последний )?месяц\b', text) and not re.search(r'\d+[./-]\d+|недел|вчера|сегодня|завтра', text):
        end = report_periods.today()
        # Same convention as the module filters: an explicitly labelled rolling 30-day window.
        return report_periods.normalize({'from': (end - timedelta(days=29)).isoformat(), 'to': end.isoformat()}), None
    period, error = report_periods.extract(text)
    if not error and not period and follow and previous.get('requested_period') and not re.search(r'без периода|за все время|последни[ейх] данны', text):
        try:
            period = report_periods.normalize(previous['requested_period'])
        except ValueError:
            return None, 'Укажите период ещё раз: сохранённый диапазон дат не распознан.'
    return period, error


def _intent(q, previous, explicit):
    complaint = bool(_COMPLAINT.search(q))
    situation = bool(_SITUATION.search(q))
    previous_action = previous.get('agent_action')
    follow = previous_action in ('complaint_causes', 'platform_report') and bool(_CONTINUATION.search(q))
    reason = bool(_CAUSE.search(q))
    detailed_lookup = _TOPIC_HINT.search(q) or 'collective' in explicit or re.search(r'[№#]\s*\d|\bномер\w*\s+\d', q)
    if complaint and detailed_lookup and not (reason or situation):
        return None
    brief = len(q.split()) <= 10 and re.search(r'\b(?:по|в|за)\b', q)
    if complaint and (reason or _QUERY.search(q) or situation or brief):
        return 'complaint_causes', follow
    if follow and previous_action == 'complaint_causes' and not explicit:
        return 'complaint_causes', follow
    if situation or (follow and not explicit):
        return 'platform_report', follow
    return None


def _requested_modules(q, explicit, causes, selection, previous, follow):
    if causes and _TOPIC_HINT.search(q):
        return ['mingkh']
    if causes and re.search(r'\b(?:карт[ауые] (?:жалоб|добродел)|добродел)\w*', q):
        if not re.search(r'\b(?:еддс|ецур|коллективн\w*)\b', q):
            return ['mingkh']
    if explicit:
        return explicit
    chosen = selection.get('module')
    if chosen and chosen not in ('*', 'all'):
        return [chosen]
    return previous.get('modules', []) if follow else []


def _unknown_place(text, q, follow):
    for match in _REGION_OR_PLACE.finditer(text):
        if normalized(match[0]) in q:
            return match[1]
    if follow:
        possible = re.fullmatch(r'(?:а |теперь |и )?(?:по|в|для)\s+([а-я-]+)[?.!]*', q)
        if possible and not re.fullmatch(r'причин\w*|жалоб\w*|област\w*|округ\w*|всем|регион\w*|целом|данн\w*', possible[1]):
            return possible[1]
    return None


def _source(q, modules, previous, follow):
    if _TOPIC_HINT.search(q):
        return 'treatment'
    if modules == ['mingkh']:
        if re.search(r'\b(?:карт[ауые]|добродел)\b', q):
            return 'map'
        if re.search(r'дашборд\w* (?:по )?жалоб', q):
            return 'dashboard'
    return previous.get('agent_source', '') if follow else ''


def decide(text, user, selection=None, previous=None):
    """Return a read plan, a clarification, an access denial, or None to preserve ordinary chat."""
    if not isinstance(text, str) or not text.strip() or len(text) > 2400:
        return None
    selection, previous = selection or {}, previous or {}
    # Quotes and pasted code are evidence, never requests from the current user.
    q = normalized(report_periods._request_text(text))
    if not q or _OUTSIDE.search(q) or _MUTATION.search(q):
        return None
    explicit, _ = report_context._detected_blocks(q)
    intent = _intent(q, previous, explicit)
    if not intent:
        return None
    action, follow = intent
    causes = action == 'complaint_causes'
    allowed = set(effective_modules(user))
    candidates = COMPLAINT_MODULES if causes else tuple(report_context.MODULES)
    named = _requested_modules(q, explicit, causes, selection, previous, follow)
    if not named and not causes and not re.search(r'\b(?:по|в|для|из|округ|область|омсу)\b', q):
        return None
    if any(module not in candidates for module in named):
        return None
    if any(module not in allowed for module in named):
        return _result('denied', message='У вашей учётной записи нет доступа к запрошенному источнику.')
    if causes and not named and 'mingkh' in allowed:
        named = ['mingkh']
    modules = [module for module in candidates if module in allowed and (not named or module in named)]
    if not modules:
        return _result('denied', message='У вашей учётной записи нет доступного источника для этого запроса.') if causes else None

    region = bool(_GLOBAL.search(q))
    compound = re.search(r'(?<![\w-])([а-яё]+-\d+)(?![\w-])', q)
    if compound:
        return _result('clarify', message='Уточните муниципалитет для населённого пункта «' + compound[1] + '». Свод доступен по ОМСУ.')
    names = _names(q, modules)
    if len(names) > 1:
        return _result('clarify', message='Укажите один муниципалитет для этого анализа: ' + ', '.join(names) + '.')
    city = names[0] if names else ''
    if not city and not region:
        unknown = _unknown_place(text, q, follow)
        if unknown:
            return _result('clarify', message='Не удалось сопоставить муниципалитет «' + unknown + '» со справочником. Уточните его название.')
        if follow and not (_COMPLAINT.search(q) or _SITUATION.search(q)):
            if not any(report_periods.extract(q)) and not re.search(r'подробнее|почему так|последн\w* (?:данн|месяц)|без периода|за все время', q):
                return None
        selected_city = selection.get('municipality')
        prior_city = previous.get('municipality') if follow else ''
        candidate = '' if selected_city in ('*', 'all') else selected_city or prior_city
        if candidate:
            resolved = _names(normalized(candidate), modules)
            if len(resolved) != 1:
                return _result('clarify', message='Не удалось подтвердить выбранный муниципалитет. Уточните название.')
            city = resolved[0]
    if not city and not region:
        platform = explicit or follow or selection.get('module') or re.search(r'\b(?:в системе|в нейроне|на платформе|жители)\b', q)
        if not causes or not (platform or _CAUSE.search(q)):
            return None
    period, error = _period(q, previous, follow)
    if error:
        return _result('clarify', message=error)
    source = _source(q, modules, previous, follow and not explicit) if causes else ''
    return _result(action, modules, city, period, source=source)
