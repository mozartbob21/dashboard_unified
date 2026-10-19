"""Explicit report dates; collection timestamps never stand in for data periods."""
from datetime import date, datetime, timedelta, timezone
import re

MOSCOW = timezone(timedelta(hours=3))
_DATE = r'(?:\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})'


def today():
    return datetime.now(MOSCOW).date()


def _date(value):
    if not isinstance(value, str):
        raise ValueError('Дата должна быть текстом.')
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', value):
        return date.fromisoformat(value)
    if re.fullmatch(r'\d{2}\.\d{2}\.\d{4}', value):
        return datetime.strptime(value, '%d.%m.%Y').date()
    raise ValueError('Укажите даты в формате ГГГГ-ММ-ДД или ДД.ММ.ГГГГ.')


def normalize(value):
    """Whitelist safe persisted metadata; reject malformed/future/inverted ranges."""
    if not isinstance(value, dict):
        raise ValueError('Не удалось определить период отчёта.')
    first, last = _date(value.get('from')), _date(value.get('to'))
    if first > last or last > today():
        raise ValueError('Начало периода должно быть не позже окончания; будущие даты недоступны.')
    label = first.strftime('%d.%m.%Y') if first == last else f'{first:%d.%m.%Y}–{last:%d.%m.%Y}'
    return {'from': first.isoformat(), 'to': last.isoformat(), 'label': label, 'timezone': 'Europe/Moscow'}


def _request_text(question):
    """Quoted source text is not a date constraint on the user's data request."""
    q = str(question or '').casefold().replace('ё', 'е')
    q = re.sub(r'```.*?(?:```|$)', ' ', q, flags=re.S)
    q = re.sub(r'(?m)^\s*>.*$', ' ', q)
    def quoted_period(match):
        quoted = match[0]
        value = quoted[1:-1].strip()
        # A bare date explicitly introduced as a period remains a constraint.
        # General source quotations and all inline code remain opaque.
        date_range = _DATE + r'(?:\s*(?:по|до|[—–-])\s*' + _DATE + r')?'
        if quoted[0] != '`' and re.search(r'\b(?:за|на|с|по|до)\s*$', q[:match.start()]) and re.fullmatch(
                r'(?:' + date_range + r'|сегодня|вчера|позавчера|завтра|послезавтра)', value):
            return ' ' + value + ' '
        return ' '
    q = re.sub(r'`[^`]*`|«[^»]*»|“[^”]*”|"[^"\n]*"', quoted_period, q)
    return q


def _common_period(q):
    """Unambiguous rolling/calendar periods, in Moscow and inclusive of dates."""
    end = today()
    week = end - timedelta(days=end.weekday())
    month = end.replace(day=1)
    candidates = []
    for match in re.finditer(r'\b(?:за\s+(?:последни[ех]\s+)?|последни[ех]\s+)(\d+)\s*(?:день|дня|дней)\b', q):
        days = int(match[1])
        if not 1 <= days <= 3660:
            raise ValueError('Недопустимая длительность периода.')
        candidates.append((end - timedelta(days=days-1), end))
    patterns = (
        (r'\bза\s+(?:последнюю\s+)?неделю\b', (end-timedelta(days=6), end)),
        (r'\b(?:текущ\w*|эт\w*)\s+недел\w*\b', (week, end)),
        (r'\b(?:прошл\w*|предыдущ\w*)\s+недел\w*\b', (week-timedelta(days=7), week-timedelta(days=1))),
        (r'\b(?:текущ\w*|эт\w*)\s+месяц\w*\b', (month, end)),
        (r'\b(?:прошл\w*|предыдущ\w*)\s+месяц\w*\b', ((month-timedelta(days=1)).replace(day=1), month-timedelta(days=1))),
    )
    candidates.extend(bounds for pattern, bounds in patterns if re.search(pattern, q))
    distinct = set(candidates)
    if len(distinct) > 1:
        raise ValueError('Несколько периодов для сравнения.')
    if distinct:
        start, finish = distinct.pop()
        return normalize({'from':start.isoformat(), 'to':finish.isoformat()})
    return None


def extract(question):
    """Return (period|None,error|None), without guessing ambiguous month names."""
    q = _request_text(question)
    if re.search(r'\b(?:завтра|послезавтра)\b', q):
        return None, 'Будущий период ещё не завершён. Укажите дату не позже сегодняшней; сохранённые итоги за другой период не подставлены.'
    values = re.findall(r'(?<![\w-])' + _DATE + r'(?![\w-])', q)
    relative = re.findall(r'\b(сегодня|вчера|позавчера)\b', q)
    if values and relative or len(values) > 2 or len(set(relative)) > 1:
        return None, 'Укажите один период отчёта: дату или диапазон «с ДД.ММ.ГГГГ по ДД.ММ.ГГГГ».'
    try:
        common = _common_period(q)
        if common and (values or (relative and not (set(relative) == {'сегодня'} and re.search(r'включая\s+сегодня', q)))):
            return None, 'Укажите один период отчёта; для сравнения нужны отдельные срезы.'
        if common:
            return common, None
        if values:
            if len(values) == 2:
                between = re.search('(' + _DATE + r')\s*(.*?)\s*(' + _DATE + ')', q)
                if not between or not re.fullmatch(r'(?:по|до|[—–-])', between[2]):
                    return None, 'Для сравнения отдельных дат нужны отдельные срезы. Укажите один период «с ДД.ММ.ГГГГ по ДД.ММ.ГГГГ».'
            return normalize({'from': values[0], 'to': values[-1]}), None
        if relative:
            day = today() - timedelta(days={'сегодня':0, 'вчера':1, 'позавчера':2}[relative[0]])
            return normalize({'from':day.isoformat(), 'to':day.isoformat()}), None
        if re.search(r'\b(?:за|на|с)\s+(?:(?:прошл\w*|предыдущ\w*|текущ\w*|эт\w*|последн\w*)\s+)?(?:\d+\s+)?(?:недел\w*|месяц\w*|год\w*|дн\w*|январ\w*|феврал\w*|март\w*|апрел\w*|ма[йяе]|июн\w*|июл\w*|август\w*|сентябр\w*|октябр\w*|ноябр\w*|декабр\w*)\b', q):
            return None, 'Укажите точные даты периода — например, «с 01.10.2026 по 15.10.2026». Итоги за другой период не будут подставлены.'
    except ValueError:
        return None, 'Укажите один корректный период: начало не позже окончания, без будущих дат.'
    return None, None


def matches(available, requested):
    if not isinstance(available, dict):
        return False
    try:
        return (_date(available.get('from')).isoformat(), _date(available.get('to')).isoformat()) == (requested['from'], requested['to'])
    except (ValueError, TypeError):
        return False


def restrict(report, requested):
    """Drop all facts with an incompatible period, preserving source identity/date."""
    if not requested:
        return report
    if report.get('period_selection') == 'exact':
        return report
    if report.get('id') == 'mingkh-appeals':
        metrics = list(report.get('metrics') or [])
        periods = report.get('periods') or {}
        for index, key in enumerate(('curr', 'prev', 'appg')):
            bounds = periods.get(key)
            if isinstance(bounds, (list, tuple)) and len(bounds) == 2 and matches(dict(zip(('from', 'to'), bounds)), requested):
                # Existing projection emits metrics only for present period entries.
                offset = sum(k in periods for k in ('curr', 'prev', 'appg')[:index])
                if offset < len(metrics):
                    report['metrics'] = [dict(metrics[offset], label='Обращения за запрошенный период')]
                    report['periods'] = {key: bounds}
                    report['period'] = {'from': requested['from'], 'to': requested['to']}
                    report['period_selection'] = 'exact'
                    return report
    if matches(report.get('period'), requested) and report.get('period_selection') != 'unavailable':
        report['period_selection'] = 'exact'
        return report
    for key in ('metrics', 'rows', 'entities'):
        report[key] = []
    for key in ('tables', 'details', 'status_counts', 'category_counts', 'municipality_table', 'matching_rows', 'omitted_rows', 'signatures_coverage'):
        report.pop(key, None)
    report['period_selection'] = 'unavailable'
    report['requested_period'] = dict(requested)
    report['status'] = 'missing'
    report['warning'] = ('Нет подтверждённой выборки за запрошенный период ' + requested['label']
                         + '. Дата сохранения не означает, что данные относятся к этому периоду; показатели за другой период не подставлены.')
    return report
