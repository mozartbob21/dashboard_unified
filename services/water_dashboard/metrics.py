"""Explicit summary metrics from recognised DataLens indicators and tables.

The first item is always the source's main metric. A missing main value must
stay missing; an unrelated indicator must never take its place.
"""
import math
import re


def _normal(value):
    return ' '.join(str(value or '').casefold().replace('ё', 'е').split()).rstrip(':')


def _number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = value
    else:
        text = re.sub(r'\s+', '', str(value)).replace(',', '.')
        match = re.fullmatch(r'([+-]?\d+(?:\.\d+)?)([%KКMМB]?)', text)
        if not match:
            return None
        number = float(match.group(1))
        number *= {'K': 1000, 'К': 1000, 'M': 1000000, 'М': 1000000,
                   'B': 1000000000}.get(match.group(2), 1)
    if not math.isfinite(number) or number < 0:
        return None
    return int(number) if number == int(number) else number


def _count(value):
    # A percentage is not a count, even if its displayed value is integral.
    if isinstance(value, str) and '%' in value:
        return None
    number = _number(value)
    return number if isinstance(number, int) else None


def _single_count(value):
    if isinstance(value, str) and len(value.strip().splitlines()) != 1:
        return None
    return _count(value)


def _percent(value, bounded=False):
    number = _number(value)
    return None if bounded and number is not None and number > 100 else number


def _metric(key, label, value, unit):
    return {'id': key, 'label': label, 'value': value, 'unit': unit}


def _widget(data, labels, *, parse=_number, caption=None):
    """Match within indicator containers only; conflicting duplicates are unknown."""
    names = {_normal(label) for label in labels} if not callable(labels) else None
    values = []
    for item in data.get('widgets') or []:
        label = _normal(item.get('label'))
        if not (labels(label) if names is None else label in names):
            continue
        if caption is not None and _normal(item.get('caption')) != _normal(caption):
            continue
        value = parse(item.get('value'))
        if value is not None:
            values.append(value)
    return values[0] if values and all(value == values[0] for value in values) else None


def _has_column(data, names):
    expected = {_normal(name) for name in names}
    municipalities = {'омсу', 'муниципалитет', 'муниципальный округ', 'городской округ'}
    for table in data.get('tables') or []:
        headers = {_normal(header) for header in table.get('headers') or []}
        if headers & municipalities and headers & expected:
            return True
    return False


def _aggregate(data, table, field, columns, *, attendance=False):
    if not _has_column(data, columns):
        return None
    values = []
    for row in table or []:
        raw = row.get(field)
        if raw is None:
            if attendance:
                continue
            # A missing municipality count makes the region total incomplete.
            return None
        value = _percent(raw, bounded=True) if attendance else _count(raw)
        if value is None:
            return None
        values.append(value)
    if not values:
        return None
    return round(sum(values) / len(values), 1) if attendance else sum(values)


def _ratio(numerator, denominator):
    return round(numerator / denominator * 100, 1) if numerator is not None and denominator else None


def primary_available(metrics):
    """Zero is valid, but a secondary metric cannot replace a missing primary."""
    return bool(metrics) and _number(metrics[0].get('value')) is not None


def source_metrics(sid, data, table):
    """Return ordered typed metrics; ``table`` is build_table({sid: data}).

    ``widgets`` must come from actual indicator containers, not page text.
    For NВОС the caller may also pass ``nvos=parse_nvos_kpis(text)``.
    """
    data = data or {}
    if sid == 'valves':
        inserted = _widget(data, lambda name: bool(re.fullmatch(r'внесено всего на \d{4} год', name)), parse=_count)
        plan = _widget(data, ['Должно быть внесено'], parse=_count)
        addressed = _widget(data, ['Внесено с корректным адресом'], parse=_count)
        return [
            _metric('inserted', 'Внесено задвижек', inserted, 'шт.'),
            _metric('plan', 'Должно быть внесено', plan, 'шт.'),
            _metric('correct_address', 'Внесено с корректным адресом', addressed, 'шт.'),
            _metric('completion_pct', 'Внесено к плану', _ratio(inserted, plan), '%'),
        ]
    if sid == 'flush':
        # Accept only this verified indicator if/when the dataset recovers.
        # Chart series, percentages and document/control-row counts are not it.
        completed = _widget(data, ['Кол-во выполненных промывок от общего кол-ва'], parse=_single_count)
        return [_metric('completed', 'Выполнено промывок', completed, 'шт.')]
    if sid == 'edo_rso':
        return [
            _metric('signers_ecp_pct', 'Подписанты с ЭЦП', _widget(data,
                ['Доля (%) должностных лиц, имеющих право подписи и ЭЦП'],
                parse=lambda value: _percent(value, bounded=True)), '%'),
            _metric('signers_total', 'Должностные лица с правом подписи', _widget(data,
                ['Кол-во должностных лиц с правом подписи'], parse=_count), 'чел.'),
            _metric('signers_ecp', 'Подписанты с ЭЦП', _widget(data,
                ['Наличие ЭЦП у подписывающих, кол-во'], parse=_count), 'чел.'),
            _metric('electronic_documents_pct', 'Электронные документы в общем количестве', _widget(data,
                ['Доля (%) ЭД в общем кол-ве документов'],
                parse=lambda value: _percent(value, bounded=True)), '%'),
        ]
    if sid == 'tasks':
        value = _aggregate(data, table, 'tasks', [
            'Кол-во задач', 'Количество задач', 'Просроченные задачи',
            'Количество просроченных задач', 'Кол-во просроченных задач',
        ])
        return [_metric('overdue_tasks', 'Просроченные задачи', value, 'шт.')]
    if sid in ('sys_vs', 'sys_kr'):
        suffix = 'VS' if sid == 'sys_vs' else 'KR'
        return [
            _metric('system_addresses', 'Системные адреса', _aggregate(data, table, 'sys' + suffix,
                ['Системных', 'Системные адреса', 'Количество системных адресов', 'Кол-во системных адресов']), 'адр.'),
            _metric('resonant_addresses', 'Резонансные адреса', _aggregate(data, table, 'res' + suffix,
                ['Резонансных', 'Резонансные адреса', 'Количество резонансных адресов', 'Кол-во резонансных адресов']), 'адр.'),
        ]
    if sid == 'meetings':
        value = _aggregate(data, table, 'att',
            ['Присутствовали на перекличках, %', 'Явка, %', 'Явка (%)'], attendance=True)
        return [
            _metric('attendance_pct', 'Средняя явка по ОМСУ', value, '%'),
            _metric('meetings', 'Проведено совещаний', _widget(data,
                ['Проведено совещаний'], parse=_count), 'шт.'),
        ]
    if sid == 'nvos':
        parsed = data.get('nvos') or {}

        def nvos_value(key, labels, *, parse=_number, caption=None):
            value = _widget(data, labels, parse=parse, caption=caption)
            return value if value is not None else parse(parsed.get(key))

        return [
            _metric('collection_pct', 'Собираемость платы за негативное воздействие', nvos_value('sbor',
                ['Доля (%)'], caption='собираемости платы за негативное воздействие'), '%'),
            _metric('nvv_pct', 'Получено платы от НВВ', nvos_value('nvv_pct', ['% от НВВ']), '%'),
            _metric('paid_accrued', 'Получено / начислено', parsed.get('pay_str') or None, ''),
            _metric('samples_fact_year', 'Отбор проб: факт за год', nvos_value('fact_year',
                ['Факт отборов проб'], caption='(год)', parse=_count), 'шт.'),
            _metric('samples_plan_year', 'Отбор проб: план на год', nvos_value('plan_year',
                ['План отборов проб'], caption='(год)', parse=_count), 'шт.'),
            _metric('samples_fact_week', 'Отбор проб: факт за неделю', nvos_value('fact_week',
                ['Факт отборов проб'], caption='(неделя)', parse=_count), 'шт.'),
            _metric('samples_plan_week', 'Отбор проб: план на неделю', nvos_value('plan_week',
                ['План отборов проб'], caption='(неделя)', parse=_count), 'шт.'),
            _metric('nvv_total', 'Сумма НВВ', nvos_value('sum_nvv', ['Сумма НВВ']), '₽'),
        ]
    return []
