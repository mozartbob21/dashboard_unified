"""Readable, deterministic analysis of saved observations. Opens no source and runs no model."""
from datetime import datetime
from core.roles import MODULE_NAMES, MODULE_ALIASES
from services import check_history, check_metrics

CHECK_NAMES = {**MODULE_NAMES, 'mingkh-water-map':'Карта жалоб МинЖКХ',
               'camera_prescriptions':'Предписания по камерам'}
STATUS = {'error':'Проверка завершилась ошибкой.', 'interrupted':'Проверка прервана; результат не подтверждён.',
          'skipped':'Проверка пропущена.', 'partial':'Проверка завершилась частично.',
          'running':'Проверка ещё выполняется.', 'unconfirmed':'Завершение проверки не подтверждено.'}

def stamp(value):
    return datetime.fromisoformat(value).astimezone(check_history.MOSCOW).strftime('%d.%m.%Y %H:%M')

def number(value):
    return format(value, ',.3f').rstrip('0').rstrip('.').replace(',', ' ').replace('.', ',')

def describe(result, *, compact=False):
    latest, older = result['latest'], result['baseline']
    module = result['module_id']
    title = CHECK_NAMES.get(module, MODULE_NAMES.get(MODULE_ALIASES.get(module), module))
    if latest and latest['params'].get('operation') == 'treatment':
        title = 'Жалобы по очистным сооружениям'
    period = ('За последний месяц: '+stamp(result['start'])+' — '+stamp(result['end'])+' МСК'
              if result['window']=='month' else 'Две последние проверки')
    dates = ('Сравнение: '+stamp(older['at'])+' → '+stamp(latest['at'])+' МСК.') if older and latest else ''
    metrics = []
    if latest and latest['result_code'] in ('measured','partial'):
        for key,value in latest['metrics'].items():
            label,unit,_ = check_metrics.metric_info(key)
            metrics.append({'label':label,'unit':unit,'value':value})
    status = result['status']
    if not latest:
        status, summary = 'empty', 'За этот период сохранённых проверок нет.'
    elif latest['status'] in STATUS:
        summary = STATUS[latest['status']]+' Динамику по этой проверке определить нельзя.'
        if latest.get('long_running'):
            summary += ' Выполнение длится более четырёх часов; завершение ещё не подтверждено.'
    elif latest['result_code'] in ('legacy','unconfirmed'):
        summary = 'Запуск сохранён без подтверждённых показателей. Динамику определить нельзя.'
    elif status != 'compared':
        summary = result['reason']
    else:
        changes = sorted(result['deltas'], key=lambda x: (x['direction']=='same', x['assessment']=='neutral', x['key']))
        lines = []
        for item in changes[:2 if compact else 20]:
            delta=item['delta']
            change = ('без изменения' if not delta else ('+' if delta > 0 else '−')+number(abs(delta))+
                      (' п. п.' if item['unit']=='%' else ' '+item['unit']))
            evaluation = {'improved':', улучшение','worsened':', ухудшение'}.get(item['assessment'],'')
            lines.append(f"{item['label']}: {number(item['previous'])} → {number(item['current'])} {item['unit']} ({change}{evaluation}).")
        summary = ' '.join(lines)
        if not compact and len(changes)>20:
            summary += ' Показаны первые 20 сопоставимых показателей из '+str(len(changes))+'.'
    return {'title':title,'period_label':period+(' · '+dates if dates else ''),'summary':summary,
            'latest_at':latest['at'] if latest else None,'status':status,'metrics':metrics}

def answer(result):
    data=describe(result)
    lines=[data['title']+'. '+data['period_label'].rstrip('.')+'.']
    if data['latest_at'] and not result['baseline']:
        lines.append('Последняя проверка: '+stamp(data['latest_at'])+' МСК.')
    lines.append(data['summary'])
    if result['status']!='compared' and data['metrics']:
        lines.append('Последние зафиксированные значения: '+'; '.join(
            f"{m['label']}: {number(m['value'])} {m['unit']}" for m in data['metrics'][:8])+'.')
    return '\n'.join(lines)
