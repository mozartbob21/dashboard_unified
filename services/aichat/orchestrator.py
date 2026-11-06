"""Execute small validated plans in the existing durable chat worker.

No replay, arbitrary code or autonomous tool calls from a result. Numerical
answers are computed on the server; Qwen is not asked to invent missing facts.
"""
from __future__ import annotations

from datetime import date
from itertools import islice
import math
import time

from fastapi import HTTPException

from . import action_planner, jobs, private_storage, tool_catalog
from .live_mingkh_tool import TOPICS as MINGKH_TOPICS


def describe(plan):
    lines = []
    for index, step in enumerate(plan['steps'], 1):
        spec = tool_catalog.spec(step['tool'])
        params = step['params']
        details = []
        if params.get('start'):
            details.append(params['start'] + ' — ' + params['end'])
        if spec['operation'] == 'transform':
            details.append('текст сообщения пользователя')
        if spec['operation'] in ('analytics', 'query', 'check'):
            details += [params.get('municipality') or 'вся область',
                        ({'all': 'ХВС, ГВС и водоотведение', 'hvs': 'ХВС', 'gvs': 'ГВС', 'vo': 'водоотведение'}[params.get('service', 'all')]
                         if spec['operation'] == 'analytics' else MINGKH_TOPICS.get(params.get('topic'), ('текущее состояние', {}))[0])]
        lines.append(str(index) + '. ' + spec['title'] + (' — ' + ', '.join(details) if details else ' — сохранённые настройки блока'))
    return '\n'.join(lines)


def proposal(plan, requirements):
    return {'ok': True, 'response_kind': 'clarification', 'required_modules': requirements,
            'answer': 'План подготовлен. Проверьте действия перед запуском. Обновление сохранит общий результат блоков. Фильтр ОМСУ применяется к анализу этого запуска.\n\n' + describe(plan),
            'action': {'kind': 'agent_plan', 'plan': plan, 'label': 'Подтверждение плана',
                       'summary': describe(plan)}}


def _number(value):
    return type(value) in (int, float) and math.isfinite(value)


def comparison(left, right):
    """Comparable definitions/units/filters/snapshot and complete requested windows only."""
    reasons = []
    if any(item.get('status') not in ('ok', 'success') or item.get('coverage') != 'full' for item in (left, right)):
        reasons.append('Один из периодов неполный либо данные отсутствуют. Разница за полные периоды не рассчитана.')
    if left.get('filters') != right.get('filters'):
        reasons.append('Фильтры выборок различаются.')
    if left.get('source', {}).get('id') != right.get('source', {}).get('id'):
        reasons.append('Источники имеют разные определения показателей.')
    independent_live = all(item.get('comparison_consistency') == 'independent_live_reads' for item in (left,right))
    if not independent_live and (not left.get('snapshot_id') or left.get('snapshot_id') != right.get('snapshot_id')):
        reasons.append('Снимок источника изменился между расчётами. Повторите сравнение по одному снимку.')
    if reasons:
        return {'status': 'incomplete', 'reasons': reasons, 'metrics': []}
    a = {item['id']: item for item in left.get('metrics', [])}
    b = {item['id']: item for item in right.get('metrics', [])}
    output = []
    if set(a) != set(b) or not a:
        return {'status': 'incomplete', 'reasons': ['Набор показателей различается или отсутствует.'], 'metrics': []}
    for key in a:
        first, second = a[key], b[key]
        if not first.get('definition') or first.get('definition') != second.get('definition') or first.get('unit') != second.get('unit'):
            return {'status': 'incomplete', 'reasons': ['Определения или единицы показателей различаются.'], 'metrics': []}
        x, y = first.get('value'), second.get('value')
        if not _number(x) or not _number(y):
            return {'status': 'incomplete', 'reasons': ['Один из показателей не подтверждён числом.'], 'metrics': []}
        output.append({'id': key, 'label': first['label'], 'unit': first['unit'], 'first': x, 'second': y,
                       'absolute': y - x, 'percent': round((y - x) / x * 100, 1) if x != 0 else None})
    warnings = []
    if independent_live:
        warnings.append('Данные получены отдельными запросами. Источник не предоставляет общую версию среза; изменения между запросами возможны.')
    if any(item.get('completeness_verified') is False for item in (left, right)):
        warnings.append('Сравниваются записи сохранённого архива. Полнота поступления всех жалоб в архив не подтверждена.')
    try:
        periods = [item['requested_period'] for item in (left, right)]
        lengths = [(date.fromisoformat(item.get('to', item.get('end'))) - date.fromisoformat(item.get('from', item.get('start')))).days + 1 for item in periods]
        if any(length <= 0 for length in lengths) or any(item.get('effective_period') != item.get('requested_period') for item in (left, right)):
            return {'status': 'incomplete', 'reasons': ['Фактические границы отличаются от запрошенных или некорректны.'], 'metrics': []}
        if lengths[0] != lengths[1]:
            warnings.append('Продолжительность периодов различается (' + str(lengths[0]) + ' и ' + str(lengths[1]) + ' дней). Это разница суммарных значений, без нормирования на день.')
    except (KeyError, TypeError, ValueError):
        return {'status': 'incomplete', 'reasons': ['Границы периодов не подтверждены.'], 'metrics': []}
    return {'status': 'complete', 'metrics': output, 'warnings': warnings, 'reasons': []}


def _fmt(number):
    return str(number).replace('.', ',')


def _window(value):
    if not isinstance(value, dict):
        return 'не подтверждён'
    if value.get('label'):
        return str(value['label'])
    return str(value.get('from', value.get('start', '—'))) + ' — ' + str(value.get('to', value.get('end', '—')))


def _distribution(items, heading):
    """Bounded plain labels and integer counts; no source HTML or Markdown."""
    rows = []
    for label, count in islice(items, 100):
        if (not isinstance(label, str) or type(count) is not int or not 0 <= count <= 2**63-1
                or not 1 <= len(label.strip()) <= 160 or not label.isprintable()
                or any(char in label for char in '<>[]`*_\\#') or '://' in label):
            continue
        rows.append('• ' + label.strip() + ' — ' + str(count) + '.')
        if len(rows) == 20:
            break
    return [heading, *rows] if rows else []


def render(state):
    lines = []
    for index, step in enumerate(state['steps'], 1):
        spec = tool_catalog.spec(step['tool'])
        params = step['params']
        lines.append(str(index) + '. ' + spec['title'])
        filters = ', '.join(str(params[key]) for key in ('municipality', 'start', 'end') if params.get(key))
        service = {'all':'ХВС, ГВС и водоотведение','hvs':'ХВС','gvs':'ГВС','vo':'водоотведение'}.get(params.get('service'))
        if filters or service:
            lines.append('Фильтры — ' + filters + ('; услуга — ' + service if service else ''))
        result = step.get('result')
        if not result:
            lines.append({'queued':'Ожидает выполнения.', 'running':'Выполняется.', 'failed':'Результат не подтверждён.'}.get(step['status'], 'Выполнение остановлено.'))
            continue
        if step.get('run_id'):
            lines.append('Запуск ' + step['run_id'])
        if spec['operation'] in ('analytics', 'query', 'check'):
            if result.get('requested_period'):
                lines.append('Запрошено ' + _window(result.get('requested_period')) + '. Данные выборки — ' + _window(result.get('effective_period')) + '.')
            else:
                lines.append('Текущее состояние, полученное при выполнении этой задачи.')
            if params.get('topic'):
                lines.append('Тематика — ' + MINGKH_TOPICS.get(params['topic'], ('тема не подтверждена', {}))[0] + '.')
            if params.get('status'):
                lines.append('Статус — ' + params['status'] + '.')
            lines.append('ОМСУ — ' + (params.get('municipality') or 'вся область') + '.')
            for metric in result.get('metrics', []):
                if _number(metric.get('value')):
                    lines.append(metric['label'] + ' — ' + _fmt(metric['value']) + ' ' + metric['unit'] + '.')
            if isinstance(result.get('status_counts'), dict):
                lines.extend(_distribution(result['status_counts'].items(), 'По статусам'))
            if isinstance(result.get('topic_breakdown'), list):
                lines.extend(_distribution(((row.get('label'), row.get('count')) for row in result['topic_breakdown']
                                            if isinstance(row, dict)), 'Тематика обращений по классификации источника'))
            lines.extend(result.get('warnings', []))
            if result.get('message'):
                lines.append(result['message'])
            if result.get('answer'):
                lines.append(result['answer'])
            if spec['operation'] == 'check':
                for source in (result.get('evidence') or {}).get('sources', []):
                    for metric in source.get('metrics', [])[:12]:
                        if _number(metric.get('value')):
                            lines.append(metric['label'] + ' — ' + _fmt(metric['value']) + ' ' + str(metric.get('unit') or '') + '.')
            for key, label in (('available_municipalities','Доступные ОМСУ'),('available_statuses','Доступные статусы')):
                if result.get(key):
                    lines.append(label + ' — ' + ', '.join(str(item) for item in result[key][:30]) + '.')
            if result.get('rating_rows') and not params.get('municipality'):
                for row in result['rating_rows'][:60]:
                    lines.append(str(row.get('municipality',row.get('name',''))) + ' — ' + str(row.get('value', '—')) + ' ' + str(row.get('unit','')) + ', место ' + str(row.get('place_label', row.get('place','—'))) + '.')
        else:
            lines.append(result.get('answer', 'Результат не подтверждён.'))
            for source in (result.get('evidence') or {}).get('sources', []):
                for metric in source.get('metrics', [])[:12]:
                    if _number(metric.get('value')):
                        lines.append(metric['label'] + ' — ' + _fmt(metric['value']) + ' ' + str(metric.get('unit') or '') + '.')
        lines.append('[Открыть блок](' + spec['url'] + ')')
        lines.append('')
    diff = state.get('comparison')
    if diff:
        if diff['status'] == 'complete':
            lines.append('Разница второго периода относительно первого')
            for metric in diff['metrics']:
                percent = (_fmt(metric['percent']) + '%') if metric['percent'] is not None else 'процент не рассчитан — исходное значение равно нулю'
                lines.append(metric['label'] + ' — ' + ('+' if metric['absolute'] > 0 else '') + _fmt(metric['absolute']) + ' ' + metric['unit'] + ', ' + percent + '.')
            lines.extend(diff.get('warnings', []))
        else:
            lines.append('Сравнение не завершено.')
            lines.extend(diff['reasons'])
    return '\n'.join(lines)


def execute(user, plan, ticket, publish, *, confirmed=False, runner=None):
    user = jobs.current_user(user)
    plan = action_planner.validate_plan(user, plan)
    specs = [tool_catalog.spec(step['tool']) for step in plan['steps']]
    if any(item['requires_confirmation'] for item in specs) and confirmed is not True:
        raise HTTPException(409, 'Подтвердите сохранённый план перед запуском.')
    required = sorted({item['module_id'] for item in specs})
    ident = ticket['job_id']
    budget = min(21600, 300 + sum(item.get('timeout_seconds', 120) for item in specs))
    if not jobs.set_budget(ident, budget):
        raise HTTPException(409, 'Задача уже завершена или прервана.')
    state = {'version': 1, 'compare': plan['compare'], 'status': 'running', 'started_at': time.time(),
             'steps': [{**step, 'status': 'queued'} for step in plan['steps']]}

    def result():
        return {'ok': state['status'] == 'succeeded', 'response_kind': 'answer' if state['status'] == 'succeeded' else 'error',
                'required_modules': required, 'agent_plan': state, 'answer': render(state),
                'report_scope': {'agent_action': 'execution_plan', 'modules': required},
                'report_sources': [{'id': item['id'], 'title': item['title'], 'source_url': item['url']} for item in specs]}

    if not jobs.plan_checkpoint(ident, result()):
        raise HTTPException(409, 'Не удалось сохранить план задачи.')
    for index, (step, spec) in enumerate(zip(state['steps'], specs), 1):
        user = jobs.current_user(user)
        if not private_storage.can_view(user, {'required_modules': required}):
            raise HTTPException(403, 'Права на данные изменились. Выполнение остановлено.')
        if not jobs.alive(ident):
            raise HTTPException(409, 'Задача прервана. Повторный запуск не выполнялся.')
        tool_catalog.validate(user, step['tool'], step['params'])
        step.update(status='running', started_at=time.time())
        def stage(label):
            publish({'id': 'plan-' + str(index), 'kind': 'tool', 'label': 'Шаг ' + str(index) + ' из ' + str(len(specs)) + ' — ' + label, 'status': 'running'})
        stage(spec['title'])
        if spec['operation'] in ('refresh','check'):
            from services import run_history
            step['run_id'] = run_history.record_start(spec['runner_module'], user='ИИ-чат', params={key:value for key,value in step['params'].items() if key in ('start','end')})
        if not jobs.plan_checkpoint(ident, result()):
            if step.get('run_id'):
                run_history.record_finish(step['run_id'], 'interrupted')
            raise HTTPException(409, 'Выполнение прервано до запуска шага.')
        try:
            if spec['operation'] in ('refresh','check') and runner is None:
                from app import run_subprocess_worker
                runner = run_subprocess_worker
            outcome = tool_catalog.execute(user, step['tool'], step['params'], run_id=step.get('run_id'),
                                           on_progress=stage, run_subprocess=runner, confirmed=confirmed)
            outcome.pop('_output_version', None)
            step['result'] = outcome
            if spec['operation'] in ('refresh', 'check'):
                success = outcome.get('status') == 'success' and outcome.get('fresh_result') is True
                step['status'] = 'succeeded' if success else 'failed'
                run_history.record_finish(step['run_id'], 'success' if success else 'skipped' if outcome.get('error_code') == 'BLOCK_BUSY' else 'error')
            else:
                step['status'] = 'succeeded' if outcome.get('status') == 'ok' else 'partial' if outcome.get('status') == 'partial' else 'failed'
        except (tool_catalog.AccessDenied, HTTPException):
            if step.get('run_id'):
                run_history.record_finish(step['run_id'], 'interrupted')
            raise
        except Exception:
            step.update(status='failed', result={'status': 'error', 'message': 'Не удалось выполнить шаг. Проверьте источник на странице блока.',
                                                'answer': 'Не удалось выполнить шаг. Проверьте источник на странице блока.'})
            if step.get('run_id'):
                run_history.record_finish(step['run_id'], 'error')
        step['finished_at'] = time.time()
        # Rights may have changed while the source was being read. Do not persist
        # a new result after revocation; old checkpoints remain access-protected.
        user = jobs.current_user(user)
        if not private_storage.can_view(user, {'required_modules': required}):
            raise HTTPException(403, 'Права на данные изменились. Выполнение остановлено.')
        publish({'id': 'plan-' + str(index), 'kind': 'tool', 'label': 'Шаг ' + str(index) + ' — ' +
                 ('результат подтверждён' if step['status'] == 'succeeded' else 'данные подтверждены частично' if step['status'] == 'partial' else 'результат не подтверждён'),
                 'status': 'done' if step['status'] == 'succeeded' else 'error'})
        if not jobs.plan_checkpoint(ident, result()):
            raise HTTPException(409, 'Задача прервана; новые действия не выполнялись.')
    if state['compare']:
        state['comparison'] = comparison(*(step.get('result') or {} for step in state['steps']))
    state['finished_at'] = time.time()
    state['status'] = 'succeeded' if all(step['status'] == 'succeeded' for step in state['steps']) and (not state['compare'] or state['comparison']['status'] == 'complete') else 'partial'
    final = result()
    # Fully executed plans, including known partial/error outcomes, are final
    # evidence. A crash while returning cannot turn them into an unknown result.
    jobs.checkpoint(ident, final)
    return final
