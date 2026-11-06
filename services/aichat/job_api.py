"""Authorized task reads and confirmation of server-stored launch proposals."""
from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse

from . import jobs, block_execution

router = APIRouter(prefix='/aichat/api')


def identity(request: Request):
    user = getattr(request.state, 'user', None)
    if not user:
        raise HTTPException(401, 'Требуется авторизация.')
    return user


@router.get('/blocks')
def blocks(user=Depends(identity)):
    return {'items': block_execution.catalog(user)}


@router.get('/tools')
def tools(user=Depends(identity)):
    from . import tool_catalog
    return JSONResponse({'items': tool_catalog.catalog(user), 'blocks': tool_catalog.inventory(user)}, headers={'Cache-Control': 'private, no-store'})


@router.get('/jobs')
def list_jobs(dialog_id: str = None, user=Depends(identity)):
    return JSONResponse({'items': jobs.list_jobs(user, dialog_id)}, headers={'Cache-Control': 'private, no-store'})


@router.get('/jobs/{ident}')
def status(ident: str, user=Depends(identity)):
    return JSONResponse(jobs.get(user, ident), headers={'Cache-Control': 'private, no-store'})


@router.post('/jobs/{ident}/confirm')
async def confirm(ident: str, request: Request, user=Depends(identity)):
    # Body cannot override any operation, owner, parameters or output path.
    try:
        payload = await request.json()
    except ValueError:
        payload = None
    if payload != {}:
        raise HTTPException(400, 'Подтвердите сохранённый запрос без изменения параметров.')
    parent = jobs.get(user, ident)
    proposal = (parent.get('result') or {}).get('action')
    if parent['status'] != 'succeeded' or not isinstance(proposal, dict):
        raise HTTPException(409, 'В этом сообщении нет ожидающего подтверждения запуска.')
    if proposal.get('kind') == 'agent_plan':
        from . import action_planner, orchestrator, tool_catalog
        try:
            plan = action_planner.validate_plan(jobs.current_user(user), proposal.get('plan'))
        except block_execution.AccessDenied as exc:
            raise HTTPException(403, str(exc)) from None
        except block_execution.InvalidParameters as exc:
            raise HTTPException(400, str(exc)) from None
        specs = [tool_catalog.spec(step['tool']) for step in plan['steps']]
        if not any(item['requires_confirmation'] for item in specs):
            raise HTTPException(409, 'Этот план не требует подтверждения запуска.')
        requirements = sorted({item['module_id'] for item in specs})
        # Preserve global single-module dedup for a one-step launch. Multi-step
        # plans also reuse the existing runners' cross-process execution locks.
        module_id = specs[0].get('runner_module') if len(specs) == 1 else None
        ticket, fresh = jobs.submit(user, 'confirm_' + ident, {'proposal': ident},
                                   dialog_id=parent['dialog_id'], module_id=module_id,
                                   confirmation_id=ident, required_modules=requirements)
        if fresh:
            jobs.start(ticket, lambda publish: orchestrator.execute(user, plan, ticket, publish, confirmed=True))
        return JSONResponse(ticket, status_code=202, headers={'Cache-Control': 'private, no-store'})
    module_id, params = proposal.get('module_id'), proposal.get('params')
    try:
        params = block_execution.validate(user, module_id, params)
    except block_execution.AccessDenied as exc:
        raise HTTPException(403, str(exc)) from None
    except block_execution.InvalidParameters as exc:
        raise HTTPException(400, str(exc)) from None
    ticket, fresh = jobs.submit(user, 'confirm_' + ident, {'proposal': ident},
                                dialog_id=parent['dialog_id'], module_id=module_id,
                                confirmation_id=ident, required_modules=[module_id])
    if fresh:
        jobs.start(ticket, lambda publish: run_block(user, module_id, params, ticket, publish))
    return JSONResponse(ticket, status_code=202, headers={'Cache-Control': 'private, no-store'})


def run_block(user, module_id, params, ticket, publish, *, runner=None, model=None):
    from services import run_history
    from .report_context import to_prompt
    user = jobs.current_user(user)
    block_execution.validate(user, module_id, params)
    if not jobs.alive(ticket['job_id']):
        raise HTTPException(409, 'Задача уже завершена или прервана.')
    def stage(label):
        publish({'id': 'block', 'kind': 'tool', 'label': label, 'status': 'running'})
    if runner is None:
        from app import run_subprocess_worker
        runner = run_subprocess_worker
    run_id = run_history.record_start(module_id, user='ИИ-чат', params=params)
    try:
        outcome = block_execution.execute(user, module_id, params, run_id=run_id,
                                          on_progress=stage, run_subprocess=runner)
    except Exception:
        run_history.record_finish(run_id, 'error', 'Выполнение из чата не подтверждено.')
        raise
    if outcome.get('error_code') == 'BLOCK_BUSY':
        run_history.record_finish(run_id, 'skipped', 'Проверка уже запущена; повтор не выполнялся.')
        result = {'ok': True, 'response_kind': 'clarification', 'run_id': run_id,
                  'module_id': module_id, 'required_modules': [module_id], 'already_running': True,
                  'answer': 'Проверка этого блока уже запущена. Повторный запуск не выполнялся. '
                            'Её состояние можно посмотреть на странице блока.', 'url': outcome['url']}
        jobs.checkpoint(ticket['job_id'], result)
        publish({'id': 'block', 'kind': 'tool', 'label': 'Проверка уже запущена; повтор не выполнялся', 'status': 'done'})
        return result
    ok = outcome['status'] == 'success'
    if not ok:
        # Includes runners refusing an already active operation.
        run_history.record_finish(run_id, 'error', 'Блок не подтвердил успешное завершение.')
    result = {'ok': ok, 'response_kind': 'answer' if ok else 'error', 'run_id': run_id,
              'answer': outcome['answer'], 'error_code': outcome.get('error_code'), 'module_id': module_id,
              'required_modules': [module_id], 'url': outcome['url']}
    if outcome.get('evidence'):
        result['evidence'] = outcome['evidence']
        result['report_scope'] = outcome['evidence'].get('selection')
        lines = []
        for source in outcome['evidence'].get('sources', []):
            lines.append(source.get('title', 'Результат блока'))
            for metric in source.get('metrics', [])[:12]:
                if metric.get('value') is not None:
                    lines.append('• ' + metric['label'] + ': ' + str(metric['value']) +
                                 (' ' + metric['unit'] if metric.get('unit') else ''))
        result['answer'] += '\n\n' + '\n'.join(lines)
    # Durable checkpoint precedes the optional model call. A server restart must
    # preserve the confirmed outcome, even if the model never returns.
    jobs.checkpoint(ticket['job_id'], result)
    if ok and outcome.get('evidence') and jobs.alive(ticket['job_id']):
        # The job and its proven result survive a separate model failure.
        stage('Результат сохранён; Нейрона готовит вывод')
        try:
            block_execution.validate(jobs.current_user(user), module_id, params)
            if model is None:
                from .engine import ask
                model = ask
            summary = model([{'role': 'user', 'content': 'Дай краткий вывод по фактическому результату завершённого блока. Не выдумывай показатели или действия.'}],
                            platform_context=to_prompt(outcome['evidence']))
            result['answer'] += '\n\n' + summary
        except Exception:
            result['answer'] += '\n\nВывод ИИ сейчас недоступен. Подтверждённый результат сохранён в блоке.'
    publish({'id': 'block', 'kind': 'tool', 'label': 'Выполнение подтверждено' if ok else 'Блок завершился с ошибкой',
             'status': 'done' if ok else 'error'})
    return result
