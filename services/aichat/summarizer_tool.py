"""Invoke the existing Summarizer engine from a permitted current-user request.

The caller extracts the text from an explicit instruction and keeps it as data.
No agent planner is called on the source text or on the generated summary.
"""
from datetime import datetime, timezone

from fastapi import HTTPException

from core.roles import check_module_access
from services.summarizer import engine, storage

ID = 'transform:summarizer'
MIN_TEXT = 50  # Same minimum as the module's normal HTTP endpoint.
MAX_TEXT = min(40000, storage.MAX_SOURCE_CHARS)  # Current chat message bound.


def validate(user, tool_id, params):
    if tool_id != ID:
        raise ValueError('Этот инструмент суммирования не подключён.')
    if not user or not user.get('is_active', True) or not check_module_access(user, 'summarizer'):
        raise HTTPException(403, 'Нет доступа к блоку «Сумматор».')
    if not isinstance(params, dict) or set(params) - {'text', 'backend'}:
        raise ValueError('Передайте текст для сумматора. Произвольные настройки модели недоступны.')
    text = params.get('text')
    if not isinstance(text, str) or not MIN_TEXT <= len(text.strip()) <= MAX_TEXT:
        raise ValueError('Для сумматора нужен текст от 50 до 40 000 символов. Разделите большой текст на части.')
    result = {'text': text.strip()}
    if 'backend' in params:
        if type(params['backend']) is not str or params['backend'] not in ('qwen', 'algo'):
            raise ValueError('Доступен Госчат/Qwen или штатный локальный алгоритм сумматора.')
        result['backend'] = params['backend']
    return result


def execute(user, tool_id, params, on_progress=None):
    params = validate(user, tool_id, params)
    result = {'tool_id': ID, 'module': 'summarizer', 'status': 'missing',
              'source': {'id': ID, 'title': 'Сумматор — оперативная сводка', 'url': '/summarizer'},
              'origin': 'block_execution', 'metrics': [], 'warnings': [],
              'used_model': False, 'fallback': False}
    if callable(on_progress):
        on_progress('Передаём текст штатному сумматору')
    try:
        summary = engine.summarize(params['text'], backend=params.get('backend'))
    except Exception:
        # Source text and exception bodies may contain sensitive data.
        result.update(error_code='SUMMARIZER_EXECUTION_FAILED',
                      answer='Блок «Сумматор» не завершил обработку. Результат не подтверждён; повторите запрос позднее.')
        return result
    if (not isinstance(summary, dict) or summary.get('ok') is not True or
            summary.get('backend') not in ('qwen', 'algo') or
            not isinstance(summary.get('report_text'), str) or not summary['report_text'].strip()):
        result.update(error_code='SUMMARIZER_RESULT_UNCONFIRMED',
                      answer='Блок «Сумматор» не вернул подтверждённую сводку. Проверьте исходный текст и состояние блока.')
        return result
    backend = summary['backend']
    warning = summary.get('warning')
    if isinstance(warning, str) and warning.strip():
        result['warnings'].append(warning[:1000])
    if backend == 'algo' and not result['warnings']:
        result['warnings'].append('Сводка подготовлена штатным локальным алгоритмом сумматора. ИИ-модель не использовалась.')
    text = summary['report_text']
    if len(text) > 64000:
        text = text[:64000]
        result['warnings'].append('Сводка слишком большая для одного сообщения. В чате показаны первые 64 000 символов.')
    result.update(status='ok', backend=backend, used_model=backend == 'qwen',
                  fallback=backend == 'algo' and summary.get('requested_backend') == 'qwen',
                  report_text=text, answer='\n\n'.join([*result['warnings'], text]),
                  as_of=datetime.now(timezone.utc).isoformat())
    if isinstance(summary.get('ai_error_code'), str):
        # Only the shared adapter's fixed error code may accompany a fallback.
        from core.ai_errors import MESSAGES
        code = summary['ai_error_code']
        if code in MESSAGES:
            result['ai_error_code'] = code
    if callable(on_progress):
        on_progress('Штатный сумматор подготовил сводку' if backend == 'qwen' else 'Сумматор подготовил алгоритмическую сводку')
    return result
