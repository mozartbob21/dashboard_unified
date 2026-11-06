"""Chat adapters for the existing, fixed module runners.

Only the original user message may be passed to ``recognize``. Recognition
creates a proposal, never executes it. ``execute`` is called only for a stored,
explicitly confirmed proposal; models/documents cannot supply commands or URLs.
"""
from datetime import date, datetime, timedelta
from pathlib import Path
import os
import re
from zoneinfo import ZoneInfo

from core.roles import check_module_access

BASE = Path(__file__).resolve().parents[2]
SPECS = {
    'edds': {'title': 'ЕДДС — обновить жалобы', 'path': 'data/edds/water_daily.json', 'module': 'edds', 'url': '/edds', 'timeout': 1900},
    'mingkh-water-map': {'title': 'МинЖКХ — обновить карту жалоб', 'path': 'data/mingkh/water-map.json.gz', 'module': 'mingkh', 'url': '/mingkh/water-map', 'timeout': 1300},
    'edo': {'title': 'Заполненность данных', 'path': 'data/edo/result.json', 'module': 'edo', 'url': '/edo', 'timeout': 3600},
    'overdue': {'title': 'Просроченные задачи', 'path': 'data/overdue/final_result.json', 'module': 'overdue', 'url': '/overdue', 'timeout': 3600},
    'watercontrol': {'title': 'Контроль воды', 'path': 'data/watercontrol/result.json', 'module': 'watercontrol', 'url': '/watercontrol', 'timeout': 3600},
    'utnkr': {'title': 'Технадзор УТНКР', 'path': 'data/utnkr/violators.json', 'module': 'utnkr', 'url': '/utnkr', 'timeout': 3600},
    'cameras': {'title': 'Проверка камер', 'path': 'data/cameras/state/dashboard_state.json', 'module': 'cameras', 'url': '/cameras', 'timeout': 7200},
    'cds': {'title': 'Обращения 1С', 'path': 'data/cds/result.json', 'module': 'cds', 'url': '/cds', 'timeout': 3600},
}
ERROR_MESSAGES = {
    'BLOCK_TIMEOUT': 'Достигнут предел времени выполнения блока. Исполнитель остановлен; новый результат не подтверждён.',
    'BLOCK_EXIT_FAILED': 'Исполнитель блока завершился с ошибкой. Проверьте доступ к источнику и состояние проверки на странице блока.',
    'BLOCK_BUSY': 'Блок уже выполняется. Дождитесь завершения текущей проверки; второй запуск не создан.',
    'BLOCK_FAILED': 'Не удалось выполнить блок. Проверьте его состояние и настройки источника.',
    'BLOCK_CREDENTIALS_MISSING': 'Доступ к Доброделу не настроен. Администратору нужно сохранить логин и пароль в разделе настроек доступа.',
    'BLOCK_SOURCE_NOT_CONFIGURED': 'Источник технадзора УТНКР не настроен. Администратору нужно указать адрес источника; старые локальные строки не считаются новой проверкой.',
}
_ALIASES = {
    'edds': r'\bеддс\b',
    'mingkh-water-map': r'\bкарт[ауые]\s+(?:жалоб(?:\s+мин\s*жкх)?|мин\s*жкх)\b',
    'edo': r'\b(?:заполненность\s+данных|эдо)\b',
    'overdue': r'\b(?:просроченн\w*\s+задач\w*|просрочк\w*)\b',
    'watercontrol': r'\bконтрол[ья]\s+воды\b',
    'utnkr': r'\b(?:утнкр|технадзор)\b',
    'cameras': r'\bкамер\w*\b',
    'cds': r'\b(?:обращения\s+1с|цдс|1с)\b',
}
_INTENT = re.compile(r'^(?:пожалуйста[, ]+)?(?:(?:можешь|можете)\s+)?(?:запусти(?:ть|те)?|обнови(?:ть|те)?|проверь(?:те)?|проверить|выполни(?:ть|те)?)\s+', re.I)
_DATE = re.compile(r'(?<!\d)(?:\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})(?!\d)')


class InvalidParameters(ValueError):
    pass


class AccessDenied(PermissionError):
    pass


def permitted(user, module_id):
    if module_id not in SPECS or not check_module_access(user, module_id):
        return False
    if module_id == 'mingkh-water-map':
        from services.auth.accounts import is_account_manager
        return is_account_manager(user)
    return True


def catalog(user):
    return [{'module_id': key, 'title': spec['title'], 'parameters': ['start', 'end'] if key == 'cds' else [],
             'description': 'Получение данных из настроенного источника; без отправки сообщений и публикации документов.',
             'requires_confirmation': True}
            for key, spec in SPECS.items() if permitted(user, key)]


def validate(user, module_id, params):
    if not isinstance(module_id, str) or module_id not in SPECS:
        raise InvalidParameters('Этот блок не поддерживает запуск из чата. Используйте его страницу.')
    if not permitted(user, module_id):
        raise AccessDenied('Недостаточно прав для запуска этого блока.')
    if not isinstance(params, dict):
        raise InvalidParameters('Некорректные параметры запуска.')
    if module_id != 'cds':
        if params:
            raise InvalidParameters('Этот блок запускается целиком с сохранёнными настройками, без дополнительных параметров.')
        return {}
    if set(params) != {'start', 'end'}:
        raise InvalidParameters('Укажите начало и конец периода для «Обращения 1С», например 2026-11-01 — 2026-11-04.')
    try:
        if any(type(params[key]) is not str or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', params[key]) for key in params):
            raise ValueError()
        first, last = date.fromisoformat(params['start']), date.fromisoformat(params['end'])
        if first > last or (last - first).days > 365:
            raise ValueError()
    except ValueError:
        raise InvalidParameters('Период должен содержать корректные даты от начала к концу и не превышать 366 дней.') from None
    return {'start': first.isoformat(), 'end': last.isoformat()}


def _as_date(value):
    return (datetime.strptime(value, '%d.%m.%Y').date() if '.' in value else date.fromisoformat(value)).isoformat()


def recognize(text, user):
    """Propose only anchored direct commands; never inspect history/attachments."""
    if not isinstance(text, str) or len(text) > 2000:
        return None
    text = text.strip()
    prefix = _INTENT.match(text)
    if not prefix:
        return None
    body = text[prefix.end():].casefold().replace('ё', 'е')
    # Quoted instructions, multiline documents, shell/URL fragments are data.
    if any(fragment in body for fragment in ('\n', '```', '://', '«', '»', '"', '<', '>')):
        return None
    found = [(key, re.search(pattern, body)) for key, pattern in _ALIASES.items()]
    found = [(key, match) for key, match in found if match]
    if len(found) != 1:
        if not found:
            return None
        return {'clarification': 'Укажите один блок для запуска. Каждый запуск подтверждается отдельно.'}
    module_id, match = found[0]
    if not permitted(user, module_id):
        return {'clarification': 'Недостаточно прав для запуска этого блока. Обратитесь к администратору.'}
    remainder = (body[:match.start()] + ' ' + body[match.end():]).strip()
    remainder = re.sub(r'\b(?:блок|раздел|данные|сбор|проверку|обновление|пожалуйста)\b', ' ', remainder)
    remainder = re.sub(r'\bи\s+(?:покажи|пришли)\s+результат\b', ' ', remainder)
    params = {}
    if module_id == 'cds':
        dates = _DATE.findall(remainder)
        try:
            if len(dates) in (1, 2):
                params = {'start': _as_date(dates[0]), 'end': _as_date(dates[-1])}
                remainder = _DATE.sub(' ', remainder)
                remainder = re.sub(r'\b(?:с|по|за|период|от|до)\b', ' ', remainder)
            elif re.search(r'\b(?:вчера|сегодня)\b', remainder):
                today = datetime.now(ZoneInfo('Europe/Moscow')).date()
                day = today - timedelta(days=1) if 'вчера' in remainder else today
                params = {'start': day.isoformat(), 'end': day.isoformat()}
                remainder = re.sub(r'\b(?:вчера|сегодня|за)\b', ' ', remainder)
        except ValueError:
            return {'clarification': 'В периоде есть некорректная дата. Укажите начало и конец периода.'}
    if remainder.strip(' .,!?:;—–-'):
        return {'clarification': 'Уточните параметры запуска. Для «Обращения 1С» доступны даты начала и конца; остальные блоки запускаются целиком с сохранёнными настройками.'}
    try:
        params = validate(user, module_id, params)
    except (InvalidParameters, AccessDenied) as exc:
        return {'clarification': str(exc)}
    summary = SPECS[module_id]['title']
    if params:
        summary += ': ' + params['start'] + ' — ' + params['end']
    return {'module_id': module_id, 'params': params, 'summary': summary}


def _fingerprint(path):
    try:
        value = path.stat()
        return value.st_mtime_ns, value.st_size, value.st_ino
    except OSError:
        return None


def _evidence(module_id):
    """Reuse reviewed local projections, never feed raw subprocess output to AI."""
    from services.aichat import report_context
    spec = SPECS[module_id]
    result = report_context.prepare('', [spec['module']], requested={'modules': [spec['module']]})
    if module_id == 'mingkh-water-map':
        result['sources'] = [source for source in result.get('sources', []) if source.get('id') == 'mingkh-water-map']
    if module_id == 'edds':
        result['sources'] = [source for source in result.get('sources', []) if source.get('id') == 'edds-complaints']
    return result


def execute(user, module_id, params, *, run_id, on_progress=None, run_subprocess=None):
    """Run existing module code, with its existing validation/locks/credentials.

    The caller persists and authorizes the confirmed action, supplies a server
    history ID and owns the worker lease. A failed result is never auto-retried.
    The injected subprocess worker is the same one used by normal module buttons.
    """
    params = validate(user, module_id, params)
    if not isinstance(run_id, str) or not re.fullmatch(r'[a-zA-Z0-9_-]{16,100}', run_id):
        raise InvalidParameters('Не удалось определить идентификатор запуска.')
    spec = SPECS[module_id]
    path = BASE / spec['path']
    before = _fingerprint(path)
    def progress(stage):
        if callable(on_progress):
            on_progress(stage)
    def failed(code):
        code = code if code in ERROR_MESSAGES else 'BLOCK_FAILED'
        return {'status': 'error', 'error_code': code, 'run_id': run_id, 'module_id': module_id,
                'url': spec['url'], 'answer': ERROR_MESSAGES[code], 'evidence': None}
    if module_id == 'utnkr' and not os.environ.get('SITE_TABLE_URL', '').strip().strip("\"'"):
        return failed('BLOCK_SOURCE_NOT_CONFIGURED')
    if module_id in ('edds', 'mingkh-water-map'):
        from services.auth.integrations import credentials
        try:
            configured = credentials('edds')
            ready = isinstance(configured, dict) and all(isinstance(configured.get(key), str) and configured[key].strip()
                                                        for key in ('username', 'password'))
        except Exception:
            ready = False
        if not ready:
            return failed('BLOCK_CREDENTIALS_MISSING')
    progress('Запуск существующего блока')
    if module_id == 'edds':
        from services.edds.runner import AlreadyRunning, run
        try:
            ok = run(user='ИИ-чат', run_id=run_id, raise_on_busy=True)
        except AlreadyRunning:
            return failed('BLOCK_BUSY')
        outcome = {'status': 'success' if ok else 'error'}
    elif module_id == 'mingkh-water-map':
        from services.mingkh.water_map_runner import run
        outcome = run(user='ИИ-чат', run_id=run_id)
    else:
        from services.scheduler import MODULE_COMMANDS
        command = list(MODULE_COMMANDS[module_id])
        if module_id == 'utnkr':
            # The collector reloads configuration in the child process. Enforce
            # live-source mode there too if its configuration changed meanwhile.
            command += ['--require-live-source']
        if module_id == 'cds':
            command += [date.fromisoformat(params[key]).strftime('%d.%m.%Y') for key in ('start', 'end')]
        if not callable(run_subprocess):
            raise RuntimeError('Не подключён штатный исполнитель блоков.')
        outcome = run_subprocess(module_id, command, run_id=run_id,
                                 progress_callback=on_progress, timeout_seconds=spec['timeout'])
    if not isinstance(outcome, dict) or outcome.get('status') != 'success':
        # Never return raw exception bodies, remote HTML, subprocess logs or URLs.
        code = outcome.get('code') if isinstance(outcome, dict) else None
        if isinstance(outcome, dict) and outcome.get('status') in ('busy', 'skipped'):
            code = 'BLOCK_BUSY'
        return failed(code)
    after = _fingerprint(path)
    result = {'status': 'success', 'run_id': run_id, 'module_id': module_id, 'url': spec['url'],
              'answer': 'Блок «' + spec['title'] + '» завершил выполнение.', 'evidence': None, 'fresh_result': False}
    if after is not None and after != before:
        try:
            evidence = _evidence(module_id)
            sources = evidence.get('sources') or []
            if sources and all(source.get('status') == 'saved' for source in sources):
                result.update(evidence=evidence, fresh_result=True, _output_version=list(after))
                result['answer'] += ' Новый сохранённый результат подтверждён.'
        except Exception:
            pass
    if not result['fresh_result']:
        result['answer'] += ' Новый сохранённый результат не подтверждён; прежние показатели не выдаются за результат этого запуска.'
    return result
