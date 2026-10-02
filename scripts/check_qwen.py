"""Read-only Qwen configuration diagnostics; never display configuration values.

Run from the project: python scripts/check_qwen.py
Add --probe for one tiny request to the configured AI service. The request does
not include chat history, documents or platform data. Neither the response text
nor credentials, URLs, model names or key fingerprints are printed.
"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
from collections import Counter

PROJECT_ROOT = Path(__file__).resolve().parents[1]
ENV_FILE = PROJECT_ROOT / '.env'
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Restrict names too: an accidentally pasted secret must not become an output key.
VARIABLES = ('QWEN_API_BASE', 'QWEN_BASE_URL', 'QWEN_API_KEY', 'QWEN_MODEL', 'AI_CA_BUNDLE',
             'SUMMARIZER_QWEN_API_KEY')


def _file_settings(env_file):
    from dotenv.main import DotEnv
    from dotenv.parser import parse_stream
    values, duplicates, invalid = {}, [], 0
    if not env_file.is_file():
        return values, duplicates, invalid, 'not_found'
    try:
        with env_file.open(encoding='utf-8') as source:
            bindings = list(parse_stream(source))
        counts = Counter(binding.key for binding in bindings
                         if binding.key in VARIABLES and not binding.error)
        duplicates = sorted(name for name, count in counts.items() if count > 1)
        invalid = sum(bool(binding.error) for binding in bindings)
        # Same interpolation precedence and UTF-8 encoding as app.load_dotenv().
        # DotEnv.dict reads values but does not put them into os.environ.
        import io
        source = ''.join(binding.original.string for binding in bindings if not binding.error)
        values = DotEnv(dotenv_path=None, stream=io.StringIO(source),
                        override=False, encoding='utf-8').dict()
        return values, duplicates, invalid, None
    except (OSError, UnicodeError):
        return values, duplicates, invalid, 'unreadable'


def diagnose(env_file=None):
    """Return presence, precedence and conflict flags only; do not mutate state."""
    env_file = Path(env_file) if env_file is not None else ENV_FILE
    file_values, duplicates, invalid, file_error = _file_settings(env_file)
    variables = {}
    effective = {}
    for name in VARIABLES:
        process_present = name in os.environ
        file_present = name in file_values and file_values[name] is not None
        process_value = os.environ.get(name, '').strip()
        file_value = (file_values.get(name) or '').strip()
        source = 'process' if process_present else 'dotenv' if file_present else 'unset'
        effective[name] = process_value if process_present else file_value
        variables[name] = {
            'in_process': process_present,
            'in_dotenv': file_present,
            'process_has_value': bool(process_value),
            'dotenv_has_value': bool(file_value),
            'source': source,
            'conflict': bool(process_present and file_present and process_value != file_value),
        }
    endpoint_variable = next((name for name in ('QWEN_API_BASE', 'QWEN_BASE_URL')
                              if effective[name]), None)
    key = effective['QWEN_API_KEY']
    warnings = []
    if any(item['conflict'] for item in variables.values()):
        warnings.append('PROCESS_ENV_OVERRIDES_DOTENV')
    if key.lower().startswith('bearer '):
        warnings.append('KEY_CONTAINS_BEARER_PREFIX')
    if duplicates:
        warnings.append('DUPLICATE_VARIABLES_IN_DOTENV')
    if invalid:
        warnings.append('DOTENV_PARSE_ERRORS')
    if not key:
        warnings.append('KEY_NOT_CONFIGURED')
    return {
        'env_file_found': env_file.is_file(),
        'env_file_error': file_error,
        'variables': variables,
        'endpoint_source': (endpoint_variable + ':' + variables[endpoint_variable]['source'])
                           if endpoint_variable else 'default',
        'model_source': variables['QWEN_MODEL']['source'] if effective['QWEN_MODEL'] else 'default',
        'summarizer_override': {
            'configured': bool(effective['SUMMARIZER_QWEN_API_KEY']),
            'differs_from_chat_key': bool(effective['SUMMARIZER_QWEN_API_KEY']
                                          and effective['SUMMARIZER_QWEN_API_KEY'] != key),
            'source': variables['SUMMARIZER_QWEN_API_KEY']['source'],
        },
        'duplicate_variables': duplicates,
        'dotenv_parse_errors': invalid,
        'warnings': warnings,
    }


def probe(env_file=None, *, summarizer=False):
    """Send only a synthetic greeting through the same transport as the chat."""
    from dotenv import load_dotenv
    from core.ai_errors import describe_ai_error
    env_file = Path(env_file) if env_file is not None else ENV_FILE
    try:
        # Preserve the app's precedence; diagnostics must not silently fix conflicts.
        load_dotenv(dotenv_path=env_file, override=False, encoding='utf-8')
        from services.summarizer.engine import _qwen_chat
        override_key = (os.getenv('SUMMARIZER_QWEN_API_KEY', '').strip() or None) if summarizer else None
        _qwen_chat([{'role': 'user', 'content': 'Ответь: OK'}], max_tokens=256,
                   key=override_key)
    except Exception as exc:
        failure = describe_ai_error(exc)
        return {'ok': False, 'code': failure['code'], 'http_status': failure['http_status']}
    # _qwen_chat returns text only; do not fabricate an unobserved HTTP status.
    return {'ok': True, 'code': 'AI_OK', 'http_status': None}


def main(argv=None):
    parser = argparse.ArgumentParser(description='Безопасная проверка настроек ИИ без вывода значений.')
    parser.add_argument('--probe', action='store_true',
                        help='Отправить один короткий тестовый вопрос ИИ; без данных платформы.')
    args = parser.parse_args(argv)
    report = diagnose()
    if args.probe:
        report['probe'] = probe()
        if report['summarizer_override']['configured'] and report['summarizer_override']['differs_from_chat_key']:
            report['summarizer_probe'] = probe(summarizer=True)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if report.get('probe', {}).get('ok') is False else 0


if __name__ == '__main__':
    raise SystemExit(main())
