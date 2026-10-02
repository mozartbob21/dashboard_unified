"""GigaChat REST client, used only by the summarizer.

OAuth: https://developers.sber.ru/docs/ru/gigachat/api/reference/rest/post-token
Chat: https://developers.sber.ru/docs/ru/gigachat/api/reference/rest/post-chat
"""
import hashlib
import math
import os
import ssl
import threading
import time
import uuid

import httpx

from core.ai_errors import AIServiceError
from core.privacy import PrivacyError

AUTH_URL = 'https://ngw.devices.sberbank.ru:9443/api/v2/oauth'
DEFAULT_BASE = 'https://api.giga.chat/v1'
ALLOWED_BASES = {DEFAULT_BASE, 'https://gigachat.devices.sberbank.ru/api/v1'}
_TOKENS = {}
_LOCK = threading.Lock()


def _tls():
    try:
        return ssl.create_default_context(cafile=os.getenv('GIGACHAT_CA_BUNDLE_FILE') or os.getenv('AI_CA_BUNDLE') or None)
    except (OSError, ValueError):
        raise AIServiceError('AI_TLS_ERROR') from None


def _token(credentials, scope, tls, *, rejected=None):
    cache_key = hashlib.sha256((credentials + '\0' + scope).encode()).hexdigest()
    with _LOCK:
        cached = _TOKENS.get(cache_key)
        if cached and cached['expires_at'] > time.time() + 60 and cached['access_token'] != rejected:
            return cached['access_token']
        _TOKENS.pop(cache_key, None)
        response = httpx.post(
            AUTH_URL, data={'scope': scope},
            headers={'Authorization': 'Basic ' + credentials, 'RqUID': str(uuid.uuid4()), 'Accept': 'application/json'},
            timeout=30, verify=tls, trust_env=False, follow_redirects=False)
        response.raise_for_status()
        try:
            data = response.json()
            token = data['access_token']
            expires = float(data['expires_at']) / 1000
            if not isinstance(token, str) or not token.strip() or not math.isfinite(expires) or expires <= time.time():
                raise ValueError()
        except (KeyError, TypeError, ValueError):
            raise AIServiceError('AI_INVALID_RESPONSE') from None
        if len(_TOKENS) >= 8:
            _TOKENS.clear()
        _TOKENS[cache_key] = {'access_token': token, 'expires_at': expires}
        return token


def chat(messages, creds=None, model=None, max_tokens=2000):
    credentials = (creds or os.getenv('SUMMARIZER_GIGACHAT_CREDENTIALS', '').strip()
                   or os.getenv('GIGACHAT_CREDENTIALS', '').strip()).strip()
    if not credentials:
        raise AIServiceError('GIGACHAT_CREDENTIALS_MISSING')
    scope = os.getenv('GIGACHAT_SCOPE', '').strip() or 'GIGACHAT_API_PERS'
    if scope not in {'GIGACHAT_API_PERS', 'GIGACHAT_API_B2B', 'GIGACHAT_API_CORP'}:
        raise AIServiceError('GIGACHAT_SCOPE_INVALID')
    base = os.getenv('GIGACHAT_BASE_URL', '').strip().rstrip('/') or DEFAULT_BASE
    if base not in ALLOWED_BASES:
        raise PrivacyError('Недопустимый адрес GigaChat.')
    selected_model = (model or os.getenv('GIGACHAT_MODEL', '').strip()
                      or os.getenv('GIGACHAT_MODE', '').strip() or 'GigaChat-2-Pro')
    tls = _tls()
    token = _token(credentials, scope, tls)
    payload = {'model': selected_model, 'messages': messages, 'temperature': 0.4,
               'max_tokens': max_tokens, 'stream': False}
    for attempt in range(2):
        response = httpx.post(base + '/chat/completions', json=payload,
                              headers={'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json'},
                              timeout=180, verify=tls, trust_env=False, follow_redirects=False)
        if response.status_code == 401 and attempt == 0:
            token = _token(credentials, scope, tls, rejected=token)
            continue
        response.raise_for_status()
        try:
            content = response.json()['choices'][0]['message']['content']
            if not isinstance(content, str):
                raise TypeError()
        except (KeyError, IndexError, TypeError, ValueError):
            raise AIServiceError('AI_INVALID_RESPONSE') from None
        if not content.strip():
            raise AIServiceError('AI_EMPTY_RESPONSE')
        return content.strip()
