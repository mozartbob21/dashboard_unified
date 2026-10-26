"""Cached, model-generated educational facts, independent of every conversation.

The public functions never wait for inference. A single daemon worker refreshes
the generic cache only when no answer is active; an in-flight HTTP call remains
bounded by its own timeout. Importing this module does not start work.
"""
import json
import math
import os
import re
import tempfile
import threading
import time
from datetime import datetime, timezone

from services.aichat.storage import DATA_DIR


_TTL_SECONDS = 24 * 60 * 60
_COOLDOWN_SECONDS = 15 * 60
_MAX_CACHE_BYTES = 24000
_LOCK = threading.RLock()
_cache = None
_loaded = False
_enabled = False
_preparing = False
_active_answers = 0
_retry_at = 0.0

_PROMPT = (
    'Создай 8 коротких общеобразовательных фактов на русском языке: 4 об ИИ и 4 о ЖКХ. '
    'Это самостоятельные карточки для чтения во время ожидания. Используй только '
    'устойчивые, общеизвестные сведения, которые не зависят от даты, региона или '
    'текущих событий. Не используй статистику, числа, цены, нормативы, законы, '
    'медицинские, юридические или финансовые советы. Не давай инструкций по '
    'самостоятельному ремонту оборудования или работе с газом и электричеством. '
    'Не упоминай запрос пользователя, переписку, платформу или процесс ожидания. '
    'Не придумывай источники, не добавляй ссылки, HTML или Markdown. '
    'Каждый факт — одно-два простых предложения длиной от 20 до 320 символов. '
    'Верни только JSON без пояснений: {"items":[{"topic":"ИИ","text":"..."},'
    '{"topic":"ЖКХ","text":"..."}]}. Поле topic допускает только ИИ или ЖКХ.'
)
_FORBIDDEN_TEXT = re.compile(
    r'[<>\x00-\x1f\x7f]|\d|https?://|www\.|mailto:|\[[^\]]*\]\s*\('
    r'|\b(?:[a-z0-9-]+\.)+(?:[a-z]{2,}|рф)\b'
    r'|\b(?:сегодня|вчера|завтра|сейчас|ныне)\b', re.IGNORECASE)


def _validated_items(value):
    if not isinstance(value, list) or not 6 <= len(value) <= 8:
        raise ValueError('Invalid fact count')
    items, seen = [], set()
    counts = {'ИИ': 0, 'ЖКХ': 0}
    for item in value:
        if not isinstance(item, dict) or set(item) != {'topic', 'text'}:
            raise ValueError('Invalid fact object')
        topic, text = item['topic'], item['text']
        if not isinstance(topic, str) or topic not in counts or not isinstance(text, str):
            raise ValueError('Invalid fact fields')
        text = text.strip()
        if not 20 <= len(text) <= 320 or _FORBIDDEN_TEXT.search(text):
            raise ValueError('Invalid fact text')
        normalized = ' '.join(text.casefold().split())
        if normalized in seen:
            raise ValueError('Duplicate fact')
        seen.add(normalized)
        counts[topic] += 1
        items.append({'topic': topic, 'text': text})
    if min(counts.values()) < 2:
        raise ValueError('Both fact topics are required')
    return items


def _timestamp(value):
    if not isinstance(value, str) or len(value) > 40:
        raise ValueError('Invalid fact timestamp')
    stamp = datetime.fromisoformat(value.replace('Z', '+00:00'))
    if stamp.tzinfo is None:
        raise ValueError('Fact timestamp needs a timezone')
    result = stamp.timestamp()
    if not math.isfinite(result) or result > time.time() + 300:
        raise ValueError('Invalid fact timestamp')
    return result


def _load_locked():
    global _cache, _loaded
    if _loaded:
        return
    _loaded = True
    try:
        with (DATA_DIR / 'waiting_facts.json').open('rb') as handle:
            raw = handle.read(_MAX_CACHE_BYTES + 1)
        if len(raw) > _MAX_CACHE_BYTES:
            return
        value = json.loads(raw)
        if not isinstance(value, dict) or value.get('version') != 1:
            return
        items = _validated_items(value.get('items'))
        _timestamp(value.get('generated_at'))
        _cache = {'items': items, 'generated_at': value['generated_at']}
    except (OSError, ValueError, TypeError, OverflowError):
        pass


def _stale_locked():
    if _cache is None:
        return True
    try:
        return time.time() - _timestamp(_cache['generated_at']) >= _TTL_SECONDS
    except (ValueError, OverflowError):
        return True  # A clock correction must not break this optional feature.


def _generate_items():
    # This import and the network call happen only in the background worker.
    # No question, conversation, attachment or platform data enters this call.
    import httpx
    from services.summarizer.engine import _qwen_chat
    raw = _qwen_chat(
        [{'role': 'system', 'content': 'Ты составляешь проверяемые образовательные карточки. Следуй указанному JSON-формату.'},
         {'role': 'user', 'content': _PROMPT}],
        module_id="aichat", max_tokens=1200, timeout=httpx.Timeout(connect=10.0, read=45.0, write=10.0, pool=5.0))
    if not isinstance(raw, str) or len(raw.encode('utf-8')) > _MAX_CACHE_BYTES:
        raise ValueError('Invalid fact response')
    raw = raw.strip()
    if raw.startswith('```'):
        fenced = re.fullmatch(r'```(?:json)?\s*\n([\s\S]*?)\n```', raw, re.IGNORECASE)
        if not fenced:
            raise ValueError('Invalid fact response wrapper')
        raw = fenced.group(1)
    value = json.loads(raw)
    if not isinstance(value, dict) or set(value) != {'items'}:
        raise ValueError('Invalid fact response object')
    return _validated_items(value['items'])


def _save(value):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=DATA_DIR,
                                         prefix='.waiting-facts-', suffix='.tmp', delete=False) as handle:
            temporary = handle.name
            json.dump({'version': 1, **value}, handle, ensure_ascii=False)
        os.replace(temporary, DATA_DIR / 'waiting_facts.json')
    finally:
        if temporary and os.path.exists(temporary):
            os.unlink(temporary)


def _refresh():
    global _cache, _preparing, _retry_at
    # Release single-flight ownership under the same lock as the active check.
    # Otherwise the final answer could finish between the check and cleanup,
    # observe a worker still preparing, and miss its deferred refresh.
    with _LOCK:
        if _active_answers:
            _preparing = False
            return
    try:
        items = _generate_items()
        value = {'items': items, 'generated_at': datetime.fromtimestamp(
            time.time(), timezone.utc).isoformat(timespec='seconds')}
        _save(value)
        with _LOCK:
            _cache = value
            _retry_at = 0.0
    except Exception:
        # Never log prompts, provider responses, credentials or request details.
        # A refresh failure leaves the previously validated cache untouched.
        with _LOCK:
            _retry_at = time.monotonic() + _COOLDOWN_SECONDS
    finally:
        with _LOCK:
            _preparing = False


def _warm_locked():
    global _preparing, _retry_at
    if not _enabled or _active_answers or _preparing or time.monotonic() < _retry_at:
        return
    _load_locked()
    if not _stale_locked():
        return
    _preparing = True
    try:
        threading.Thread(target=_refresh, name='aichat-waiting-facts', daemon=True).start()
    except Exception:
        _preparing = False
        _retry_at = time.monotonic() + _COOLDOWN_SECONDS


def get_facts(*, warm=True):
    """Return facts immediately; warm=False reads without activating generation."""
    global _enabled
    with _LOCK:
        if warm:
            _enabled = True
        _load_locked()
        if warm:
            _warm_locked()
        return {'items': [dict(item) for item in _cache['items']] if _cache else [],
                'generated_at': _cache['generated_at'] if _cache else None,
                'preparing': _preparing}


def begin_answer():
    """Defer new background inference while a chat answer is active."""
    global _active_answers
    with _LOCK:
        _active_answers += 1


def end_answer(*, warm=True):
    """Release a chat; cancelled callers use warm=False to avoid starting work."""
    global _active_answers
    with _LOCK:
        _active_answers = max(0, _active_answers - 1)
        if warm:
            _warm_locked()
