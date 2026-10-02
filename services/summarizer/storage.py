"""Shared local reports: source, result, revisions and approval workflow."""
import copy
import json
import os
import tempfile
import threading
import time
import uuid
from pathlib import Path

DATA = Path(__file__).resolve().parents[2] / 'data' / 'summarizer'
FILE = DATA / 'reports.json'
MIN_APPROVALS = 2
MAX_SOURCE_CHARS = 200000
_LOCK = threading.RLock()


class ReportConflict(Exception):
    """The approval state changed while a new report was being generated."""


def _now():
    return time.strftime('%d.%m.%Y %H:%M:%S')


def _load():
    if not FILE.exists():
        return []
    # Do not overwrite a damaged archive with an empty history.
    items = json.loads(FILE.read_text(encoding='utf-8'))
    if not isinstance(items, list):
        raise ValueError('Некорректный формат истории сумматора')
    return items


def _save(items):
    DATA.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix='.reports-', suffix='.tmp', dir=DATA)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as stream:
            json.dump(items, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(name, FILE)
    finally:
        if os.path.exists(name):
            os.unlink(name)


def _version(item, *, source=None, comment='', user=None):
    return {
        'created_at': item.get('updated_at') or item.get('created_at', ''),
        'author': user or item.get('author', '—'),
        'source': item.get('source', '') if source is None else source,
        'comment': comment,
        'result': copy.deepcopy(item.get('result') or {}),
    }


def _details(item):
    item = copy.deepcopy(item)
    if not item.get('versions'):
        item['versions'] = [_version(item)]
        # Old archives kept at most 20,000 characters; lost text cannot be restored.
        item['source_may_be_truncated'] = len(item.get('source') or '') >= 20000
    return item


def create_report(text, result, user):
    if len(text) > MAX_SOURCE_CHARS:
        raise ValueError('Превышен допустимый размер исходного текста')
    with _LOCK:
        items = _load()
        item = {
            'id': uuid.uuid4().hex[:12], 'created_at': _now(), 'author': user,
            'source': text, 'source_complete': True, 'result': result,
            'approvals': {}, 'rejects': [], 'status': 'pending', 'revision_comment': '', 'revision': 0,
        }
        item['versions'] = [_version(item)]
        items.insert(0, item)
        _save(items)
        return copy.deepcopy(item)


def get_report(rid):
    with _LOCK:
        item = next((it for it in _load() if it.get('id') == rid), None)
        return _details(item) if item else None


def list_reports(limit=20):
    with _LOCK:
        return [_details(it) for it in _load()[:limit]]


def delete_report(rid):
    """Delete one saved report, including its source and all versions."""
    with _LOCK:
        items = _load()
        remaining = [item for item in items if item.get('id') != rid]
        if len(remaining) == len(items):
            return False
        _save(remaining)
        return True


def clear_reports():
    """Clear the entire shared archive, including records beyond the first page."""
    with _LOCK:
        items = _load()
        if items:
            _save([])
        return len(items)


def history_page(limit=20, offset=0):
    def preview(text):
        value = ' '.join(str(text or '').split())
        return value[:180] + ('…' if len(value) > 180 else '')
    with _LOCK:
        all_items = _load()
        items = []
        for item in all_items[offset:offset + limit]:
            result = item.get('result') or {}
            items.append({key: item.get(key, '') for key in ('id', 'created_at', 'updated_at', 'author', 'status')})
            items[-1].update(source_preview=preview(item.get('source')),
                             result_preview=preview(result.get('report_text')),
                             backend=result.get('backend', ''), version_count=len(item.get('versions') or [None]))
        return {'items': items, 'total': len(all_items), 'has_more': offset + len(items) < len(all_items)}


def approve(rid, user):
    with _LOCK:
        items = _load()
        for item in items:
            if item.get('id') != rid:
                continue
            if item.get('status') == 'pending':
                item.setdefault('approvals', {})[user] = True
                if len(item['approvals']) >= MIN_APPROVALS:
                    item['status'] = 'approved'
                item['revision'] = item.get('revision', 0) + 1
                _save(items)
            return _details(item)
    return None


def reject(rid, user, comment):
    if not comment.strip():
        raise ValueError('Укажите комментарий для доработки')
    with _LOCK:
        items = _load()
        for item in items:
            if item.get('id') == rid:
                item.setdefault('rejects', []).append({'user': user, 'comment': comment, 'at': _now()})
                item['status'] = 'revision'
                item['revision_comment'] = comment
                item['revision'] = item.get('revision', 0) + 1
                _save(items)
                return _details(item)
    return None


def regenerate(rid, result, user, request_text, comment='', *, expected_revision=None):
    """Persist the new result and retain every generated version in one write."""
    with _LOCK:
        items = _load()
        for item in items:
            if item.get('id') != rid:
                continue
            if expected_revision is not None and item.get('revision', 0) != expected_revision:
                raise ReportConflict()
            if not item.get('versions'):
                item['versions'] = [_version(item)]
                item['source_may_be_truncated'] = len(item.get('source') or '') >= 20000
            item['result'] = copy.deepcopy(result)
            item['updated_at'] = _now()
            item['versions'].append(_version(item, source=request_text, comment=comment, user=user))
            item['status'] = 'pending'
            item['approvals'] = {}
            item['revision_comment'] = ''
            item['revision'] = item.get('revision', 0) + 1
            _save(items)
            return _details(item)
    return None


def to_pending(rid):
    """Compatibility for existing local callers."""
    with _LOCK:
        items = _load()
        for item in items:
            if item.get('id') == rid:
                item['status'] = 'pending'
                item['approvals'] = {}
                item['revision'] = item.get('revision', 0) + 1
                _save(items)
                return _details(item)
    return None
