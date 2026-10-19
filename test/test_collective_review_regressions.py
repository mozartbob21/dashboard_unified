"""Focused review regressions; synthetic data and an isolated database only."""
from datetime import datetime, timezone

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from services.collective import store
from services.collective.client import HEADERS
from routers import collective
from utils import db


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'review.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(store, 'now', lambda: datetime(2026, 10, 19, tzinfo=timezone.utc))


def row(primary='ECUR-1', secondary='SRC-1', source='ЕЦУР'):
    result = [''] * len(HEADERS)
    result[0], result[1], result[2], result[5], result[6], result[8], result[10], result[14] = (
        source, '10.10.2026', 'Тестовый округ', primary, secondary, 'Синтетический текст', '10', 'В работе')
    return result


def save(rows=None):
    return store.save_snapshot('2026-10-01', '2026-10-19', [['1', 'МинЖКХ']], rows or [row()])


def legacy(key='ECUR-1|SRC-1', history=None):
    return {'people': [], 'archived': [], 'since': '2026-10-01', 'items': {key: {
        'report': 'Импортированный доклад', 'report_at': '11.10.2026 10:00',
        'history': history or [],
    }}}


def test_import_after_source_load_keeps_old_report_revisions():
    key = save()['keys'][0]
    store.import_legacy(legacy(history=[{'report': 'Ранний доклад', 'at': '10.10.2026 10:00', 'by': 'Редактор'}]), 'local:1')
    assert store.get_workflow(key)['report'] == 'Импортированный доклад'
    assert any(entry['workflow']['report'] == 'Ранний доклад' for entry in store.workflow_history(key))


def test_secondary_only_import_after_source_load_attaches_to_existing_complaint():
    key = save()['keys'][0]
    store.import_legacy(legacy('|SRC-1'), 'local:1')
    assert store.get_workflow(key)['report'] == 'Импортированный доклад'
    # A later normal refresh must still succeed and preserve the imported data.
    assert save()['keys'] == [key]


def test_secondary_only_legacy_record_cannot_merge_two_distinct_source_ids():
    store.import_legacy(legacy('|SRC-1'), 'local:1')
    first = save([row('ECUR-1', 'SRC-1', 'ЕЦУР')])
    try:
        second = save([row('ECUR-1', 'SRC-1', 'ЕЦУР'), row('MSED-2', 'SRC-1', 'МСЭД')])
    except store.ValidationError:
        assert store.get_snapshot('2026-10-01', '2026-10-19')['keys'] == first['keys']
        return
    assert len(set(second['keys'])) == 2


def test_export_cannot_bypass_server_manual_data_by_omitting_key(monkeypatch):
    key = save()['keys'][0]
    store.update_workflow(key, 0, {'report': 'Действительный доклад'}, 'local:1')
    app = FastAPI()
    @app.middleware('http')
    async def session(request, call_next):
        request.state.user = {'id': 2, 'modules': ['collective']}
        return await call_next(request)
    app.include_router(collective.router)
    seen = []
    monkeypatch.setattr(collective.presentation, 'build', lambda rows, *args: seen.extend(rows) or b'pptx')
    forged = dict(zip(HEADERS, row()))
    forged.update(_theme='ii', _report='Клиентская подмена', _assignee='Клиентский ответственный')
    response = TestClient(app).post('/mingkh/collective/api/pptx', json={
        'start': '2026-10-01', 'end': '2026-10-19', 'rows': [forged],
    })
    assert response.status_code == 400 or (response.status_code == 200 and seen[0]['_report'] == 'Действительный доклад')
