"""Repair an intermediate workspace schema without rewriting stored human work."""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

import pytest

from services.collective import store
from services.collective.client import HEADERS
from utils import db


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'upgrade.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(store, 'now', lambda: datetime(2026, 10, 19, 10, 20, tzinfo=timezone.utc))
    root = Path(__file__).resolve().parents[1]
    monkeypatch.setattr(store, 'MIGRATION', root / 'migrations/20261019_collective_workspace.sql')
    monkeypatch.setattr(store, 'SECONDARY_MIGRATION', root / 'migrations/20261019_collective_secondary_index.sql')


def source(secondary='SRC-OLD'):
    row = [''] * len(HEADERS)
    row[0], row[1], row[2], row[5], row[6], row[8], row[10], row[14] = (
        'ЕЦУР', '19.10.2026', 'Тестовый округ', 'ECUR-1', secondary, 'Обращение', '123', 'В работе')
    return row


BUSINESS_TABLES = ('collective_settings', 'collective_complaints', 'collective_identity_aliases',
                   'collective_assignees', 'collective_workflows', 'collective_workflow_history',
                   'collective_snapshots')


def business_state():
    with db.get_db_connection() as conn:
        return {name: [tuple(row) for row in conn.execute('SELECT * FROM ' + name + ' ORDER BY rowid')]
                for name in BUSINESS_TABLES}


def prepare_intermediate(*, keep_table=False, keep_marker=False):
    week = store.save_snapshot('2026-10-13', '2026-10-19', [['1', 'МинЖКХ']], [source()])
    key = week['keys'][0]
    person = store.create_assignee('Ответственный', 'local:1')
    store.update_workflow(key, 0, {'report': 'Первый доклад', 'theme': 'kr', 'assignee_id': person['id']}, 'local:1')
    store.update_workflow(key, 1, {'report': 'Последний сохранённый доклад'}, 'local:2')
    # Later source data changes the source number, while the old alias remains.
    store.save_snapshot('2026-10-01', '2026-10-19', [['1', 'МинЖКХ']], [source('SRC-NEW')])
    before = business_state()
    with db.get_db_connection() as conn:
        assert conn.execute("SELECT 1 FROM app_schema_migrations WHERE name='20261019_collective_workspace'").fetchone()
        if not keep_table:
            conn.execute('DROP TABLE collective_secondary_ids')
        if not keep_marker:
            conn.execute('DELETE FROM app_schema_migrations WHERE name=?', (store.SECONDARY_MIGRATION_ID,))
    return key, person, before


def test_upgrade_keeps_old_week_and_manual_data_then_loads_today():
    key, person, before = prepare_intermediate()
    # Reading the previous period repairs a derivative missing from an
    # intermediate installation; it does not require querying the portal.
    week = store.get_snapshot('2026-10-13', '2026-10-19')
    assert week['keys'] == [key]
    assert week['rows'][0][6] == 'SRC-OLD'
    assert week['workflows'][key]['report'] == 'Последний сохранённый доклад'
    assert week['workflows'][key]['assignee_id'] == person['id']
    assert week['workflows'][key]['theme'] == 'kr'
    assert business_state() == before
    with db.get_db_connection() as conn:
        ids = {tuple(row) for row in conn.execute('SELECT * FROM collective_secondary_ids')}
        assert ids == {('src-old', key), ('src-new', key)}
        assert conn.execute('SELECT count(*) FROM app_schema_migrations WHERE name=?', (store.SECONDARY_MIGRATION_ID,)).fetchone()[0] == 1
    today = store.save_snapshot('2026-10-19', '2026-10-19', [['1', 'МинЖКХ']], [source('SRC-NEW')])
    assert today['keys'] == [key]
    assert today['workflows'][key]['report'] == 'Последний сохранённый доклад'
    assert store.get_snapshot('2026-10-13', '2026-10-19')['rows'] == week['rows']
    after = business_state()
    store.initialize()
    store.initialize()
    assert business_state() == after
    assert any(item['workflow']['report'] == 'Первый доклад' for item in store.workflow_history(key))


def test_complete_database_is_unchanged_except_forward_marker():
    _, _, before = prepare_intermediate(keep_table=True)
    with db.get_db_connection() as conn:
        secondary_before = [tuple(row) for row in conn.execute('SELECT * FROM collective_secondary_ids ORDER BY 1,2')]
    store.initialize()
    store.initialize()
    assert business_state() == before
    with db.get_db_connection() as conn:
        assert [tuple(row) for row in conn.execute('SELECT * FROM collective_secondary_ids ORDER BY 1,2')] == secondary_before


def test_missing_table_can_be_repaired_even_after_forward_marker_exists():
    key, _, before = prepare_intermediate(keep_marker=True)
    store.initialize()
    assert business_state() == before
    with db.get_db_connection() as conn:
        assert conn.execute('SELECT count(*) FROM collective_secondary_ids WHERE complaint_key=?', (key,)).fetchone()[0] == 2


def test_failed_backfill_rolls_back_table_and_marker_without_losing_source_or_reports(monkeypatch):
    key, _, before = prepare_intermediate()
    original = store._secondary_index_entries
    def fail_after_one(conn):
        yield 'synthetic-partial', key
        raise RuntimeError('synthetic failure')
    monkeypatch.setattr(store, '_secondary_index_entries', fail_after_one)
    with pytest.raises(RuntimeError, match='synthetic failure'):
        store.initialize()
    assert business_state() == before
    with db.get_db_connection() as conn:
        assert not conn.execute("SELECT 1 FROM sqlite_master WHERE name='collective_secondary_ids'").fetchone()
        assert not conn.execute('SELECT 1 FROM app_schema_migrations WHERE name=?', (store.SECONDARY_MIGRATION_ID,)).fetchone()
    monkeypatch.setattr(store, '_secondary_index_entries', original)
    store.initialize()
    assert business_state() == before


def test_concurrent_initializers_preserve_every_record():
    _, _, before = prepare_intermediate()
    with ThreadPoolExecutor(max_workers=4) as pool:
        list(pool.map(lambda _: store.initialize(), range(8)))
    assert business_state() == before
    with db.get_db_connection() as conn:
        assert conn.execute('SELECT count(*) FROM app_schema_migrations WHERE name=?', (store.SECONDARY_MIGRATION_ID,)).fetchone()[0] == 1

