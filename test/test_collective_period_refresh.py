"""Period expansion must not merge different complaints sharing a source number."""
from datetime import datetime, timezone
import logging

import pytest

from services.collective import store, refresh
from services.collective.client import HEADERS
from utils import db


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'period.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(store, 'now', lambda: datetime(2026, 10, 19, tzinfo=timezone.utc))
    monkeypatch.setattr(refresh, 'kick', lambda: None)
    monkeypatch.setattr(refresh, 'credentials', lambda key: {'username': 'synthetic', 'password': 'do-not-log'})


def row(primary, secondary='SRC-SHARED', source='ЕЦУР', annotation='Синтетический текст'):
    data = [''] * len(HEADERS)
    data[0], data[1], data[2], data[5], data[6], data[8], data[10], data[14] = (
        source, '15.10.2026', 'Тестовый округ', primary, secondary, annotation, '10', 'В работе')
    return data


def save(rows, start='2026-10-13'):
    return store.save_snapshot(start, '2026-10-19', [['1', 'МинЖКХ']], rows)


def test_month_expansion_keeps_distinct_primary_ids_sharing_secondary():
    week = save([row('ECUR-1')])
    key = week['keys'][0]
    assigned = store.create_assignee('Ответственный', 'local:1')
    store.update_workflow(key, 0, {'report': 'Сохранить доклад', 'assignee_id': assigned['id'], 'theme': 'kr'}, 'local:1')
    month = save([row('ECUR-1'), row('ECUR-2')], '2026-10-01')
    assert len(set(month['keys'])) == 2
    assert month['keys'][0] == key
    assert month['workflows'][key]['report'] == 'Сохранить доклад'
    assert month['workflows'][key]['assignee_id'] == assigned['id']
    assert month['workflows'][month['keys'][1]]['report'] == ''
    assert store.get_snapshot('2026-10-13', '2026-10-19')['rows'] == week['rows']
    repeated = save([row('ECUR-2'), row('ECUR-1')], '2026-10-01')
    assert repeated['keys'] == month['keys'][::-1]


def test_new_shared_source_numbers_work_independent_of_first_load_order():
    first = save([row('ECUR-2'), row('ECUR-1')], '2026-10-01')
    assert len(set(first['keys'])) == 2
    second = save([row('ECUR-1'), row('ECUR-2')])
    assert second['keys'] == first['keys'][::-1]


def test_later_missing_primary_does_not_select_first_of_ambiguous_secondary_ids():
    month = save([row('ECUR-1'), row('ECUR-2')], '2026-10-01')
    store.update_workflow(month['keys'][0], 0, {'report': 'Первый доклад'}, 'local:1')
    with pytest.raises(store.ValidationError, match='номер|сопостав'):
        save([row('')], '2026-10-01')
    assert store.get_snapshot('2026-10-01', '2026-10-19')['keys'] == month['keys']


@pytest.mark.parametrize('reverse', [False, True])
def test_month_cannot_attach_unnumbered_legacy_report_to_arbitrary_first_row(reverse):
    store.import_legacy({'people': [], 'archived': [], 'since': '2026-10-01', 'items': {
        '|SRC-SHARED': {'report': 'Не потерять старый доклад'},
    }}, 'local:1')
    with db.get_db_connection() as conn:
        original = conn.execute('SELECT complaint_key FROM collective_workflows').fetchone()[0]
    incoming = [row('ECUR-1'), row('ECUR-2')]
    if reverse:
        incoming.reverse()
    with pytest.raises(store.ValidationError, match='однозначно'):
        save(incoming, '2026-10-01')
    assert store.get_workflow(original)['report'] == 'Не потерять старый доклад'
    assert store.get_snapshot('2026-10-01', '2026-10-19') is None
    with db.get_db_connection() as conn:
        assert conn.execute('SELECT count(*) FROM collective_complaints').fetchone()[0] == 1


def test_real_strong_alias_conflict_still_aborts_atomically():
    original = save([row('ECUR-1', 'SRC-1'), row('ECUR-2', 'SRC-2')])
    with db.get_db_connection() as conn:
        conn.execute('INSERT INTO collective_identity_aliases(alias, complaint_key) VALUES (?, ?)',
                     ('pair:["ecur-1","src-2"]', original['keys'][1]))
    with pytest.raises(store.ValidationError, match='однозначно'):
        save([row('ECUR-1', 'SRC-2')])
    assert store.get_snapshot('2026-10-13', '2026-10-19')['keys'] == original['keys']


def test_refresh_displays_safe_validation_reason_and_logs_no_content(monkeypatch, caplog):
    original = save([row('ECUR-1')])
    secret = 'do-not-log-source-or-report'
    store.update_workflow(original['keys'][0], 0, {'report': secret}, 'local:1')
    monkeypatch.setattr(refresh.CollectivePortal, 'fetch', lambda *a: ([['1', 'МинЖКХ']], [[secret]]))
    ticket = refresh.submit('2026-10-01', '2026-10-19', 'synthetic-error-key-01')
    with caplog.at_level(logging.WARNING):
        refresh._drain()
    status = refresh.status(ticket['job_id'])
    assert status['status'] == 'error'
    assert 'Структура обращений источника изменилась' in status['error']
    assert store.get_workflow(original['keys'][0])['report'] == secret
    assert 'ValidationError' in caplog.text
    assert 'COLLECTIVE_VALIDATION' in caplog.text
    assert secret not in caplog.text and 'do-not-log' not in caplog.text


def test_unexpected_exception_never_exposes_message_to_logs_or_ui(monkeypatch, caplog):
    def fail(*args):
        raise RuntimeError('private-source-and-password')
    monkeypatch.setattr(refresh.CollectivePortal, 'fetch', fail)
    ticket = refresh.submit('2026-10-01', '2026-10-19', 'synthetic-error-key-02')
    with caplog.at_level(logging.WARNING):
        refresh._drain()
    status = refresh.status(ticket['job_id'])
    assert status['status'] == 'error'
    assert 'RuntimeError' in caplog.text and 'COLLECTIVE_UNEXPECTED' in caplog.text
    assert 'private-source-and-password' not in caplog.text + status['error']
