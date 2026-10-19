"""Server persistence tests: synthetic records, real isolated SQLite, no portals."""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path
import json
import threading

import httpx
import pytest

from utils import db
from services.collective import store
from services.collective.client import HEADERS

MIGRATION = Path(__file__).resolve().parents[1] / 'migrations/20261019_collective_workspace.sql'
START, END = '2026-10-01', '2026-10-19'


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'dashboard.db')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(store, 'MIGRATION', MIGRATION)
    monkeypatch.setattr(store, 'now', lambda: datetime(2026, 10, 19, 10, 20, tzinfo=timezone.utc))
    monkeypatch.setattr(httpx.HTTPTransport, 'handle_request', lambda *a, **kw: pytest.fail('Live network is forbidden'))


def record(primary='ECUR-1', secondary='SRC-1', status='В работе', received='10.10.2026', annotation='Синтетическое обращение'):
    row = [''] * len(HEADERS)
    for index, value in {0:'МСЭД', 1:received, 2:'Тестовый округ', 5:primary, 6:secondary, 8:annotation, 10:'123', 14:status}.items():
        row[index] = value
    return row


def snapshot(rows=None, start=START, end=END):
    return store.save_snapshot(start, end, [['curator-test', 'Синтетическое ведомство']], rows if rows is not None else [record()])


def key():
    return snapshot()['keys'][0]


def person(name='Синтетический ответственный'):
    return store.create_assignee(name, 'local:1')


def legacy(**changes):
    return {'people':['Синтетический ответственный'], 'archived':[], 'since':'2026-09-01',
            'items':{'ECUR-1|SRC-1':{'resp':'Синтетический ответственный', 'report':'Доклад из старой программы',
                                  'report_at':'10.10.2026 12:00', 'report_by':'Редактор', 'theme':'ii'}}, **changes}


def test_migration_idempotent_cutoff_is_persisted_across_sessions(monkeypatch):
    first = snapshot()
    assert first['since'] == '2026-10-19'
    monkeypatch.setattr(store, 'now', lambda: datetime(2026, 10, 25, 5, tzinfo=timezone.utc))
    store.initialize()
    store.initialize()
    assert store.get_snapshot(START, END)['since'] == '2026-10-19'
    with db.get_db_connection() as conn:
        assert conn.execute('SELECT count(*) FROM collective_settings').fetchone()[0] == 1
        assert conn.execute("SELECT count(*) FROM app_schema_migrations WHERE name='20261019_collective_workspace'").fetchone()[0] == 1


def test_snapshot_preserves_period_and_manual_fields_across_source_refresh():
    original = snapshot()
    ident = original['keys'][0]
    assignee = person()
    saved = store.update_workflow(ident, 0, {'assignee_id':assignee['id'], 'report':'Ручной доклад', 'theme':'kr'}, 'local:1')
    changed = record(status='Решено', annotation='Обновлённый исходный текст')
    next_period = snapshot([changed], '2026-10-02', END)
    assert next_period['keys'] == [ident]
    assert next_period['workflows'][ident]['report'] == 'Ручной доклад'
    assert next_period['workflows'][ident]['theme'] == 'kr'
    assert next_period['workflows'][ident]['version'] == saved['version']
    old_period = store.get_snapshot(START, END)
    assert old_period['rows'][0][8] == 'Синтетическое обращение'
    assert old_period['workflows'][ident]['report'] == 'Ручной доклад'
    assert old_period['workflows'][ident]['assignee_id'] == assignee['id']
    assert store.latest_snapshot()['keys'] == [ident]


def test_source_refresh_does_not_reset_manual_versions_or_remove_missing_complaints():
    ident = key()
    saved = store.update_workflow(ident, 0, {'report':'Есть доклад'}, 'local:1')
    assert snapshot([])['rows'] == []
    assert store.get_workflow(ident)['report'] == saved['report']
    assert snapshot()['workflows'][ident]['version'] == 1


def test_source_id_appears_later_retains_same_stable_key():
    first = snapshot([record(primary='')])
    ident = first['keys'][0]
    store.update_workflow(ident, 0, {'report':'Сохранить'}, 'local:1')
    second = snapshot([record(primary='NEW-PRIMARY')])
    assert second['keys'] == [ident]
    assert second['workflows'][ident]['report'] == 'Сохранить'


def test_idless_rows_do_not_share_manual_fields_and_mutable_status_keeps_key():
    first = snapshot([record('', '', annotation='Первая'), record('', '', annotation='Вторая')])
    assert len(set(first['keys'])) == 2
    store.update_workflow(first['keys'][0], 0, {'theme':'kr'}, 'local:1')
    second = snapshot([record('', '', status='Решено', annotation='Первая'), record('', '', annotation='Вторая')])
    assert second['keys'] == first['keys']
    assert second['workflows'][first['keys'][1]]['theme'] is None


def test_known_primary_survives_source_number_change_without_merging_other_complaint():
    original = snapshot([record('A', 'A1'), record('B', 'B1')])
    first, second = original['keys']
    store.update_workflow(first, 0, {'report': 'Доклад А'})
    store.update_workflow(second, 0, {'report': 'Доклад Б'})
    changed = snapshot([record('A', 'B1'), record('B', 'B1')])
    assert changed['keys'] == original['keys']
    assert changed['workflows'][first]['report'] == 'Доклад А'
    assert changed['workflows'][second]['report'] == 'Доклад Б'


@pytest.mark.parametrize('status,received,eligible', [('', '01.01.2020', True), ('В работе', '01.01.2020', True),
    ('Ответ в процессе подготовки', '01.01.2020', True), ('Решено', '18.10.2026', False),
    ('Решено', '19.10.2026', True), ('Решено', '2026-10-19', True), ('Решено', '99.99.2026', False)])
def test_cutoff_and_open_status_enforced_on_server(status, received, eligible):
    data = snapshot([record(status=status, received=received)])
    ident = data['keys'][0]
    assert data['workflows'][ident]['can_assign'] is eligible
    assignee = person()
    if eligible:
        assert store.update_workflow(ident, 0, {'assignee_id':assignee['id']})['assignee_id'] == assignee['id']
    else:
        with pytest.raises(store.ValidationError, match='Закрытое'):
            store.update_workflow(ident, 0, {'assignee_id':assignee['id']})
        assert store.update_workflow(ident, 0, {'report':'Доклад разрешён и по закрытому'})['report']


def test_archiving_preserves_assignments_but_prevents_new_assignment_and_restores():
    data = snapshot([record('A', 'A1'), record('B', 'B1')])
    first, second = data['keys']
    assignee = person()
    saved = store.update_workflow(first, 0, {'assignee_id':assignee['id']})
    archived = store.update_assignee(assignee['id'], assignee['version'], {'archived':True})
    assert store.get_workflow(first)['assignee_archived'] is True
    same = store.update_workflow(first, saved['version'], {'assignee_id':assignee['id'], 'report':'Можно обновить'})
    assert same['assignee_id'] == assignee['id']
    with pytest.raises(store.ValidationError, match='архиве'):
        store.update_workflow(second, 0, {'assignee_id':assignee['id']})
    restored = store.update_assignee(assignee['id'], archived['version'], {'active':True})
    assert not restored['archived']
    assert store.update_workflow(second, 0, {'assignee_id':assignee['id']})['assignee_id'] == assignee['id']


def test_existing_assignee_can_be_changed_after_old_complaint_is_closed():
    ident, first, second = key(), person('Первый'), person('Второй')
    store.update_workflow(ident, 0, {'assignee_id':first['id']})
    snapshot([record(status='Решено', received='01.01.2020')])
    assert store.get_workflow(ident)['can_assign']
    changed = store.update_workflow(ident, 1, {'assignee_id':second['id']})
    assert changed['assignee_id'] == second['id']
    cleared = store.update_workflow(ident, 2, {'assignee_id':None})
    assert cleared['assignee_id'] is None and not cleared['can_assign']


def test_directory_name_uniqueness_rename_and_optimistic_version():
    first = person('Иван Иванов')
    with pytest.raises(store.ValidationError, match='уже есть'):
        person('  ИВАН   ИВАНОВ ')
    renamed = store.update_assignee(first['id'], first['version'], {'name':'Иван Петров'})
    with pytest.raises(store.ConflictError) as error:
        store.update_assignee(first['id'], first['version'], {'active':False})
    assert error.value.current == renamed
    assert store.list_assignees()[0]['name'] == 'Иван Петров'


def test_two_concurrent_editors_cannot_overwrite_each_other():
    ident = key()
    barrier = threading.Barrier(2)
    def edit(text):
        barrier.wait()
        try:
            return ('saved', store.update_workflow(ident, 0, {'report':text}, 'local:1'))
        except store.ConflictError as exc:
            return ('conflict', exc.current)
    with ThreadPoolExecutor(max_workers=2) as pool:
        results = list(pool.map(edit, ['Первый текст', 'Второй текст']))
    assert sorted(result[0] for result in results) == ['conflict', 'saved']
    assert results[0][1]['report'] == results[1][1]['report']
    assert store.get_workflow(ident)['version'] == 1


def test_repeat_same_version_returns_conflict_then_current_noop_is_idempotent():
    ident = key()
    saved = store.update_workflow(ident, 0, {'report':'Текст'}, 'local:1')
    with pytest.raises(store.ConflictError):
        store.update_workflow(ident, 0, {'report':'Текст'}, 'local:1')
    assert store.update_workflow(ident, 1, {'report':'Текст'}, 'local:1') == saved
    assert len(store.workflow_history(ident)) == 1


def test_history_is_bounded_and_retains_replaced_reports():
    ident = key()
    for version in range(25):
        store.update_workflow(ident, version, {'report':f'Доклад {version}'}, 'local:1')
    history = store.workflow_history(ident)
    assert len(history) == 20
    assert history[0]['workflow']['report'] == 'Доклад 23'
    assert store.get_workflow(ident)['report'] == 'Доклад 24'


def test_plain_text_and_line_breaks_preserved_without_html_execution_or_logs(capsys):
    ident = key()
    text = '<script>alert(1)</script>\r\n=SUM(1,2)'
    result = store.update_workflow(ident, 0, {'report':text}, 'keycloak:synthetic')
    assert result['report'] == '<script>alert(1)</script>\n=SUM(1,2)'
    assert '<script>' not in capsys.readouterr().out
    assert store.get_snapshot(START, END)['workflows'][ident]['report'] == result['report']


@pytest.mark.parametrize('patch', [{}, {'user_id':1}, {'report':None}, {'report':[]}, {'report':'x'*20001},
    {'report':'bad\x00'}, {'report':'bad\ud800'}, {'theme':''}, {'theme':[]}, {'theme':True},
    {'assignee_id':0}, {'assignee_id':False}, {'assignee_id':'1'}])
def test_invalid_patch_rejected_without_mutation(patch):
    ident = key()
    with pytest.raises(store.ValidationError):
        store.update_workflow(ident, 0, patch)
    assert store.get_workflow(ident)['version'] == 0


@pytest.mark.parametrize('version', [None, False, -1, '0', 1.0])
def test_invalid_versions_rejected(version):
    ident = key()
    with pytest.raises(store.ValidationError):
        store.update_workflow(ident, version, {'report':'Доклад'})


@pytest.mark.parametrize('rows', [None, {}, [[]], [[{}]*16], [[float('nan')]*16], [['x'*200001]*16]])
def test_invalid_source_cannot_replace_previous_snapshot(rows):
    original = snapshot()
    with pytest.raises(store.ValidationError):
        store.save_snapshot(START, END, [], rows)
    assert store.get_snapshot(START, END)['rows'] == original['rows']


@pytest.mark.parametrize('start,end', [(0, END), (START, False), ('2026-10-30', END), (START, '2026-10-20'), ('20261001', END)])
def test_invalid_period_does_not_write_database(start, end):
    with pytest.raises(store.ValidationError):
        store.save_snapshot(start, end, [], [])
    assert not db.DB_FILE.exists()


def test_missing_snapshot_returns_none_and_unknown_workflow_not_found():
    assert store.get_snapshot(START, END) is None
    with pytest.raises(store.NotFoundError):
        store.get_workflow('c_' + 'a'*64)


def test_legacy_import_keeps_existing_manual_data_and_is_idempotent():
    ident = key()
    person('Существующий')
    store.update_workflow(ident, 0, {'report':'Новый серверный доклад', 'theme':'kr'}, 'local:1')
    result = store.import_legacy(legacy(), 'local:1')
    current = store.get_workflow(ident)
    assert result['entries_updated'] == 1 and result['assignees_added'] == 1
    assert current['report'] == 'Новый серверный доклад' and current['theme'] == 'kr'
    assert current['assignee_name'] == 'Синтетический ответственный'
    second = store.import_legacy(legacy(), 'local:1')
    assert second['entries_preserved'] == 1 and second['assignees_added'] == 0
    assert store.get_workflow(ident)['version'] == current['version']
    assert store.get_snapshot(START, END)['since'] == '2026-09-01'


def test_legacy_import_does_not_restore_intentionally_cleared_report_or_assignee():
    ident, assigned = key(), person()
    store.update_workflow(ident, 0, {'report':'Старое', 'assignee_id':assigned['id'], 'theme':'ii'})
    store.update_workflow(ident, 1, {'report':'', 'assignee_id':None, 'theme':None})
    store.import_legacy(legacy())
    result = store.get_workflow(ident)
    assert result['report'] == '' and result['assignee_id'] is None and result['theme'] is None


def test_legacy_import_before_source_load_preserves_history_and_unknown_archived_person():
    data = legacy(people=[], items={'ECUR-1|SRC-1':{'resp':'Бывший сотрудник', 'report':'Текущий',
        'history':[{'report':'Ранее', 'at':'09.10.2026 10:00', 'by':'Редактор'}]}})
    result = store.import_legacy(data)
    assert result['entries_added'] == 1
    fetched = snapshot()
    ident = fetched['keys'][0]
    workflow = fetched['workflows'][ident]
    assert workflow['report'] == 'Текущий' and workflow['assignee_archived'] is True
    assert store.workflow_history(ident)[0]['workflow']['report'] == 'Ранее'
    assert workflow['version'] == 1


def test_legacy_secondary_only_then_primary_arrives_retains_report():
    store.import_legacy(legacy(items={'|SRC-1':{'report':'Только исходный номер'}}))
    fetched = snapshot()
    assert fetched['workflows'][fetched['keys'][0]]['report'] == 'Только исходный номер'


def test_import_cleared_report_keeps_old_history_and_theme_only_import():
    store.import_legacy(legacy(items={'ECUR-1|SRC-1':{'report':'', 'report_at':'10.10.2026 12:00',
                'history':[{'report':'Старый', 'at':'09.10.2026 12:00', 'by':'Редактор'}]}}, themes={'ECUR-1|SRC-1':'kr'}))
    fetched = snapshot()
    ident = fetched['keys'][0]
    assert fetched['workflows'][ident]['report'] == '' and fetched['workflows'][ident]['theme'] == 'kr'
    assert store.workflow_history(ident)[0]['workflow']['report'] == 'Старый'


def test_legacy_invalid_item_aborts_whole_import_before_any_data_change():
    key()
    data = legacy(items={'ECUR-1|SRC-1':{'report':'Валидный'}, 'OTHER|SRC':{'report':'x'*20001}})
    with pytest.raises(store.ValidationError):
        store.import_legacy(data)
    assert store.list_assignees() == []
    assert store.get_snapshot(START, END)['since'] == '2026-10-19'


@pytest.mark.parametrize('data', [legacy(since='2099-01-01'), legacy(since=False), legacy(people='name'),
    legacy(items={'|':{}}), legacy(items={' | ':{}}), legacy(items={'A|B':{'theme':'bad'}}),
    legacy(items={'A|B':{'history':[{}]*21}}), legacy(items={'A|B':{'unknown':'value'}}),
    legacy(items={'A|B':{}, 'a|b':{}}), legacy(people=['x'*121])])
def test_invalid_legacy_shape_rejected(data):
    with pytest.raises(store.ValidationError):
        store.import_legacy(data)
    assert not db.DB_FILE.exists()


def test_two_source_updates_and_edit_keep_human_data():
    ident = key()
    barrier = threading.Barrier(3)
    def operation(index):
        barrier.wait()
        if index == 0:
            return store.update_workflow(ident, 0, {'report':'Переживёт обновления'}, 'local:1')
        return snapshot([record(annotation=f'Источник {index}')])
    with ThreadPoolExecutor(max_workers=3) as pool:
        list(pool.map(operation, range(3)))
    result = store.get_snapshot(START, END)
    assert result['workflows'][ident]['report'] == 'Переживёт обновления'
    assert result['snapshot']['version'] == 3


def test_legacy_history_import_after_source_before_any_human_edit():
    ident = key()
    store.import_legacy(legacy(items={'ECUR-1|SRC-1':{'report':'Текущий', 'history':[
        {'report':'Первый', 'at':'08.10.2026 12:00', 'by':'Редактор'},
        {'report':'Второй', 'at':'09.10.2026 12:00', 'by':'Редактор'}]}}))
    assert [item['workflow']['report'] for item in store.workflow_history(ident)] == ['Второй', 'Первый']
    assert store.get_workflow(ident)['report'] == 'Текущий'
    assert store.get_workflow(ident)['version'] == 2


def test_legacy_history_import_after_manual_theme_preserves_both_histories():
    ident = key()
    store.update_workflow(ident, 0, {'theme':'kr'}, 'local:1')
    store.import_legacy(legacy(items={'ECUR-1|SRC-1':{'report':'Текущий', 'history':[
        {'report':'Старый', 'at':'08.10.2026 12:00', 'by':'Редактор'}]}}))
    current = store.get_workflow(ident)
    assert current['report'] == 'Текущий' and current['theme'] == 'kr'
    assert any(item['workflow']['report'] == 'Старый' for item in store.workflow_history(ident))
    assert len(store.workflow_history(ident)) == 3


def test_secondary_only_legacy_import_after_live_source_uses_existing_workflow():
    ident = key()
    store.import_legacy(legacy(items={'|SRC-1':{'report':'Только исходный номер'}}))
    assert store.get_workflow(ident)['report'] == 'Только исходный номер'
    with db.get_db_connection() as conn:
        assert conn.execute('SELECT count(*) FROM collective_complaints').fetchone()[0] == 1


def test_ambiguous_secondary_only_import_rolls_back_all_changes():
    first, second = record('A', 'SHARED'), record('B', 'SHARED')
    second[0] = 'Другая исходная система'
    snapshot([first, second])
    with pytest.raises(store.ValidationError, match='противоречащие'):
        store.import_legacy(legacy(items={'|SHARED':{'report':'Нельзя угадать'}}))
    assert store.list_assignees() == []
    assert store.get_snapshot(START, END)['since'] == '2026-10-19'
