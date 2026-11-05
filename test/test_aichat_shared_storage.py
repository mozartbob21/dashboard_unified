"""Shared chat boundary tests: real temporary SQLite, no working data or network."""
from concurrent.futures import ThreadPoolExecutor
import copy
import json
from pathlib import Path

from fastapi import HTTPException
import pytest

from core.roles import ALL_MODULE_IDS
from services.aichat import private_storage as store
from services.aichat.engine import GREETING
from utils import db

EDITOR = {'id': 1, 'username': 'first', 'role': 'user', 'modules': ['edo']}
PEER = {'id': 2, 'username': 'second', 'role': 'user', 'modules': ['edo']}
RESTRICTED = {'id': 3, 'username': 'third', 'role': 'Администратор', 'modules': []}


@pytest.fixture
def temporary_store(tmp_path, monkeypatch):
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'shared.sqlite')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    monkeypatch.setattr(store, 'LEGACY_FILE', tmp_path / 'dialogs.json')
    with db.get_db_connection() as conn:
        for user in (EDITOR, PEER, RESTRICTED):
            conn.execute('INSERT INTO users(id,username,password_hash,role,modules) VALUES (?,?,?,?,?)',
                         (user['id'], user['username'], 'unused', user['role'], json.dumps(user['modules'])))
    return tmp_path


def write_legacy(items):
    payload = json.dumps(items, ensure_ascii=False).encode()
    store.LEGACY_FILE.write_bytes(payload)
    return payload


def test_accounts_share_dialog_and_changes_without_owner_claim(temporary_store):
    dialog = store.create_dialog('Общая тема', owner='local:1')
    store.append_message(dialog['id'], 'user', 'Общий вопрос', owner='local:1')
    store.append_message(dialog['id'], 'assistant', 'Общий ответ', owner='local:2')
    assert store.get_dialog(dialog['id'], 'local:2')['messages'][-1]['content'] == 'Общий ответ'
    assert dialog['id'] in {item['id'] for item in store.list_dialogs('local:3')}
    # Use the server's own normalizer, without trusting custom role definitions.
    updated = store.set_team(dialog['id'], {'preset': 'council'}, owner='local:2')
    assert updated['team']['preset'] == 'council'
    assert store.get_dialog(dialog['id'], 'local:1')['team'] == updated['team']


def test_restricted_message_redacts_all_payload_metadata_without_mutating_raw(temporary_store):
    raw = {'role': 'assistant', 'content': 'SECRET_CONTENT', 'required_modules': ['edo'],
           'file': 'SECRET_FILE.xlsx', 'file_name': 'SECRET_FILE.xlsx', 'ts': 12, 'event_id': 'event',
           'report_scope': {'modules': ['edo'], 'municipalities': ['SECRET_CITY']},
           'skill_run': {'warnings': ['SECRET_WARNING']}, 'evidence': {'secret': 'SECRET_EVIDENCE'},
           'action': {'params': {'secret': 'SECRET_PARAM'}}}
    before = copy.deepcopy(raw)
    hidden = store.visible_message(RESTRICTED, raw)
    assert hidden.get('restricted') is True
    assert 'SECRET_' not in json.dumps(hidden, ensure_ascii=False)
    assert hidden.get('role') == 'assistant'
    assert store.visible_message(PEER, raw)['content'] == 'SECRET_CONTENT'
    assert raw == before


def test_explicit_empty_requirements_do_not_override_known_module_scope(temporary_store):
    message = {'role': 'assistant', 'content': 'SECRET', 'required_modules': [], 'module_id': 'edo'}
    assert not store.can_view(RESTRICTED, message)
    assert store.can_view(PEER, message)


@pytest.mark.parametrize('scope', [
    {'required_modules': ['mingkh-water-map']},
    {'module_id': 'mingkh-water-map'},
    {'report_scope': {'modules': ['mingkh-water-map']}},
])
def test_aliases_share_the_same_module_permission(temporary_store, scope):
    message = {'role': 'assistant', 'content': 'Map data', **scope}
    assert set(store.required_modules(message)) == {'mingkh'}
    assert store.can_view({'id': 4, 'modules': ['mingkh']}, message)
    assert not store.can_view(RESTRICTED, message)


@pytest.mark.parametrize('wildcard', ['all', '*'])
def test_legacy_wildcard_never_means_public_data(temporary_store, wildcard):
    message = {'role': 'assistant', 'content': 'SECRET', 'report_scope': {'module': wildcard}}
    assert set(store.required_modules(message)) == set(ALL_MODULE_IDS)
    assert not store.can_view(PEER, message)
    assert store.can_view({'id': 4, 'modules': list(ALL_MODULE_IDS)}, message)


def test_dialog_title_and_messages_use_same_rights(temporary_store):
    raw = {'id': 'example', 'title': 'SECRET_TITLE', 'messages': [
        {'role': 'user', 'content': 'SECRET_QUERY', 'required_modules': ['edo']},
        {'role': 'assistant', 'content': 'SECRET_ANSWER', 'required_modules': ['edo']},
    ]}
    hidden = store.visible_dialog(RESTRICTED, raw)
    assert 'SECRET_' not in json.dumps(hidden, ensure_ascii=False)
    assert store.visible_dialog(PEER, raw)['title'] == 'SECRET_TITLE'


def test_legacy_import_preserves_file_and_propagates_known_scope(temporary_store):
    original = write_legacy([{'id': 'legacy-id', 'title': 'Старый общий диалог', 'messages': [
        {'role': 'assistant', 'content': GREETING},
        {'role': 'user', 'content': 'Отчёт', 'report_scope': {'modules': ['edo']}},
        {'role': 'assistant', 'content': 'SECRET_FIRST_REPORT'},
        {'role': 'user', 'content': 'Сократи прежний ответ'},
        {'role': 'assistant', 'content': 'SECRET_FOLLOWUP'},
    ]}])
    raw = store.get_dialog('legacy-id', 'local:2')
    assert raw is not None
    hidden = store.visible_dialog(RESTRICTED, raw)
    assert hidden['messages'][0]['content'] == GREETING
    assert 'SECRET_' not in json.dumps(hidden, ensure_ascii=False)
    visible = store.visible_dialog(PEER, raw)
    assert visible['messages'][-1]['content'] == 'SECRET_FOLLOWUP'
    assert store.LEGACY_FILE.read_bytes() == original
    store.initialize()
    assert len(store.list_dialogs('local:3')) == 1


def test_unknown_legacy_answer_is_not_mislabeled_public(temporary_store):
    write_legacy([{'id': 'unknown', 'title': 'История без происхождения ответа', 'messages': [
        {'role': 'user', 'content': 'Вопрос'}, {'role': 'assistant', 'content': 'UNKNOWN_LEGACY_DATA'}]}])
    raw = store.get_dialog('unknown', 'local:1')
    assert raw['messages'][-1]['content'] == 'UNKNOWN_LEGACY_DATA'
    assert 'UNKNOWN_LEGACY_DATA' not in json.dumps(store.visible_dialog(RESTRICTED, raw))
    assert store.visible_dialog({'id': 4, 'modules': list(ALL_MODULE_IDS)}, raw)['messages'][-1]['content'] == 'UNKNOWN_LEGACY_DATA'


@pytest.mark.parametrize('known_scope', [{'module_id': 'edo'}, {'report_scope': {'modules': ['edo']}}])
def test_legacy_answer_with_own_known_scope_does_not_require_unrelated_modules(temporary_store, known_scope):
    original = write_legacy([{'id': 'own-scope', 'title': 'Отчёт', 'messages': [
        {'role': 'user', 'content': 'Покажи результат'},
        {'role': 'assistant', 'content': 'KNOWN_SOURCE_REPORT', **known_scope},
    ]}])
    raw = store.get_dialog('own-scope', 'local:2')
    assert raw['messages'][-1]['required_modules'] == ['edo']
    assert store.visible_dialog(PEER, raw)['messages'][-1]['content'] == 'KNOWN_SOURCE_REPORT'
    assert store.visible_dialog(RESTRICTED, raw)['messages'][-1]['restricted'] is True
    assert store.LEGACY_FILE.read_bytes() == original


def test_legacy_import_is_concurrent_idempotent_and_deletion_does_not_resurrect(temporary_store):
    write_legacy([{'id': 'once', 'title': 'Общая история', 'messages': []}])
    with ThreadPoolExecutor(max_workers=6) as pool:
        list(pool.map(lambda _: store.initialize(), range(6)))
    assert [item['id'] for item in store.list_dialogs('local:1')] == ['once']
    assert store.delete_dialog('once', owner='local:2')
    store.initialize()
    assert store.get_dialog('once', 'local:1') is None
    assert store.LEGACY_FILE.exists()


def test_any_accounts_active_task_blocks_shared_dialog_deletion(temporary_store):
    from services.aichat import jobs
    dialog = store.create_dialog('Выполняемая задача', owner='local:1')
    ticket, _ = jobs.submit(EDITOR, 'shared_delete_guard_1234', {'text': 'test'}, dialog_id=dialog['id'])
    with pytest.raises(HTTPException) as error:
        store.delete_dialog(dialog['id'], owner='local:2')
    assert error.value.status_code == 409
    assert store.get_dialog(dialog['id'], 'local:2') is not None
    jobs._finish(ticket['job_id'], 'failed', {'ok': False, 'answer': 'Synthetic stopped', 'required_modules': []})
    assert store.delete_dialog(dialog['id'], owner='local:2')
