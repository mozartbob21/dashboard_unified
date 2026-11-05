"""Shared API privacy tests with actual handlers and a synthetic model.

No live model, credentials, portal, or working data are used. The model records
its input so a passing response cannot mask a leak in the supplied history.
"""
import json

import pytest

from services.aichat import private_storage as store
from services.auth.home_preferences import account_key
from test.test_aichat_jobs import (  # Reuse isolated DB / actual route harness.
    USER, OTHER, actual_chat_api, database, wait_done,
)


def shared_dialog():
    dialog = store.create_dialog('Общее обсуждение', owner=account_key(USER))
    store.append_message(dialog['id'], 'user', 'Общий вопрос', owner=account_key(USER))
    store.append_message(dialog['id'], 'assistant', 'PUBLIC_ANSWER', owner=account_key(USER))
    store.append_message(
        dialog['id'], 'assistant', 'SECRET_CDS_REPORT', owner=account_key(USER),
        required_modules=['cds'], file='SECRET_FILE.xlsx', file_name='SECRET_FILE.xlsx',
        report_scope={'modules': ['cds'], 'municipality': 'SECRET_CITY'},
        skill_run={'warnings': ['SECRET_WARNING']},
    )
    return dialog['id']


def send(http, did, request_id, **extra):
    return http.post('/aichat/api/send', data={
        'dialog_id': did, 'text': 'Сократи предыдущий ответ', 'include_context': 'false',
        'background': 'true', 'request_id': request_id, **extra,
    })


def test_shared_followup_retains_source_rights_on_job_and_dialog(actual_chat_api):
    http, identity, model, calls = actual_chat_api
    did = shared_dialog()
    model.return_value = 'SECRET_SHORT_SUMMARY'
    response = send(http, did, 'shared_high_followup_12345')
    assert response.status_code == 202
    done = wait_done(response.json()['job_id'])
    assert done['status'] == 'succeeded'
    assert done['result']['required_modules'] == ['cds']
    assert 'SECRET_CDS_REPORT' in repr(model.call_args)
    assert not calls

    identity['user'] = dict(OTHER)
    restricted = http.get('/aichat/api/dialogs/' + did)
    assert restricted.status_code == 200
    assert 'SECRET_' not in restricted.text
    assert 'PUBLIC_ANSWER' in restricted.text
    assert restricted.json()['messages'][-1]['restricted'] is True
    assert http.get('/aichat/api/jobs/' + done['job_id']).status_code == 403
    assert done['job_id'] not in http.get('/aichat/api/jobs').text
    assert 'SECRET_' not in http.get('/aichat/api/dialogs').text


@pytest.mark.parametrize('transport', ['background', 'legacy'])
@pytest.mark.parametrize('set_team', [False, True])
def test_low_rights_followup_never_supplies_hidden_body_or_metadata_to_model(
        actual_chat_api, transport, set_team):
    http, identity, model, calls = actual_chat_api
    did = shared_dialog()
    identity['user'] = dict(OTHER)
    model.return_value = 'PUBLIC_NEW_ANSWER'
    options = {'team': 'null'} if set_team else {}
    if transport == 'legacy':
        options['background'] = 'false'
    response = send(http, did, 'shared_low_followup_12345', **options)
    if transport == 'background':
        assert response.status_code == 202
        result = wait_done(response.json()['job_id'])['result']
    else:
        assert response.status_code == 200
        result = response.json()
    assert model.called and not calls
    assert 'SECRET_' not in repr(model.call_args)
    assert 'PUBLIC_ANSWER' in repr(model.call_args)
    assert result['answer'] == 'PUBLIC_NEW_ANSWER'
    assert result['required_modules'] == []
    after = http.get('/aichat/api/dialogs/' + did).json()
    assert 'SECRET_' not in json.dumps(after, ensure_ascii=False)
    assert after['messages'][-1]['content'] == 'PUBLIC_NEW_ANSWER'
    assert not after['messages'][-1].get('restricted')


def test_shared_team_patch_redacts_history_without_erasing_it(actual_chat_api):
    http, identity, _, _ = actual_chat_api
    did = shared_dialog()
    identity['user'] = dict(OTHER)
    changed = http.patch('/aichat/api/dialogs/' + did + '/team', json={'team': None})
    assert changed.status_code == 200
    assert 'SECRET_' not in changed.text
    assert changed.json()['messages'][-1]['restricted'] is True
    identity['user'] = dict(USER)
    still_present = http.get('/aichat/api/dialogs/' + did)
    assert 'SECRET_CDS_REPORT' in still_present.text
    assert 'SECRET_FILE.xlsx' in still_present.text


def test_shared_list_title_is_filtered_with_first_question(actual_chat_api):
    http, identity, _, _ = actual_chat_api
    dialog = store.create_dialog('SECRET_TITLE', owner=account_key(USER))
    store.append_message(dialog['id'], 'user', 'SECRET_QUERY', owner=account_key(USER), required_modules=['cds'])
    store.append_message(dialog['id'], 'assistant', 'SECRET_ANSWER', owner=account_key(USER), required_modules=['cds'])
    identity['user'] = dict(OTHER)
    listing = http.get('/aichat/api/dialogs')
    assert listing.status_code == 200
    assert dialog['id'] in listing.text
    assert 'SECRET_' not in listing.text
    assert 'SECRET_' not in http.get('/aichat/api/dialogs/' + dialog['id']).text


def test_revocation_during_report_preparation_cannot_reuse_old_history(actual_chat_api, monkeypatch):
    from services.aichat import block_execution, report_context
    from utils import db
    http, _, model, _ = actual_chat_api
    did = shared_dialog()

    def revoke_after_history_was_selected(text, user):
        # Simulate a separate administrator changing grants between the history
        # read and the context collector refreshing the executing identity.
        with db.get_db_connection() as conn:
            conn.execute('UPDATE users SET modules=? WHERE id=?', (json.dumps(['edo']), USER['id']))
        return None

    monkeypatch.setattr(block_execution, 'recognize', revoke_after_history_was_selected)
    monkeypatch.setattr(report_context, 'prepare', lambda *args, **kwargs: {
        'selection': {'modules': ['edo']}, 'sources': [],
    })
    response = send(http, did, 'shared_revoked_followup_12345', include_context='true')
    assert response.status_code == 202
    done = wait_done(response.json()['job_id'])
    assert done['status'] == 'failed'
    assert not model.called


def test_shared_delete_checks_all_message_rights_on_server(actual_chat_api):
    import ast
    from pathlib import Path
    from fastapi import HTTPException, Request
    http, identity, _, _ = actual_chat_api
    source = Path(__file__).resolve().parents[1] / 'app.py'
    node = next(node for node in ast.parse(source.read_text(encoding='utf-8-sig')).body
                if isinstance(node, ast.AsyncFunctionDef) and node.name == 'aichat_del')
    exec(compile(ast.Module(body=[node], type_ignores=[]), str(source), 'exec'), {
        'app': http.app, 'Request': Request, 'HTTPException': HTTPException,
    })
    did = shared_dialog()
    identity['user'] = dict(OTHER)
    denied = http.delete('/aichat/api/dialogs/' + did)
    assert denied.status_code == 403
    assert 'SECRET_' not in denied.text
    assert store.get_dialog(did, account_key(USER)) is not None

    # Shared ownership stays shared: a peer may delete an ordinary chat.
    public = store.create_dialog('Общая тема', owner=account_key(USER))
    store.append_message(public['id'], 'user', 'Общий вопрос', owner=account_key(USER))
    assert http.delete('/aichat/api/dialogs/' + public['id']).json() == {'ok': True}

    identity['user'] = dict(USER)
    assert http.delete('/aichat/api/dialogs/' + did).json() == {'ok': True}
    assert http.get('/aichat/api/dialogs/' + did).status_code == 404
