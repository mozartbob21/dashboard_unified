"""Self-service module grants against an isolated SQLite DB; no app startup or env."""
import json
from types import SimpleNamespace

import pytest
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient


@pytest.fixture
def api(tmp_path, monkeypatch):
    import utils.db as db
    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'accounts.sqlite')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    from services.auth import accounts
    from routers import users
    from services.auth.registration import ensure_registration_tables
    ensure_registration_tables()
    with db.get_db_connection() as conn:
        for index, username in enumerate(['parent', 'delegate', 'othermanager', 'ordinary'], 1):
            conn.execute('INSERT INTO users(id,username,password_hash,role,modules,is_active,email) VALUES (?,?,?,?,?,?,?)',
                         (index, username, 'synthetic-hash', 'Пользователь', '["edo"]', 1, username+'@example.invalid'))
        conn.execute('UPDATE account_control SET manager_user_id=1,grants_migrated=1 WHERE id=1')
        conn.executemany('INSERT INTO account_managers(user_id) VALUES (?)', [(2,), (3,)])
    actor = {'user': {'id': 2, 'username': 'delegate', 'is_active': True, 'modules': ['edo']}}
    app = FastAPI()
    @app.middleware('http')
    async def synthetic_session(request, call_next):
        request.state.user = actor['user']
        return await call_next(request)
    app.include_router(users.router)
    with TestClient(app) as client:
        yield SimpleNamespace(client=client, actor=actor, db=db, accounts=accounts, routes=users)


def row(api, user_id=2):
    with api.db.get_db_connection() as conn:
        return dict(conn.execute('SELECT * FROM users WHERE id=?', (user_id,)).fetchone())


def update(api, modules):
    return api.client.put('/api/users/me/modules', json={'modules': modules})


def test_delegate_assigns_and_removes_own_modules_without_changing_identity(api):
    before = row(api)
    assert update(api, ['edo', 'zips', 'collective', 'zips']).status_code == 200
    after = row(api)
    assert json.loads(after.pop('modules')) == ['edo', 'zips', 'collective']
    before.pop('modules')
    assert after == before
    assert api.accounts.is_account_manager(api.actor['user'])
    assert update(api, []).status_code == 200
    assert json.loads(row(api)['modules']) == []
    assert api.accounts.is_account_manager(api.actor['user'])


def test_parent_assigns_self_modules_and_remains_protected(api):
    api.actor['user'] = {'id': 1, 'username': 'parent', 'is_active': True}
    assert update(api, ['zips']).status_code == 200
    assert json.loads(row(api, 1)['modules']) == ['zips']
    response = api.client.put('/api/users/1', json={'username': 'parent', 'modules': ['collective']})
    assert response.status_code == 200
    for payload in [{'is_active': False}, {'can_manage_users': False}]:
        assert api.client.put('/api/users/1', json={'username': 'parent', **payload}).status_code == 400
    assert api.client.post('/api/users/1/archive').status_code == 400
    assert api.accounts.is_parent_manager(api.actor['user'])


@pytest.mark.parametrize('extra', [
    {'user_id': 3}, {'id': 3}, {'username': 'othermanager'}, {'email': 'new@example.invalid'},
    {'password': 'Synthetic-new-password'}, {'is_active': False}, {'can_manage_users': True},
])
def test_self_endpoint_rejects_any_non_module_field(api, extra):
    before = row(api)
    assert api.client.put('/api/users/me/modules', json={'modules': ['zips'], **extra}).status_code == 422
    assert row(api) == before
    assert json.loads(row(api, 3)['modules']) == ['edo']


@pytest.mark.parametrize('payload,status', [({}, 422), ({'modules': None}, 422),
    ({'modules': 'zips'}, 422), ({'modules': [False]}, 422), ({'modules': [1]}, 422),
    ({'modules': ['zips']*51}, 422), ({'modules': ['not-a-module']}, 400)])
def test_invalid_module_payload_is_atomic(api, payload, status):
    before = row(api)
    assert api.client.put('/api/users/me/modules', json=payload).status_code == status
    assert row(api) == before


@pytest.mark.parametrize('user_id,username', [(1, 'parent'), (2, 'delegate'), (3, 'othermanager')])
def test_general_editor_does_not_allow_delegate_to_edit_managers(api, user_id, username):
    before = row(api, user_id)
    response = api.client.put('/api/users/'+str(user_id), json={'username': username, 'modules': ['zips']})
    assert response.status_code == 403
    assert row(api, user_id) == before
    assert api.client.post('/api/users/'+str(user_id)+'/archive').status_code in {400, 403}


def test_delegate_still_edits_ordinary_users_but_cannot_grant_management(api):
    assert api.client.put('/api/users/4', json={'username': 'ordinary', 'modules': ['zips']}).status_code == 200
    assert api.client.put('/api/users/4', json={'username': 'ordinary', 'can_manage_users': True}).status_code == 403
    assert json.loads(row(api, 4)['modules']) == ['zips']


@pytest.mark.parametrize('actor', [None, {'id':4,'username':'ordinary','is_active':True},
    {'id':2,'username':'delegate','kc_sub':'synthetic-external-id'}])
def test_non_manager_or_external_identity_cannot_assign_self(api, actor):
    before = row(api)
    api.actor['user'] = actor
    assert update(api, ['zips']).status_code == 403
    assert row(api) == before


@pytest.mark.parametrize('change', ['revoke', 'disable', 'archive'])
def test_service_rechecks_live_privileges_inside_transaction(api, change):
    # This check simulates the route dependency having succeeded before access changed.
    assert api.accounts.is_account_manager(api.actor['user'])
    with api.db.get_db_connection() as conn:
        if change == 'revoke':
            conn.execute('DELETE FROM account_managers WHERE user_id=2')
        elif change == 'disable':
            conn.execute('UPDATE users SET is_active=0 WHERE id=2')
        else:
            conn.execute('INSERT INTO account_archives(user_id) VALUES (2)')
    before = row(api)
    with pytest.raises(HTTPException) as error:
        api.accounts.save_own_modules(['zips'], actor=api.actor['user'])
    assert error.value.status_code == 403
    assert row(api) == before
    assert update(api, ['zips']).status_code == 403


def test_listing_and_update_use_server_resolved_account_not_supplied_id(api):
    api.actor['user']['id'] = 3  # Legacy/stale identity field cannot select another target.
    result = api.client.get('/api/users')
    assert result.status_code == 200
    assert result.json()['current_user_id'] == 2
    assert update(api, ['zips']).status_code == 200
    assert json.loads(row(api)['modules']) == ['zips']
    assert json.loads(row(api, 3)['modules']) == ['edo']
