"""No live MAX/source calls: real temporary SQLite + synthetic source/transport."""
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
import hashlib
from pathlib import Path
import sqlite3
import threading
from types import SimpleNamespace

import pytest

from services.max_bot.client import MaxAPIError
from services.max_bot.presentations import PresentationService
from services.max_bot.store import Store, StoreError

NOW = 1_800_000_000


@pytest.fixture
def env(tmp_path):
    accounts = tmp_path / "accounts.sqlite"
    with sqlite3.connect(accounts) as db:
        db.executescript("CREATE TABLE users(id INTEGER PRIMARY KEY,username TEXT,role TEXT,modules TEXT,is_active INTEGER,password_hash TEXT);"
            "CREATE TABLE account_control(id INTEGER PRIMARY KEY,manager_user_id INTEGER);"
            "CREATE TABLE account_managers(user_id INTEGER PRIMARY KEY);INSERT INTO account_control VALUES(1,9);")
        db.executemany("INSERT INTO users VALUES(?,?,?,'[\"presentations\"]',1,?)", [(1,"one","Пользователь","hash-one"),(2,"two","Пользователь","hash-two")])
    store = Store(SimpleNamespace(db_path=tmp_path / "bot.sqlite", accounts_path=accounts, link_ttl=300, binding_ttl=2592000))
    with store._connection() as db:
        db.executemany("INSERT INTO max_bot_bindings VALUES(?,?,?,?,?,?)", [
            (1,"101","201",hashlib.sha256(b'hash-one').hexdigest(),NOW,NOW+2592000),
            (2,"102","202",hashlib.sha256(b'hash-two').hexdigest(),NOW,NOW+2592000)])
    data = SimpleNamespace(store=store, now=NOW, items=[], sent=[], uploaded=[], failures={}, calls=0,
                           source_error=None, after_upload=None, list_hook=None, downloads=[])
    class Client:
        def upload_file(self, path):
            assert path.stat().st_mode & 0o777 == 0o600
            assert path.parent.stat().st_mode & 0o777 == 0o700
            data.uploaded.append(path.name)
            if data.after_upload:
                data.after_upload()
            return path.stem
        def send_file(self, max_id, token, caption):
            assert type(max_id) is int
            failure = data.failures.get(token)
            if failure:
                raise failure
            data.sent.append((max_id, token, caption))
            return True
    class Source:
        def list_attachments(self):
            data.calls += 1
            if data.list_hook:
                data.list_hook()
            if data.source_error:
                raise data.source_error
            return list(data.items)
        def download(self, item, directory):
            path = directory / (item.id + '.pptx')
            path.write_bytes(b'synthetic-file')
            data.downloads.append(path)
            return path
    @contextmanager
    def factory():
        yield Source()
    data.factory, data.client = factory, Client()
    data.service = PresentationService(store, source_factory=factory, client=data.client, clock=lambda:data.now, waiter=lambda _:False)
    return data


def add(env, ident):
    env.items.append(SimpleNamespace(id=str(ident),name='Presentation '+str(ident)+'.pptx',description='Synthetic',created_at='2026-10-23',path=''))


def subscribe(env, who=1, enabled=True):
    return env.service.set_subscription({'id':who}, str(100+who), str(200+who), enabled)


def advance(env, seconds=3601):
    env.now += seconds
    return env.service.tick()


def states(env):
    with env.store._connection() as db:
        return [(r['attachment_id'], r['user_id'], r['state']) for r in db.execute('SELECT * FROM max_presentations_deliveries ORDER BY id')]


def test_first_scan_is_baseline_and_only_new_files_are_sent(env):
    add(env, 1)
    subscribe(env)
    assert env.service.tick()['scanned']
    assert env.sent == []
    assert env.service.status(1)['baseline_ready']
    add(env, 2)
    assert advance(env)['sent'] == 1
    assert [s[1] for s in env.sent] == ['2']
    assert states(env) == [('2', 1, 'sent')]
    assert advance(env)['sent'] == 0
    assert all(not p.exists() and not p.parent.exists() for p in env.downloads)


def test_subscribe_is_personal_explicit_and_new_subscriber_gets_no_backlog(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    advance(env)
    subscribe(env, 2)
    subscribe(env, 2)  # Idempotent button/network retry.
    advance(env)
    assert [(s[0],s[1]) for s in env.sent] == [(101,'1')]
    add(env, 2)
    advance(env)
    assert [(s[0],s[1]) for s in env.sent] == [(101,'1'),(101,'2'),(102,'2')]
    assert '102' not in repr(env.service.status(1))


def test_partial_failure_does_not_replay_success_and_retries_definite_failure(env):
    subscribe(env)
    env.service.tick()
    add(env, 1); add(env, 2)
    env.failures['2'] = MaxAPIError('MAX_ATTACHMENT_NOT_READY','safe',retryable=True)
    advance(env)
    assert states(env) == [('1',1,'sent'),('2',1,'retry')]
    del env.failures['2']
    advance(env, 300)
    assert [s[1] for s in env.sent] == ['1','2']
    assert env.calls == 2  # Retry deliveries without re-listing the source.


def test_ambiguous_send_is_not_automatically_retried_even_after_restart(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    env.failures['1'] = MaxAPIError('MAX_READ_TIMEOUT','safe',uncertain=True)
    advance(env)
    assert states(env) == [('1',1,'uncertain')]
    assert env.service.status(1)['uncertain'] == 1
    env.failures.clear()
    env.service = PresentationService(env.store, source_factory=env.factory, client=env.client, clock=lambda:env.now, waiter=lambda _:False)
    advance(env)
    assert env.sent == []
    assert env.uploaded == ['1.pptx']


def test_crash_recovery_upload_retries_but_send_is_uncertain(env):
    subscribe(env)
    env.service.tick()
    add(env, 1); add(env, 2)
    env.failures.update({str(i):MaxAPIError('MAX_ATTACHMENT_NOT_READY','safe',retryable=True) for i in (1,2)})
    advance(env)
    with env.store._connection() as db:
        db.execute("UPDATE max_presentations_deliveries SET state='uploading' WHERE attachment_id='1'")
        db.execute("UPDATE max_presentations_deliveries SET state='sending' WHERE attachment_id='2'")
    env.failures.clear()
    advance(env, 300)
    assert states(env) == [('1',1,'sent'),('2',1,'uncertain')]
    assert [s[1] for s in env.sent] == ['1']


def test_two_service_instances_do_not_scan_or_deliver_concurrently(env):
    subscribe(env)
    entered, release = threading.Event(), threading.Event()
    def block():
        entered.set()
        assert release.wait(5)
    env.list_hook = block
    other = PresentationService(env.store, source_factory=env.factory, client=env.client, clock=lambda:env.now, waiter=lambda _:False)
    with ThreadPoolExecutor(max_workers=2) as pool:
        first = pool.submit(env.service.tick)
        assert entered.wait(5)
        assert other.tick()['busy'] is True
        release.set()
        assert first.result()['scanned'] is True
    assert env.calls == 1


@pytest.mark.parametrize('revoke', ['grant','inactive','password','unlink','expired'])
def test_revocation_or_binding_expiry_cancels_subscription_and_pending(env, revoke):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    env.failures['1'] = MaxAPIError('MAX_ATTACHMENT_NOT_READY','safe',retryable=True)
    advance(env)
    with sqlite3.connect(env.store.accounts_path) as db:
        if revoke == 'grant': db.execute("UPDATE users SET modules='[]' WHERE id=1")
        if revoke == 'inactive': db.execute('UPDATE users SET is_active=0 WHERE id=1')
        if revoke == 'password': db.execute("UPDATE users SET password_hash='new-hash' WHERE id=1")
    if revoke == 'unlink': env.store.unlink(1)
    if revoke == 'expired': env.now = NOW + 2592001
    env.failures.clear()
    advance(env, 300)
    assert env.sent == []
    assert states(env) == [('1',1,'cancelled')]
    with env.store._connection() as db:
        assert db.execute('SELECT enabled FROM max_presentations_subscriptions WHERE user_id=1').fetchone()[0] == 0


def test_relink_same_identity_requires_new_opt_in(env):
    subscribe(env)
    env.service.tick()
    with env.store._connection() as db:
        db.execute('UPDATE max_bot_bindings SET created_at=created_at+1 WHERE user_id=1')
    add(env, 1)
    advance(env)
    assert env.sent == []
    assert env.service.status(1)['enabled'] is False
    subscribe(env)
    add(env, 2)
    advance(env)
    assert [s[1] for s in env.sent] == ['2']


def test_revocation_after_upload_blocks_send(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    def revoke():
        with sqlite3.connect(env.store.accounts_path) as db:
            db.execute("UPDATE users SET modules='[]' WHERE id=1")
    env.after_upload = revoke
    advance(env)
    assert env.uploaded == ['1.pptx']
    assert env.sent == []
    assert states(env) == [('1',1,'cancelled')]


def test_unsubscribe_after_upload_blocks_send_and_reenable_has_no_backlog(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    env.after_upload = lambda:subscribe(env, enabled=False)
    advance(env)
    assert env.sent == []
    env.after_upload = None
    subscribe(env)
    advance(env)
    assert env.sent == []
    add(env, 2)
    advance(env)
    assert [s[1] for s in env.sent] == ['2']


def test_source_failure_does_not_establish_baseline_or_lose_new_files(env):
    subscribe(env)
    add(env, 1)
    env.source_error = RuntimeError('private credential MUST NOT APPEAR')
    result = env.service.tick()
    assert result['error_code'] == 'PRESENTATION_SOURCE_FAILED'
    assert env.service.status(1)['baseline_ready'] is False
    assert 'credential' not in repr(env.service.status(1))
    env.source_error = None
    advance(env, 300)
    assert env.sent == []
    add(env, 2)
    env.source_error = RuntimeError('secret')
    advance(env)
    env.source_error = None
    advance(env, 300)
    assert [s[1] for s in env.sent] == ['2']


def test_cannot_subscribe_other_account_or_without_right(env):
    with pytest.raises(StoreError):
        env.service.set_subscription({'id':2}, '101', '201', True)
    with sqlite3.connect(env.store.accounts_path) as db:
        db.execute("UPDATE users SET modules='[\"zips\"]' WHERE id=1")
        db.execute('INSERT INTO account_managers VALUES(1)')
    with pytest.raises(StoreError) as error:
        subscribe(env)
    assert error.value.code == 'MODULE_FORBIDDEN'


def test_status_only_instance_has_no_network_side_effects(env):
    subscribe(env)
    service = PresentationService(env.store, clock=lambda:env.now, waiter=lambda _:False)
    assert service.status({'id':1})['enabled'] is True
    assert env.calls == 0 and env.sent == []


def test_batch_is_bounded_and_remaining_deliveries_continue(env):
    subscribe(env)
    env.service.tick()
    for i in range(13): add(env, i)
    assert advance(env)['sent'] == 10
    assert env.service.status(1)['pending'] == 3
    assert env.service.tick()['sent'] == 3
    assert len(env.sent) == 13


def test_invalid_listing_is_atomic_not_partial_baseline(env):
    subscribe(env)
    add(env, 1); add(env, 1)
    assert env.service.tick()['error_code'] == 'PRESENTATION_SOURCE_FAILED'
    assert env.service.status(1)['baseline_ready'] is False
    with env.store._connection() as db:
        assert db.execute('SELECT count(*) FROM max_presentations_files').fetchone()[0] == 0


def test_delivery_uses_actual_transport_contract_with_private_routing(env):
    import json
    import httpx
    from services.max_bot.client import MaxClient
    requests = []
    def handle(request):
        requests.append(request)
        if request.url.path == '/uploads':
            assert request.headers['Authorization'] == 'synthetic-token'
            return httpx.Response(200, json={'url':'https://fu.oneme.ru/api/upload.do?token=signed-upload'})
        if request.url.host == 'fu.oneme.ru':
            assert 'Authorization' not in request.headers
            assert b'synthetic-file' in request.read()
            return httpx.Response(200, json={'token':'synthetic-file-token'})
        assert request.url.path == '/messages'
        assert request.url.params['user_id'] == '101'
        assert json.loads(request.content)['attachments'] == [{'type':'file','payload':{'token':'synthetic-file-token'}}]
        return httpx.Response(200, json={'message':{'body':{'mid':'synthetic-mid'}}})
    client = MaxClient('synthetic-token', transport=httpx.MockTransport(handle), sleeper=lambda _:None)
    env.service.client = client
    try:
        subscribe(env)
        env.service.tick()
        add(env, 1)
        assert advance(env)['sent'] == 1
        assert len(requests) == 3
    finally:
        client.close()


def test_oversize_source_file_is_failed_without_automatic_retry(env):
    from services.max_bot.presentation_source import SourceError
    @contextmanager
    def factory():
        class Source:
            def list_attachments(self): return list(env.items)
            def download(self, item, directory):
                raise SourceError('SOURCE_TOO_LARGE', 'safe')
        yield Source()
    env.service.source_factory = factory
    subscribe(env)
    env.service.tick()
    add(env, 1)
    advance(env)
    assert states(env) == [('1',1,'failed')]
    assert env.service.status(1)['failed'] == 1
    assert env.sent == []


def test_stale_lease_recovery_works_after_process_loss(env):
    subscribe(env)
    with env.store._connection() as db:
        db.execute("INSERT INTO max_presentations_lease VALUES(1,'other-process',?)", (env.now+120,))
    assert env.service.tick()['busy'] is True
    assert advance(env, 121)['scanned'] is True


def test_shutdown_during_upload_never_sends_and_cleans_file(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    env.after_upload = env.service.close
    assert advance(env)['sent'] == 0
    assert env.uploaded == ['1.pptx'] and env.sent == []
    assert states(env) == [('1',1,'retry')]
    assert all(not path.exists() for path in env.downloads)


def test_attachment_processing_reuses_token_and_upload(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    calls, waits = [], []
    original = env.client.send_file
    def send(max_id, token, caption):
        calls.append(token)
        if len(calls) == 1:
            raise MaxAPIError('MAX_ATTACHMENT_NOT_READY','safe',retryable=True)
        return original(max_id, token, caption)
    env.client.send_file = send
    env.service._wait = lambda delay:waits.append(delay) or False
    assert advance(env)['sent'] == 1
    assert env.uploaded == ['1.pptx'] and calls == ['1','1'] and waits == [2]


def test_processing_wait_rechecks_revocation_and_shutdown(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    env.failures['1'] = MaxAPIError('MAX_ATTACHMENT_NOT_READY','safe',retryable=True)
    attempts = []
    original = env.client.send_file
    def send(*args):
        attempts.append(1)
        return original(*args)
    env.client.send_file = send
    env.service._wait = lambda _:env.store.unlink(1) or False
    advance(env)
    assert attempts == [1] and env.sent == []
    assert states(env) == [('1',1,'cancelled')]


@pytest.mark.parametrize('uncertain', [True, False])
def test_explicit_retry_is_personal_idempotent_and_never_replays_sent(env, uncertain):
    subscribe(env); subscribe(env, 2)
    env.service.tick()
    add(env, 1)
    env.failures['1'] = MaxAPIError('MAX_RESPONSE_INVALID','safe',uncertain=uncertain)
    advance(env)
    one, two = env.service.status(1), env.service.status(2)
    assert len(one['issues']) == len(two['issues']) == 1
    own, foreign = one['issues'][0]['id'], two['issues'][0]['id']
    assert own != foreign and one['issues'][0]['retryable'] is True
    with pytest.raises(StoreError):
        env.service.retry_delivery({'id':1}, '101', '201', foreign)
    with pytest.raises(StoreError):
        env.service.retry_delivery({'id':1}, '102', '202', own)
    result = env.service.retry_delivery({'id':1}, '101', '201', own)
    assert result['issues'] == [] and result['pending'] == 1
    assert env.service.retry_delivery({'id':1}, '101', '201', own)['pending'] == 1
    env.failures.clear()
    assert env.service.tick()['sent'] == 1
    with pytest.raises(StoreError) as error:
        env.service.retry_delivery({'id':1}, '101', '201', own)
    assert error.value.code == 'DELIVERY_NOT_RETRYABLE'
    assert len(env.sent) == 1
    assert env.service.status(2)['issues'][0]['id'] == foreign


def test_retry_denied_after_revoke_or_subscription_generation_change(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    env.failures['1'] = MaxAPIError('MAX_READ_TIMEOUT','safe',uncertain=True)
    advance(env)
    issue = env.service.status(1)['issues'][0]['id']
    subscribe(env, enabled=False)
    subscribe(env)
    assert env.service.status(1)['issues'][0]['retryable'] is False
    with pytest.raises(StoreError):
        env.service.retry_delivery(1, '101', '201', issue)
    with sqlite3.connect(env.store.accounts_path) as db:
        db.execute("UPDATE users SET modules='[]' WHERE id=1")
    with pytest.raises(StoreError):
        env.service.retry_delivery(1, '101', '201', issue)


def test_issues_limit_is_five_and_safe_text_is_not_interpreted(env):
    subscribe(env)
    env.service.tick()
    for i in range(7):
        add(env, i)
        env.items[-1].name = '<img src=x onerror=alert(1)>-'+str(i)
        env.failures[str(i)] = MaxAPIError('MAX_RESPONSE_INVALID','safe',uncertain=True)
    advance(env)
    issues = env.service.status(1)['issues']
    assert len(issues) == 5 and len({i['id'] for i in issues}) == 5
    assert all(i['name'].startswith('<img') and set(i) == {'id','name','code','state','retryable'} for i in issues)


def test_old_worker_cannot_overwrite_new_owner_delivery_after_losing_lease(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    def takeover():
        with env.store._connection() as db:
            db.execute("UPDATE max_presentations_lease SET owner='new-owner',expires_at=?", (env.now+120,))
            db.execute("UPDATE max_presentations_deliveries SET state='sending',last_error=NULL")
    env.after_upload = takeover
    advance(env)
    assert env.sent == []
    assert states(env) == [('1',1,'sending')]


def test_old_successful_sender_cannot_clobber_new_owner_after_lease_loss(env):
    subscribe(env)
    env.service.tick()
    add(env, 1)
    def send(*_):
        with env.store._connection() as db:
            db.execute("UPDATE max_presentations_lease SET owner='new-owner',expires_at=?", (env.now+120,))
            db.execute("UPDATE max_presentations_deliveries SET state='uncertain',last_error='DELIVERY_INTERRUPTED'")
        return True
    env.client.send_file = send
    assert advance(env)['sent'] == 0
    assert states(env) == [('1',1,'uncertain')]


def test_old_scanner_cannot_replace_new_owners_state(env):
    def takeover():
        with env.store._connection() as db:
            db.execute("UPDATE max_presentations_lease SET owner='new-owner',expires_at=?", (env.now+120,))
        raise RuntimeError('old request failed')
    env.list_hook = takeover
    env.service.tick()
    with env.store._connection() as db:
        state = db.execute('SELECT * FROM max_presentations_state').fetchone()
    assert state['last_error'] is None and state['failures'] == 0


def test_bad_source_login_pauses_all_pending_delivery_until_success(env):
    from services.max_bot.presentation_source import SourceError
    subscribe(env)
    env.service.tick()
    for i in range(3):
        add(env, i)
        env.failures[str(i)] = MaxAPIError('MAX_ATTACHMENT_NOT_READY','safe',retryable=True)
    advance(env)
    downloads = len(env.downloads)
    env.source_error = SourceError('SOURCE_AUTH_FAILED','safe')
    advance(env)
    assert len(env.downloads) == downloads
    calls = env.calls
    for _ in range(5): advance(env, 1)
    assert env.calls == calls and len(env.downloads) == downloads
    env.source_error = None
    env.failures.clear()
    assert advance(env, 300)['sent'] == 3


def test_download_auth_failure_stops_batch_after_single_login(env):
    from services.max_bot.presentation_source import SourceError
    subscribe(env)
    env.service.tick()
    for i in range(3): add(env, i)
    attempts = []
    @contextmanager
    def factory():
        class Source:
            def list_attachments(self): return list(env.items)
            def download(self, *_):
                attempts.append(1)
                raise SourceError('SOURCE_AUTH_FAILED','safe')
        yield Source()
    env.service.source_factory = factory
    assert advance(env)['error_code'] == 'SOURCE_AUTH_FAILED'
    assert attempts == [1]
    for _ in range(5): advance(env, 1)
    assert attempts == [1]
    assert env.service.status(1)['pending'] == 3
