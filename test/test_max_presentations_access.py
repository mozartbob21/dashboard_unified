"""Separate grants and opt-in UX with synthetic accounts, never office data."""
import json
import sqlite3

from test.test_max_bot_link_api import api, post, pending
from test.test_max_bot_handler import handler, event, callback, button, link
from services.max_bot.presentations import PresentationService


def grant(accounts, modules, user=1):
    with sqlite3.connect(accounts) as conn:
        conn.execute('UPDATE users SET modules=? WHERE id=?', (json.dumps(modules), user))
        conn.execute('UPDATE account_control SET manager_user_id=NULL')
        conn.execute('DELETE FROM account_managers')


def test_nonmanager_can_link_presentations_without_stock_access(api):
    grant(api.config.accounts_path, ['presentations'])
    api.actor['user']['modules'] = ['presentations']
    response = api.client.get('/max-bot')
    assert response.status_code == 200
    assert '/presentations' in response.text
    assert 'href="/zips"' not in response.text
    assert 'synthetic-token' not in response.text
    request_id = pending(api)
    assert post(api, 'confirm', {'request_id': request_id}).status_code == 200
    assert api.store().authorize(123456789, 987654321)['modules'] == ['presentations']
    # A stale session cannot preserve a revoked permission.
    grant(api.config.accounts_path, [])
    assert api.client.get('/max-bot').status_code == 403
    assert post(api, 'link').status_code == 403


def test_ordinary_zips_account_still_cannot_link(api):
    grant(api.config.accounts_path, ['zips'])
    assert post(api, 'link').status_code == 403


def test_presentations_subscription_is_explicit_and_survives_restart(handler):
    grant(handler.store.accounts_path, ['presentations'])
    handler.presentations = PresentationService(handler.store)
    link(handler)
    account = handler.store.authorize(123456789, 321)
    assert not handler.presentations.status(account)['subscribed']
    menu = handler.handle(event('/menu'))
    page = handler.handle(callback(button(menu)))
    assert 'Подписка выключена' in page['body']['text']
    subscribed = handler.handle(callback(button(page)))
    assert 'Подписка включена' in subscribed['body']['text']
    handler.presentations = PresentationService(handler.store)
    assert handler.presentations.status(account)['subscribed']
    stopped = handler.handle(callback(button(subscribed)))
    assert 'Подписка выключена' in stopped['body']['text']


def test_presentations_never_unlocks_stock_or_contacts(handler):
    link(handler)
    contact_token = button(handler.handle(event()))
    # Even a combined grant never bypasses the existing manager requirement for Zips.
    grant(handler.store.accounts_path, ['presentations', 'zips'])
    for response in (handler.handle(event('Муфта D100')), handler.handle(callback(contact_token))):
        assert 'управляющим' in response['body']['text']
        assert 'Власиха' not in str(response)
        assert '+7' not in str(response)


def test_old_subscription_button_rechecks_revoked_grant(handler):
    grant(handler.store.accounts_path, ['presentations'])
    handler.presentations = PresentationService(handler.store)
    link(handler)
    page = handler.handle(event('/presentations'))
    grant(handler.store.accounts_path, [])
    denied = handler.handle(callback(button(page)))
    assert 'доступ' in denied['body']['text'].lower()
    with handler.store._connection() as conn:
        assert conn.execute('SELECT count(*) FROM max_presentations_subscriptions').fetchone()[0] == 0


def test_uncertain_delivery_requires_explicit_confirmation_before_retry(handler):
    class Subscription:
        calls = 0
        def status(self, account):
            return {'subscribed': True, 'issues': [] if self.calls else [
                {'id': 7, 'name': 'Презентация.pptx', 'state': 'uncertain',
                 'code': 'MAX_READ_TIMEOUT', 'retryable': True}]}
        def retry_delivery(self, account, uid, chat, delivery):
            assert (account['id'], uid, chat, delivery) == (1, 123456789, 321, 7)
            self.calls += 1
    grant(handler.store.accounts_path, ['presentations'])
    link(handler)
    handler.presentations = Subscription()
    page = handler.handle(event('/presentations'))
    confirm = handler.handle(callback(button(page, 2)))
    assert 'может создать дубликат' in confirm['body']['text']
    assert handler.presentations.calls == 0
    handler.handle(callback(button(confirm)))
    assert handler.presentations.calls == 1
    handler.handle(callback(button(confirm)))
    assert handler.presentations.calls == 1
