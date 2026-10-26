"""Permanent links keep opt-in delivery across upgrades and long inactivity.

Only synthetic sources/transports and temporary SQLite databases are used.
"""
import pytest

from services.max_bot.client import MaxAPIError
from services.max_bot.presentations import PresentationService
from services.max_bot.store import PERMANENT_BINDING_EXPIRY, Store
from test.test_max_presentations import add, advance, env, states, subscribe


def connect_permanently(env):
    code = env.store.create_link(1, now=env.now)["code"]
    pending = env.store.redeem_link(code, "101", "201", now=env.now)
    env.store.confirm_link(1, pending["request_id"], now=env.now)


def restart(env):
    env.service.close()
    env.store = Store(env.store.settings)
    env.service = PresentationService(env.store, source_factory=env.factory, client=env.client,
        clock=lambda: env.now, waiter=lambda _: False)


def subscription_row(env):
    with env.store._connection() as db:
        return tuple(db.execute("SELECT * FROM max_presentations_subscriptions WHERE user_id=1").fetchone())


def test_connected_user_keeps_subscription_and_baseline_across_years_and_restart(env):
    connect_permanently(env)
    subscribe(env)
    before = subscription_row(env)
    add(env, 1)
    env.service.tick()  # Existing files form the baseline, never sent as new.
    env.now += 5 * 365 * 86400
    restart(env)
    assert env.store.binding_status(1, now=env.now)["active"]
    assert env.service.status(1)["subscribed"]
    assert subscription_row(env) == before
    add(env, 2)
    assert env.service.tick()["sent"] == 1
    assert states(env) == [("2", 1, "sent")]
    assert [sent[1] for sent in env.sent] == ["2"]


def test_upgrade_preserves_existing_subscription_without_new_opt_in(env):
    # Emulate a valid old link and its existing presentation subscription.
    subscribe(env)
    before = subscription_row(env)
    add(env, 1)
    env.service.tick()
    with env.store._connection() as db:
        db.execute("DELETE FROM max_bot_meta WHERE key='migration_20261023_max_persistent_bindings'")
    restart(env)
    with env.store._connection() as db:
        assert db.execute("SELECT expires_at FROM max_bot_bindings WHERE user_id=1").fetchone()[0] == PERMANENT_BINDING_EXPIRY
    assert subscription_row(env) == before
    env.now += 2 * 365 * 86400
    add(env, 2)
    assert env.service.status(1)["subscribed"]
    assert env.service.tick()["sent"] == 1
    assert [sent[1] for sent in env.sent] == ["2"]


@pytest.mark.parametrize("future", [5 * 365 * 86400, PERMANENT_BINDING_EXPIRY])
def test_explicit_delivery_retry_remains_available_without_link_expiration(env, future):
    connect_permanently(env)
    subscribe(env)
    env.service.tick()
    add(env, 1)
    env.failures["1"] = MaxAPIError("MAX_READ_TIMEOUT", "synthetic", uncertain=True)
    advance(env)
    assert states(env) == [("1", 1, "uncertain")]
    env.now += future
    restart(env)
    env.failures.clear()
    # Time passing never automatically retries an ambiguous send.
    assert env.service.tick()["sent"] == 0
    issue = env.service.status(1)["issues"][0]
    assert issue["retryable"]
    env.service.retry_delivery(1, "101", "201", issue["id"])
    assert env.service.tick()["sent"] == 1
    assert states(env) == [("1", 1, "sent")]
