"""One existing polling process owns startup/shutdown of presentation monitoring."""
import pytest
from test.test_max_bot_transport import store, Endpoint, Handler, max_client
from services.max_bot import polling
from services.max_bot.client import MaxAPIError


class Monitor:
    def __init__(self): self.started = self.stopped = 0
    def start(self): self.started += 1
    def close(self): self.stopped += 1


def test_monitor_starts_once_after_confirming_bot_identity_and_stops(store):
    monitor = Monitor()
    with max_client(Endpoint()) as client:
        worker = polling.DevPoller(store, client, Handler(), presentations=monitor)
        worker.start()
        worker.start()
        assert monitor.started == 1
        worker.close()
        assert monitor.stopped == 1
        assert store.lease('next-owner')


def test_active_webhook_never_starts_presentation_thread(store):
    endpoint, monitor = Endpoint(), Monitor()
    endpoint.webhook = True
    with max_client(endpoint) as client:
        worker = polling.DevPoller(store, client, Handler(), presentations=monitor)
        with pytest.raises(MaxAPIError): worker.start()
        assert monitor.started == 0
