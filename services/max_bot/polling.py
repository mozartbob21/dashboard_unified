"""Explicit development-only MAX polling. Importing this module never connects.

Each update is durably claimed before business side effects. Claimed updates are
not automatically repeated after a crash or an uncertain send: the user can send
a fresh request. This prefers avoiding duplicate linking/messages over guessing
whether MAX accepted an interrupted request.
"""
from __future__ import annotations

import argparse
import logging
import os
import signal
import sys
import threading
import time
import uuid
from pathlib import Path

from .client import MaxAPIError, MaxClient, UPDATE_TYPES
from .store import StoreError


log = logging.getLogger("max_bot.polling")
POLL_INTERVAL = 0.5
LEASE_TTL = 120


class PollingError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = False


class DevPoller:
    def __init__(self, store, client, handler, *, owner=None, clock=time.monotonic, sleeper=time.sleep, presentations=None):
        self.store, self.client, self.handler = store, client, handler
        self.owner = owner or uuid.uuid4().hex
        self._clock, self._sleep = clock, sleeper
        self._started = False
        self._last_poll = None
        self.presentations = presentations

    def _lease(self):
        if not self.store.lease(self.owner, ttl=LEASE_TTL):
            raise PollingError("LEASE_LOST", "Другой процесс уже получает события этого бота. Тестовый запуск остановлен.")

    def _audit(self, event, outcome):
        try:
            if self.store.audit(event, outcome) is False:
                log.warning("MAX technical audit unavailable")
        except Exception:
            log.warning("MAX technical audit unavailable")

    def start(self):
        if self._started:
            return
        self._lease()
        try:
            self.client.check_no_webhook()
            me = self.client.get_me()
            # A different bot cannot reuse this bot's users, cursor or receipts.
            self.store.ensure_bot_id(me["user_id"])
            self._lease()
            self._started = True
            if self.presentations is not None:
                self.presentations.start()
        except BaseException:
            self.store.release_lease(self.owner)
            raise

    def close(self):
        try:
            if self.presentations is not None:
                self.presentations.close()
            self.store.release_lease(self.owner)
        finally:
            self._started = False

    def poll_once(self):
        self.start()
        self._lease()
        # A webhook might have been enabled after this worker started. Never
        # disable it or continue competing with it; only GET is used here.
        self.client.check_no_webhook()
        if self._last_poll is not None:
            delay = POLL_INTERVAL - (self._clock() - self._last_poll)
            if delay > 0:
                self._sleep(delay)
        self._lease()
        marker = self.store.get_marker()
        self._last_poll = self._clock()
        batch = self.client.get_updates(marker)
        self._lease()
        result = {"received": len(batch["updates"]), "sent": 0, "ignored": 0, "failed": 0, "uncertain": 0}
        for update in batch["updates"]:
            self._lease()
            if update.get("update_type") not in UPDATE_TYPES:
                result["ignored"] += 1
                continue
            try:
                key = self.handler.event_key(update)
            except Exception:
                # Malformed envelopes do not expose content in logs or stop
                # processing the next valid update in the same bounded batch.
                self._audit("poll", "invalid")
                result["ignored"] += 1
                continue
            if not key or not self.store.claim_event(key):
                result["ignored"] += 1
                continue
            try:
                reply = self.handler.handle(update)
            except Exception:
                self.store.finish_event(key, "failed")
                self._audit("poll", "error")
                log.warning("MAX handler failed; update will not be repeated automatically")
                result["failed"] += 1
                continue
            if reply is None:
                self.store.finish_event(key, "ignored")
                result["ignored"] += 1
                continue
            self._lease()
            try:
                if not isinstance(reply, dict) or not isinstance(reply.get("body"), dict):
                    raise MaxAPIError("INVALID_REPLY", "Не удалось подготовить ответ бота.")
                callback_id = reply.get("callback_id")
                if callback_id:
                    self.client.answer(callback_id, reply["body"])
                else:
                    self.client.send(reply.get("user_id"), reply["body"])
            except MaxAPIError as error:
                status = "outbound_uncertain" if error.uncertain else "failed"
                self.store.finish_event(key, status)
                result["uncertain" if error.uncertain else "failed"] += 1
                self._audit("send", "uncertain" if error.uncertain else "error")
                log.warning("MAX send failed code=%s", error.code)
                if error.code == "MAX_AUTH_FAILED":
                    raise
            except Exception:
                # The write may have happened. Do not retry or expose exception
                # strings that could include phones, queries or a bot token.
                self.store.finish_event(key, "outbound_uncertain")
                result["uncertain"] += 1
                self._audit("send", "uncertain")
                log.warning("MAX send outcome unknown")
            else:
                self.store.finish_event(key, "sent")
                result["sent"] += 1
                self._audit("send", "ok")
        # Advance only after every claimed result has been handled. A crash
        # earlier re-fetches the batch but durable claims prevent side effects.
        next_marker = batch.get("marker")
        if next_marker is not None:
            self._lease()
            self.store.set_marker(next_marker, owner=self.owner)
        self._audit("poll", "ok")
        return result

    def run(self, *, stop=None, max_iterations=None):
        stop = stop or threading.Event()
        failures = 0
        iterations = 0
        try:
            while not stop.is_set() and (max_iterations is None or iterations < max_iterations):
                iterations += 1
                try:
                    self.poll_once()
                except MaxAPIError as error:
                    if not error.retryable:
                        raise
                    self._audit("poll", "unavailable")
                    delay = min(30, 2 ** min(failures + 1, 5))
                    failures += 1
                    log.warning("MAX polling interrupted code=%s; retry in %s seconds. %s", error.code, delay, error.message)
                    self._sleep(delay)
                except StoreError as error:
                    if error.code != "SYSTEM_UNAVAILABLE":
                        raise
                    failures += 1
                    log.warning("MAX local storage unavailable; polling stopped")
                    raise
                else:
                    failures = 0
        finally:
            self.close()
        return iterations


class _SafeParser(argparse.ArgumentParser):
    def error(self, _message):
        # Do not echo unknown arguments that may accidentally contain secrets.
        self.print_usage(sys.stderr)
        self.exit(2, "Ошибка параметров запуска. Используйте --help.\n")


def _arguments(argv):
    parser = _SafeParser(description="Локальное тестирование MAX-бота. Production Long Polling не поддерживается.")
    actions = parser.add_mutually_exclusive_group(required=True)
    actions.add_argument("--run-development-only", action="store_true", help="Явно разрешить тестовые обращения к MAX")
    actions.add_argument("--check-config", action="store_true", help="Проверить настройки без запросов к MAX и без создания БД")
    parser.add_argument("--env-file", type=Path, help="Файл настроек оператора; значения окружения имеют приоритет")
    parser.add_argument("--max-polls", type=int, help="Остановиться после N тестовых циклов")
    args = parser.parse_args(argv)
    if args.max_polls is not None and (not args.run_development_only or not 1 <= args.max_polls <= 10000):
        parser.error("invalid poll limit")
    return args


def main(argv=None):
    args = _arguments(argv)
    try:
        environment = dict(os.environ)
        if args.env_file is not None:
            from dotenv import dotenv_values
            if not args.env_file.is_file() or args.env_file.stat().st_size > 64 * 1024:
                raise ValueError("invalid env file")
            values = dotenv_values(args.env_file, interpolate=False)
            environment = {**{key: value for key, value in values.items() if value is not None}, **environment}
        from .config import Settings
        settings = Settings.from_env(environment)
        if args.check_config:
            if settings.enabled:
                settings.validate_polling()
            mode_label = "локальное тестирование" if settings.mode == "development-polling" else "выключен"
            print(f"Настройки проверены. Режим: {mode_label}. "
                  + ("Бот включён." if settings.enabled else "Бот выключен.")
                  + (" Токен настроен." if settings.bot_token else " Токен не настроен."))
            return 0
        if not settings.enabled or settings.mode != "development-polling" or not settings.bot_token:
            print("Тестовый бот выключен или не настроен. Проверьте MAX_ZIP_BOT_ENABLED, MAX_ZIP_BOT_MODE и MAX_ZIP_BOT_TOKEN.", file=sys.stderr)
            return 2
        from .handler import BotHandler
        from .stock import StockService
        from .store import Store
        store = Store(settings)
        stocks = StockService(settings.published_path, settings.contacts_path, stale_days=30)
        handler = BotHandler(store, stocks)
    except Exception:
        print("Не удалось проверить конфигурацию тестового MAX-бота. Проверьте файл настроек и доступ к отдельной служебной БД.", file=sys.stderr)
        return 2

    stop = threading.Event()
    previous = {}
    for sig in (signal.SIGINT, signal.SIGTERM):
        previous[sig] = signal.signal(sig, lambda *_args: stop.set())
    try:
        with MaxClient(settings.bot_token, ca_bundle=settings.ca_file) as client:
            from .presentations import PresentationService
            from .presentation_config import source_factory
            presentations = PresentationService(store, source_factory=source_factory, client=client)
            handler.presentations = presentations
            DevPoller(store, client, handler, presentations=presentations).run(stop=stop, max_iterations=args.max_polls)
        return 0
    except KeyboardInterrupt:
        return 0
    except (MaxAPIError, PollingError, StoreError) as error:
        print(f"{error.code}: {error.message}", file=sys.stderr)
        return 1
    except Exception:
        print("Тестовый MAX-бот остановлен из-за ошибки. Проверьте состояние сервиса; секреты в журнал не выводятся.", file=sys.stderr)
        return 1
    finally:
        for sig, prior in previous.items():
            signal.signal(sig, prior)


if __name__ == "__main__":
    raise SystemExit(main())
