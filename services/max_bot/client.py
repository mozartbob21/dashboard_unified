"""Narrow MAX Bot API transport; no webhook mutation or live call on import.

API documentation rechecked 22.10.2026. Long Polling is development-only.
All callers must already have authorized the user and the outgoing data.
"""
from __future__ import annotations

import json
import socket
import ssl
import threading
import time
from pathlib import Path

import httpx


API_ORIGIN = "https://platform-api2.max.ru"
UPDATE_TYPES = ("message_created", "message_callback", "bot_started", "bot_stopped")
MAX_RESPONSE_BYTES = 2 * 1024 * 1024
MAX_BODY_BYTES = 32 * 1024
POLL_TIMEOUT = 30
SEND_INTERVAL = 0.55


class MaxAPIError(Exception):
    """Public diagnostics without request/response bodies, URLs or secrets."""

    def __init__(self, code, message, *, retryable=False, uncertain=False):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = retryable
        self.uncertain = uncertain


def classify_network_error(exc, *, write=False):
    """Inspect nested transport errors; return only fixed, non-sensitive text."""
    pending, seen, causes = [exc], set(), []
    while pending and len(seen) < 32:
        error = pending.pop()
        if not isinstance(error, BaseException) or id(error) in seen:
            continue
        seen.add(id(error))
        causes.append(error)
        pending.extend((error.__cause__, error.__context__))
        pending.extend(arg for arg in error.args if isinstance(arg, BaseException))
    # Some httpcore versions retain OpenSSL's error code only in a wrapper's
    # message. Inspect that signature, but never return/log the raw exception.
    certificate_failure = any(isinstance(error, ssl.SSLCertVerificationError)
        or isinstance(error, (ssl.SSLError, httpx.ConnectError))
        and "CERTIFICATE_VERIFY_FAILED" in str(error) for error in causes)
    if certificate_failure:
        return MaxAPIError("MAX_TLS_CERTIFICATE_FAILED",
            "Не удалось проверить TLS-сертификат MAX. Укажите согласованный PEM-файл "
            "доверенных сертификатов в MAX_ZIP_BOT_CA_FILE. Проверка TLS остаётся включённой; "
            "повторные подключения остановлены до исправления настройки.", uncertain=write)
    if any(isinstance(error, socket.gaierror) for error in causes):
        code, message = "MAX_DNS_FAILED", "Не удалось определить адрес сервера MAX. Проверьте DNS и подключение к сети."
    elif any(isinstance(error, httpx.ConnectTimeout) for error in causes):
        code, message = "MAX_CONNECT_TIMEOUT", "Истекло время установления соединения с MAX. Проверьте доступ к API из этой сети."
    elif any(isinstance(error, httpx.ReadTimeout) for error in causes):
        code, message = "MAX_READ_TIMEOUT", "MAX не ответил за отведённое время. Соединение будет проверено повторно."
    elif any(isinstance(error, ssl.SSLError) for error in causes):
        code, message = "MAX_TLS_FAILED", "Не удалось согласовать защищённое соединение с MAX. Проверьте настройки TLS и сети."
    else:
        code, message = "MAX_CONNECTION_FAILED", "Не удалось связаться с MAX. Проверьте доступность сети и API."
    return MaxAPIError(code, message, retryable=not write, uncertain=write)


def _invalid(message="Некорректные параметры операции MAX."):
    return MaxAPIError("INVALID_REQUEST", message)


def _json_object(pairs):
    obj = {}
    for key, value in pairs:
        if key in obj:
            raise ValueError("duplicate key")
        obj[key] = value
    return obj


def _nonfinite(_value):
    raise ValueError("nonfinite")


def _body(value):
    if not isinstance(value, dict) or set(value) - {"text", "attachments", "notify"}:
        raise _invalid("Некорректное сообщение MAX.")
    text = value.get("text")
    if not isinstance(text, str) or not text.strip() or len(text) > 4000:
        raise _invalid("Сообщение должно содержать от 1 до 4000 символов.")
    if "notify" in value and type(value["notify"]) is not bool:
        raise _invalid()
    attachments = value.get("attachments", [])
    if not isinstance(attachments, list) or len(attachments) > 1:
        raise _invalid()
    for attachment in attachments:
        if not isinstance(attachment, dict) or set(attachment) != {"type", "payload"} or attachment["type"] != "inline_keyboard":
            raise _invalid()
        payload = attachment["payload"]
        if not isinstance(payload, dict) or set(payload) != {"buttons"}:
            raise _invalid()
        rows = payload["buttons"]
        if not isinstance(rows, list) or len(rows) > 20:
            raise _invalid()
        for row in rows:
            if not isinstance(row, list) or not 1 <= len(row) <= 8:
                raise _invalid()
            for button in row:
                if not isinstance(button, dict) or button.get("type") != "callback" or set(button) - {"type", "text", "payload", "intent"}:
                    raise _invalid()
                if not isinstance(button.get("text"), str) or not 1 <= len(button["text"]) <= 128:
                    raise _invalid()
                if not isinstance(button.get("payload"), str) or not 1 <= len(button["payload"].encode("utf-8")) <= 256:
                    raise _invalid()
                if button.get("intent") not in (None, "positive", "negative", "default"):
                    raise _invalid()
    try:
        encoded = json.dumps(value, ensure_ascii=False, allow_nan=False).encode("utf-8")
    except (TypeError, ValueError, UnicodeError, OverflowError):
        raise _invalid() from None
    if len(encoded) > MAX_BODY_BYTES:
        raise _invalid("Сообщение слишком большое для отправки.")
    return json.loads(encoded)


class MaxClient:
    def __init__(self, bot_token, *, ca_bundle=None, transport=None, clock=time.monotonic, sleeper=time.sleep):
        if not isinstance(bot_token, str) or not bot_token.strip() or len(bot_token) > 4096 or any(ord(c) < 33 or ord(c) > 126 for c in bot_token):
            raise MaxAPIError("CONFIG_INVALID", "Токен MAX не настроен или имеет некорректный формат.")
        verify = True
        if ca_bundle is not None:
            # Use only an explicitly supplied operator-managed trust bundle.
            # Never download/install a CA and never disable verification.
            try:
                path = Path(ca_bundle)
                verify = ssl.create_default_context(cafile=str(path))
            except (OSError, ValueError, TypeError, ssl.SSLError):
                raise MaxAPIError("TLS_CONFIG_INVALID", "Не удалось загрузить согласованный файл доверенных сертификатов.") from None
        self._http = httpx.Client(
            base_url=API_ORIGIN,
            headers={"Authorization": bot_token, "Accept": "application/json", "Accept-Encoding": "identity"},
            verify=verify, trust_env=False, follow_redirects=False,
            timeout=httpx.Timeout(connect=10.0, read=35.0, write=10.0, pool=10.0),
            transport=transport,
        )
        self._clock, self._sleep = clock, sleeper
        self._send_lock = threading.Lock()
        self._last_send = None

    def close(self):
        self._http.close()

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        self.close()

    def _request(self, method, path, *, params=None, body=None):
        # No arbitrary URLs, proxies, paths, redirect targets or user-supplied
        # header forwarding are available through this adapter.
        if (method, path) not in {("GET", "/me"), ("GET", "/subscriptions"), ("GET", "/updates"),
                                  ("POST", "/messages"), ("POST", "/answers")}:
            raise _invalid()
        write = method == "POST"
        try:
            with self._http.stream(method, path, params=params, json=body) as response:
                if response.is_redirect:
                    raise MaxAPIError("REDIRECT_DENIED", "MAX вернул перенаправление. Администратору нужно проверить подключение.", uncertain=write)
                if response.status_code in (401, 403):
                    raise MaxAPIError("MAX_AUTH_FAILED", "MAX отклонил доступ. Администратору нужно проверить токен бота.")
                if response.status_code == 429:
                    raise MaxAPIError("MAX_RATE_LIMIT", "MAX ограничил частоту запросов. Повторите позже.", retryable=not write)
                if response.status_code >= 500:
                    raise MaxAPIError("MAX_UNAVAILABLE", "MAX временно недоступен. Повторите позже.", retryable=not write, uncertain=write)
                if not 200 <= response.status_code < 300:
                    raise MaxAPIError("MAX_REQUEST_REJECTED", "MAX не принял запрос. Повторите действие позже.")
                declared = response.headers.get("content-length")
                if declared is not None and (not declared.isdecimal() or int(declared) > MAX_RESPONSE_BYTES):
                    raise ValueError("response size")
                if response.headers.get("content-encoding", "identity").lower() != "identity":
                    raise ValueError("compressed response")
                content = bytearray()
                for chunk in response.iter_bytes(chunk_size=65536):
                    content.extend(chunk)
                    if len(content) > MAX_RESPONSE_BYTES:
                        raise ValueError("response size")
                result = json.loads(content.decode("utf-8"), object_pairs_hook=_json_object, parse_constant=_nonfinite)
                if not isinstance(result, dict):
                    raise ValueError("response object")
                return result
        except MaxAPIError:
            raise
        except (httpx.HTTPError, ssl.SSLError) as error:
            raise classify_network_error(error, write=write) from None
        except (ValueError, UnicodeError, RecursionError, OverflowError):
            raise MaxAPIError("MAX_RESPONSE_INVALID", "Не удалось прочитать ответ MAX. Повторите действие позже.",
                              retryable=not write, uncertain=write) from None

    def get_me(self):
        result = self._request("GET", "/me")
        if result.get("is_bot") is not True or type(result.get("user_id")) is not int or not 0 < result["user_id"] < 2**63:
            raise MaxAPIError("MAX_BOT_INVALID", "Токен не подтвердил существующего бота MAX.")
        return {"user_id": result["user_id"], "is_bot": True}

    def check_no_webhook(self):
        result = self._request("GET", "/subscriptions")
        subscriptions = result.get("subscriptions")
        if not isinstance(subscriptions, list):
            raise MaxAPIError("MAX_RESPONSE_INVALID", "Не удалось проверить подписки MAX.", retryable=True)
        if subscriptions:
            raise MaxAPIError("WEBHOOK_ACTIVE", "У бота уже настроен Webhook. Тестовый Long Polling остановлен; подписки не изменены.")

    def get_updates(self, marker=None):
        if marker is not None and (type(marker) is not int or not 0 <= marker < 2**63):
            raise _invalid("Некорректная позиция чтения событий MAX.")
        params = {"timeout": POLL_TIMEOUT, "limit": 50, "types": ",".join(UPDATE_TYPES)}
        if marker is not None:
            params["marker"] = marker
        result = self._request("GET", "/updates", params=params)
        updates, next_marker = result.get("updates"), result.get("marker")
        if (not isinstance(updates, list) or len(updates) > 50 or any(not isinstance(update, dict) for update in updates)
                or next_marker is not None and (type(next_marker) is not int or not 0 <= next_marker < 2**63)):
            raise MaxAPIError("MAX_RESPONSE_INVALID", "Не удалось прочитать события MAX.", retryable=True)
        if marker is not None and next_marker is not None and next_marker < marker:
            raise MaxAPIError("MAX_MARKER_REGRESSED", "MAX вернул устаревшую позицию событий. Чтение остановлено.")
        if updates and next_marker is None:
            raise MaxAPIError("MAX_RESPONSE_INVALID", "MAX не подтвердил позицию прочитанных событий.", retryable=True)
        return {"updates": updates, "marker": next_marker}

    def _throttle(self):
        now = self._clock()
        if self._last_send is not None:
            delay = SEND_INTERVAL - (now - self._last_send)
            if delay > 0:
                self._sleep(delay)
        self._last_send = self._clock()

    def send(self, user_id, body):
        if type(user_id) is not int or not 0 < user_id < 2**63:
            raise _invalid("Некорректный получатель сообщения MAX.")
        body = _body(body)
        with self._send_lock:
            self._throttle()
            result = self._request("POST", "/messages", params={"user_id": user_id, "disable_link_preview": "true"}, body=body)
        if not isinstance(result.get("message"), dict):
            raise MaxAPIError("MAX_RESPONSE_INVALID", "MAX не подтвердил отправку сообщения. Повторите запрос в боте.", uncertain=True)
        return True

    def answer(self, callback_id, body):
        if not isinstance(callback_id, str) or not 1 <= len(callback_id) <= 512 or any(ord(c) < 33 for c in callback_id):
            raise _invalid("Некорректная кнопка MAX.")
        body = _body(body)
        with self._send_lock:
            self._throttle()
            result = self._request("POST", "/answers", params={"callback_id": callback_id, "disable_link_preview": "true"}, body={"message": body})
        if result.get("success") is not True:
            raise MaxAPIError("MAX_ANSWER_REJECTED", "MAX не подтвердил обновление сообщения. Повторите запрос в боте.", uncertain=True)
        return True
