"""Private-dialog commands. Search uses published data; no LLM or business writes."""
from collections import OrderedDict
import hashlib
import re
import secrets
import time

from .store import StoreError
from .stock import StockError

HELP = ("Остатки ЗиП РСО · Нейрона ИИ\n\n"
        "Для подключения на главной странице Нейроны нажмите «Подключить MAX». "
        "Получите код, отправьте /link КОД и подтвердите подключение на странице Нейроны.\n\n"
        "Затем напишите, что нужно найти: название, диаметр или модель. "
        "Несколько позиций разделяйте точкой с запятой. "
        "Выберите муниципалитет в результате, чтобы получить контакт ответственного.\n\n"
        "Подтверждённое наличие: положительный остаток с датой не старше 30 дней. "
        "Старые записи требуют уточнения.\n/help — помощь\n/stop — отключить доступ. "
        "Пароль от Нейроны боту не отправляйте.")


def text(value, limit=180):
    value = re.sub(r"[\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]", " ", str(value or ""))
    return value if len(value) <= limit else value[:limit - 1] + "…"


class BotHandler:
    def __init__(self, store, stock, clock=time.monotonic):
        self.store, self.stock, self.clock = store, stock, clock
        self.actions = OrderedDict()

    @staticmethod
    def _identity(update):
        if not isinstance(update, dict):
            return None
        kind = update.get("update_type")
        message = update.get("message") or {}
        if not isinstance(message, dict) or not isinstance(update.get("callback", {}), dict):
            return None
        if kind in {"bot_started", "bot_stopped"}:
            user, chat = update.get("user") or {}, update.get("chat_id")
        elif kind in {"message_created", "message_callback"}:
            recipient = message.get("recipient") or {}
            if not isinstance(recipient, dict):
                return None
            if recipient.get("chat_type") != "dialog":
                return None
            user = (update.get("callback") or {}).get("user") if kind == "message_callback" else message.get("sender")
            user, chat = user or {}, recipient.get("chat_id")
        else:
            return None
        if not isinstance(user, dict):
            return None
        uid = user.get("user_id")
        if user.get("is_bot") or type(uid) is not int or uid <= 0 or type(chat) is not int or chat == 0:
            return None
        return uid, chat

    @classmethod
    def event_key(cls, update):
        if not isinstance(update, dict) or not cls._identity(update):
            return None
        kind = update.get("update_type")
        if kind == "message_created":
            body = (update.get("message") or {}).get("body") or {}
            if not isinstance(body, dict):
                return None
            identifier = body.get("mid")
        elif kind == "message_callback":
            identifier = (update.get("callback") or {}).get("callback_id")
        else:
            if type(update.get("timestamp")) is not int or update["timestamp"] <= 0:
                return None
            identifier = str(update.get("timestamp"))
        if not isinstance(identifier, str) or not identifier or len(identifier) > 256:
            return None
        return hashlib.sha256(f"{kind}:{cls._identity(update)}:{identifier}".encode()).hexdigest()

    def _button(self, label, identity, account, action):
        now = self.clock()
        while self.actions and (next(iter(self.actions.values()))[0] <= now or len(self.actions) >= 5000):
            self.actions.popitem(last=False)
        key = secrets.token_urlsafe(18)
        self.actions[key] = (now + 900, identity, account["id"], action)
        return [{"type": "callback", "text": text(label, 96), "payload": key}]

    def _reply(self, update, identity, value, buttons=None):
        body = {"text": value[:4000]}
        if buttons:
            body["attachments"] = [{"type": "inline_keyboard", "payload": {"buttons": buttons}}]
        reply = {"user_id": identity[0], "body": body}
        if update.get("update_type") == "message_callback":
            reply["callback_id"] = update["callback"]["callback_id"]
        return reply

    def handle(self, update):
        identity = self._identity(update) if isinstance(update, dict) else None
        if not identity:
            return None
        uid, chat = identity
        try:
            kind = update["update_type"]
            body = (update.get("message") or {}).get("body") or {}
            if not isinstance(body, dict):
                return None
            command = body.get("text") or ""
            if not isinstance(command, str):
                return None
            if kind == "bot_stopped" or (kind == "message_created" and command.strip() == "/stop"):
                self.store.unlink_max(uid, chat)
                self.actions = OrderedDict((k, v) for k, v in self.actions.items() if v[1] != identity)
                return None if kind == "bot_stopped" else self._reply(update, identity, "Доступ MAX отключён. Для повторного подключения получите новый код в Нейроне.")
            self.store.rate_limit("bot-actions", str(uid), 30, 60)
            if kind == "bot_started" or command.strip() in {"/start", "/help"}:
                return self._reply(update, identity, HELP)
            if kind == "message_created" and command.strip().startswith("/link"):
                parts = command.strip().split()
                if len(parts) != 2 or parts[0] != "/link":
                    return self._reply(update, identity, "Отправьте /link КОД. Код выдаётся на странице «Подключить MAX» в Нейроне.")
                self.store.redeem_link(parts[1], uid, chat)
                self.store.audit("link_redeem", "ok")
                return self._reply(update, identity, f"Код принят. Аккаунт MAX: •••• {str(uid)[-4:]}.\nВернитесь в Нейрону, сверьте эти цифры и нажмите «Подтвердить». До подтверждения данные недоступны.")
            account = self.store.authorize(uid, chat)
            if kind == "message_callback":
                key = (update.get("callback") or {}).get("payload")
                saved = self.actions.get(key) if isinstance(key, str) else None
                if not saved or saved[0] <= self.clock() or saved[1] != identity or saved[2] != account["id"]:
                    return self._reply(update, identity, "Эта кнопка устарела. Отправьте поисковый запрос ещё раз.")
                action = saved[3]
                if action["kind"] == "contacts":
                    return self._contacts(update, identity, account, action)
                query, page = action["query"], action["page"]
            else:
                query, page = command.strip(), 1
            if not query or len(query) > 500:
                return self._reply(update, identity, "Напишите название и характеристики позиции, до 500 символов. Например: «Нужна муфта…».")
            return self._search(update, identity, account, query, page)
        except (StoreError, StockError) as exc:
            return self._reply(update, identity, exc.message)

    def _search(self, update, identity, account, query, page):
        result = self.stock.search(query, page=page, page_size=5)
        self.store.audit("search", "ok")
        groups = result["queries"]
        if len(groups) > 1:
            buttons = [self._button(group["query"], identity, account,
                       {"kind": "search", "query": group["query"], "page": 1}) for group in groups]
            return self._reply(update, identity, "Выберите позицию для отдельной проверки. Наличие всего комплекта в одном месте не подтверждено.", buttons)
        group = groups[0]
        lines = [f"Поиск: {text(group['query'])}"]
        if group["status"] == "clarify":
            lines.append("Уточните название, диаметр или модель. Найденные варианты не считаются точным совпадением.")
        if not group["items"]:
            lines.append("Подходящих опубликованных позиций не найдено.")
        buttons = []
        for index, item in enumerate(group["items"], 1):
            match = " · похожее совпадение" if item["match_type"] == "similar" else ""
            quantity = "не указано" if item["quantity"] is None else str(item["quantity"])
            lines.append(f"\n{index}. {text(item['name'])}{match}\n"
                         f"{text(item['municipality'], 80)} · {text(item['organization'], 110)}\n"
                         f"Количество: {text(quantity, 40)} {text(item['unit'], 24)} · дата: {text(item['stock_date'] or 'не указана', 32)}\n"
                         f"{text(item['availability_label'], 150)}")
            buttons.append(self._button(f"{index}. {item['municipality']} · контакт", identity, account,
                           {"kind": "contacts", "item_id": item["id"], "municipality_id": item["municipality_id"], "page": 1}))
        pagination = group["pagination"]
        if pagination["page"] > 1:
            buttons.append(self._button("← Назад", identity, account, {"kind": "search", "query": query, "page": page - 1}))
        if pagination["has_next"]:
            buttons.append(self._button("Следующие →", identity, account, {"kind": "search", "query": query, "page": page + 1}))
        if group["items"]:
            lines.append(f"\nСтраница {pagination['page']}. Совпадений: {pagination['total']}.")
        for warning in (result.get("warnings", []) + group.get("warnings", []))[:3]:
            lines.append(text(warning, 160))
        return self._reply(update, identity, "\n".join(lines), buttons)

    def _contacts(self, update, identity, account, action):
        result = self.stock.contacts(action["item_id"], action["municipality_id"])
        self.store.audit("contacts", "ok")
        lines = [f"{text(result['item_name'])}\n{text(result['municipality'], 80)} · {text(result['organization'], 120)}"]
        contacts = result["contacts"]
        start = (action["page"] - 1) * 5
        for contact in contacts[start:start + 5]:
            lines.append(f"\n{text(contact['person'], 140)}\n{text(contact['position'], 120)}\nТелефон: {text(contact['phone'] or 'не указан', 120)}")
        if not contacts:
            lines.append("Контакт ответственного не указан в системе.")
        lines.extend(text(warning, 160) for warning in result.get("warnings", [])[:3])
        buttons = []
        if start > 0:
            buttons.append(self._button("← Контакты", identity, account, {**action, "page": action["page"] - 1}))
        if len(contacts) > start + 5:
            buttons.append(self._button("Ещё контакты →", identity, account, {**action, "page": action["page"] + 1}))
        return self._reply(update, identity, "\n".join(lines), buttons)
