"""Read-only MAX projection of the curator's existing published stock.

No ingestion, curator state, independent inventory database, network or LLM is
used here. Authentication and scope checks belong to the restricted MAX router.
Dates of publication/import never replace the report's stock date.
"""
from __future__ import annotations

import hashlib
import json
import math
import re
import unicodedata
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from threading import Lock
from zoneinfo import ZoneInfo

from services.report_municipalities import is_municipality, key as municipality_key


MAX_STOCK_BYTES = 32 * 1024 * 1024
MAX_CONTACT_BYTES = 8 * 1024 * 1024
MAX_ROWS = 100_000
MAX_CONTACTS = 30_000
MAX_QUERY_LENGTH = 1200
MAX_PART_LENGTH = 240
MAX_PARTS = 5
MAX_QUERY_TOKENS = 24
MAX_PAGE_SIZE = 50

_INTENT = frozenset({
    "нужен", "нужна", "нужно", "нужны", "нужную", "нужный", "нужное", "мне", "нам",
    "где", "есть", "найди", "найти", "найдите", "поиск", "ищу", "ищем", "покажи",
    "покажите", "пожалуйста", "наличие", "наличии", "в", "на", "складе", "у", "кого",
})
# Only inflections of specific nouns, not equivalence of different fittings.
_WORDS = {
    "муфту": "муфта", "муфты": "муфта", "муфт": "муфта",
    "насоса": "насос", "насосы": "насос", "насосов": "насос",
    "трубу": "труба", "трубы": "труба", "труб": "труба",
    "задвижку": "задвижка", "задвижки": "задвижка", "задвижек": "задвижка",
    "метры": "м", "метр": "м", "метров": "м", "штук": "шт", "штуки": "шт",
}
_HEADER = {"рсо": "organization", "округ": "municipality", "наименование": "name",
           "категория": "category", "группа": "group", "количество": "quantity",
           "ед.": "unit", "дата": "date"}
_REQUIRED = {"organization", "municipality", "name", "quantity", "unit"}
_MISSING_UNITS = {"", "—", "–", "-", "не указано", "неизвестно", "ед?", "null", "none"}
_STATUS_LABELS = {
    "available": "Есть по актуальным данным",
    "zero": "Нулевой остаток",
    "unknown": "Наличие не подтверждено",
    "stale": "Данные устарели — уточните остаток",
}


class StockError(ValueError):
    """Safe public error; never carries paths, queries, contacts or source errors."""

    def __init__(self, code: str, message: str, status_code: int = 503):
        super().__init__(message)
        self.code = code
        self.message = message
        self.status_code = status_code


def _invalid_query(message="Уточните запрос: укажите название или характеристику позиции."):
    return StockError("INVALID_QUERY", message, 400)


def _text(value, maximum=1000):
    if not isinstance(value, str) or len(value) > maximum:
        return ""
    # Values remain text, never markup. Control characters do not reach the UI.
    return re.sub(r"\s+", " ", "".join(c for c in value if c in "\r\n\t" or unicodedata.category(c) not in {"Cc", "Cs"})).strip()


def _plain(value):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", value).casefold().replace("ё", "е")).strip()


def _org_key(value):
    # Keep legal form and significant digits. Do not fuzzy-match organizations.
    return _plain(value).translate(str.maketrans({"«": '"', "»": '"', "„": '"', "“": '"', "”": '"'}))


def _tokens(value, *, query=False):
    normalized = _plain(value)
    # Normalize dimension separators only BETWEEN digits: model letters stay.
    normalized = re.sub(r"(?<=\d)\s*[xх×*]\s*(?=\d)", "x", normalized)
    normalized = re.sub(r"(?<=\d),(?=\d)", ".", normalized)
    # Keep whole model/article/dimension tokens, including their order. K50-32
    # must not become interchangeable with K32-50 or a different pump code.
    tokens = re.findall(r"(?=[a-zа-я0-9./-]*\d)[a-zа-я0-9]+(?:[./-][a-zа-я0-9]+)*|[a-zа-я]+", normalized)
    return tuple(_WORDS.get(token, token) for token in tokens if not query or token not in _INTENT)


def _search_tokens(tokens):
    values = set(tokens)
    # A plain diameter query (100) can find D100, but a query containing D100
    # still requires that complete code. Numeric prefixes are never substrings.
    for token in tokens:
        if re.search(r"\d", token):
            values.update(re.findall(r"\d+(?:\.\d+)?|[a-zа-я]+", token))
    return frozenset(values)


def _opaque(prefix, *parts):
    packed = json.dumps(parts, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return prefix + hashlib.sha256(packed).hexdigest()[:32]


def _quantity(value):
    if value is None or isinstance(value, bool) or not isinstance(value, (str, int, float)):
        return None
    if isinstance(value, str):
        value = re.sub(r"\s", "", value).replace(",", ".")
        if not re.fullmatch(r"[+]?(?:\d+(?:\.\d*)?|\.\d+)", value):
            return None
    try:
        number = float(value)
        return number if math.isfinite(number) and number >= 0 else None
    except (ValueError, TypeError, OverflowError):
        return None


def _date(value):
    if not isinstance(value, str):
        return None
    raw = value.strip()
    try:
        if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", raw):
            return datetime.strptime(raw, "%d.%m.%Y").date()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
            return date.fromisoformat(raw)
    except ValueError:
        pass
    return None


def _stamp(value):
    raw = _text(value, 64)
    try:
        datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except (ValueError, TypeError):
        return None
    return raw


def _object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate key")
        result[key] = value
    return result


def _reject_constant(value):
    raise ValueError("nonfinite JSON")


def _read_json(path, maximum, kind):
    try:
        if path.stat().st_size > maximum:
            raise ValueError("too large")
        with path.open("rb") as stream:
            raw = stream.read(maximum + 1)
        if len(raw) > maximum:
            raise ValueError("too large")
        result = json.loads(raw.decode("utf-8-sig"), object_pairs_hook=_object, parse_constant=_reject_constant)
        if not isinstance(result, dict):
            raise ValueError("object required")
        return result
    except FileNotFoundError:
        if kind == "CONTACTS":
            raise StockError("CONTACTS_NOT_CONFIGURED",
                "В Нейроне, к которой подключён этот бот, справочник контактов ещё не загружен. "
                "Загрузите его на том же компьютере или сервере: «Остатки ЗиП РСО» → "
                "«Контакты ответственных от РСО» → «Загрузить контакты».", 409) from None
        raise StockError("STOCK_UNAVAILABLE", "Опубликованные остатки временно недоступны. Попробуйте позже.") from None
    except (OSError, UnicodeError, ValueError, RecursionError, OverflowError):
        message = ("Справочник контактов временно недоступен. Попробуйте позже."
                   if kind == "CONTACTS" else "Опубликованные остатки временно недоступны. Попробуйте позже.")
        raise StockError(kind + "_UNAVAILABLE", message) from None


def _one_typo(left, right):
    """One edit/transposition for words only; never relax model numbers/dimensions."""
    if left == right:
        return True
    if min(len(left), len(right)) < 4 or not left.isalpha() or not right.isalpha() or abs(len(left) - len(right)) > 1:
        return False
    if len(left) == len(right):
        wrong = [i for i, (a, b) in enumerate(zip(left, right)) if a != b]
        return len(wrong) == 1 or (len(wrong) == 2 and wrong[1] == wrong[0] + 1
                                  and left[wrong[0]] == right[wrong[1]] and left[wrong[1]] == right[wrong[0]])
    short, long = (left, right) if len(left) < len(right) else (right, left)
    first = next((i for i, (a, b) in enumerate(zip(short, long)) if a != b), len(short))
    return short[first:] == long[first + 1:]


class StockService:
    def __init__(self, published_path, contacts_path, stale_days=30):
        if type(stale_days) is not int or not 1 <= stale_days <= 365:
            raise ValueError("stale_days must be 1..365")
        self.published_path = Path(published_path)
        self.contacts_path = Path(contacts_path)
        self.stale_days = stale_days
        self._cache_key = None
        self._cache = None
        self._cache_lock = Lock()

    def _snapshot(self):
        # Cache only a bounded read-only projection; stat every request so a
        # withdrawal/correction is immediately applied, including contact calls.
        try:
            stat = self.published_path.stat()
            key = (stat.st_mtime_ns, stat.st_ctime_ns, stat.st_size, stat.st_ino)
        except OSError:
            raise StockError("STOCK_UNAVAILABLE", "Опубликованные остатки временно недоступны. Попробуйте позже.") from None
        with self._cache_lock:
            if self._cache_key == key and self._cache is not None:
                return self._cache
            payload = _read_json(self.published_path, MAX_STOCK_BYTES, "STOCK")
            grid = payload.get("rows")
            if not isinstance(grid, list) or not grid or len(grid) > MAX_ROWS + 1 or not isinstance(grid[0], list):
                raise StockError("STOCK_UNAVAILABLE", "Опубликованный свод остатков не готов. Попробуйте позже.")
            positions = {}
            for index, heading in enumerate(grid[0]):
                field = _HEADER.get(_plain(_text(heading)))
                if field:
                    if field in positions:
                        raise StockError("STOCK_UNAVAILABLE", "Не удалось прочитать опубликованный свод остатков.")
                    positions[field] = index
            if not _REQUIRED.issubset(positions):
                raise StockError("STOCK_UNAVAILABLE", "Не удалось прочитать опубликованный свод остатков.")
            rows, omitted = [], 0
            seen = set()
            organizations = defaultdict(set)
            for row in grid[1:]:
                if not isinstance(row, list) or len(row) <= max(positions.values()):
                    omitted += 1
                    continue
                value = lambda field: row[positions[field]] if field in positions else None
                entry = {field: _text(value(field)) for field in ("organization", "municipality", "name", "category", "group", "unit")}
                if not entry["organization"] or not entry["name"]:
                    omitted += 1
                    continue
                muni_key = municipality_key(entry["municipality"]) if is_municipality(entry["municipality"]) else ""
                entry["municipality_id"] = _opaque("m_", muni_key)
                entry["_municipality_key"] = muni_key
                entry["_org_key"] = _org_key(entry["organization"])
                # Existing publications have no item IDs. Include the complete
                # identity and quantity/date version to reject withdrawn results.
                entry["quantity"] = _quantity(value("quantity"))
                stock_date = _date(value("date"))
                entry["stock_date"] = stock_date.isoformat() if stock_date else None
                entry["_date"] = stock_date
                entry["id"] = _opaque("i_", entry["_org_key"], muni_key, entry["name"], entry["category"], entry["group"],
                                      entry["unit"], entry["quantity"], entry["stock_date"])
                if entry["id"] in seen:
                    continue  # identical published copies do not multiply stock
                seen.add(entry["id"])
                entry["_tokens"] = _tokens(entry["name"])
                entry["_search_tokens"] = _search_tokens((*entry["_tokens"], *_tokens(entry["unit"])))
                rows.append(entry)
                organizations[entry["_org_key"]].add(muni_key)
            if grid[1:] and not rows:
                raise StockError("STOCK_UNAVAILABLE", "Не удалось прочитать позиции опубликованного свода.")
            result = {"rows": rows, "organizations": organizations,
                      "published_at": _stamp(payload.get("published_at")), "omitted": omitted}
            self._cache_key, self._cache = key, result
            return result

    def _availability(self, row, today):
        qty = row["quantity"]
        stock_date = row["_date"]
        if qty == 0:
            return "zero"
        if qty is None or _plain(row["unit"]) in _MISSING_UNITS or not row["_municipality_key"] or stock_date is None or stock_date > today:
            return "unknown"
        return "stale" if (today - stock_date).days > self.stale_days else "available"

    @staticmethod
    def _queries(query):
        if (not isinstance(query, str) or len(query) > MAX_QUERY_LENGTH
                or any(unicodedata.category(c) == "Cs" or ord(c) < 32 and c not in "\n\r\t" for c in query)):
            raise _invalid_query("Запрос слишком длинный или содержит недопустимые символы.")
        parts = [_text(part, MAX_PART_LENGTH) for part in re.split(r"[;\r\n]+", query) if part.strip()]
        if not parts or len(parts) > MAX_PARTS or any(not part for part in parts):
            raise _invalid_query("Введите от одной до пяти позиций; разделяйте их новой строкой или точкой с запятой.")
        result = []
        for part in parts:
            tokens = _tokens(part, query=True)
            if not tokens or len(tokens) > MAX_QUERY_TOKENS or not any(len(t) >= 2 for t in tokens):
                raise _invalid_query()
            result.append((part, tokens))
        return result

    def search(self, query: str, page: int = 1, page_size: int = 20, today=None):
        parts = self._queries(query)
        if type(page) is not int or not 1 <= page <= MAX_ROWS or type(page_size) is not int or not 1 <= page_size <= MAX_PAGE_SIZE:
            raise _invalid_query("Некорректные параметры страницы результатов.")
        today = today or datetime.now(ZoneInfo("Europe/Moscow")).date()
        if isinstance(today, datetime):
            today = today.date()
        if not isinstance(today, date):
            raise ValueError("today must be date")
        snapshot = self._snapshot()
        queries = []
        for part, tokens in parts:
            matches = []
            for row in snapshot["rows"]:
                available = row["_search_tokens"]
                if all(token in available for token in tokens):
                    match_type = "exact" if tuple(tokens) == row["_tokens"] else "terms"
                    matches.append((row, match_type))
            if not matches:
                # Suggestions are offered only when all requested numeric/model
                # tokens are preserved. They are never represented as exact.
                for row in snapshot["rows"]:
                    available = row["_search_tokens"]
                    missing = [token for token in tokens if token not in available]
                    if len(missing) == 1 and any(_one_typo(missing[0], word) for word in available):
                        matches.append((row, "similar"))
            matches.sort(key=lambda pair: (
                {"exact": 0, "terms": 1, "similar": 2}[pair[1]],
                {"available": 0, "stale": 1, "zero": 2, "unknown": 3}[self._availability(pair[0], today)],
                _plain(pair[0]["name"]), pair[0]["_municipality_key"], pair[0]["_org_key"], pair[0]["id"],
            ))
            total = len(matches)
            variants = {(tuple(row["_tokens"]), row["unit"]) for row, _ in matches}
            ambiguous = len(variants) > 1 or any(kind == "similar" for _, kind in matches)
            items = []
            start = (page - 1) * page_size
            for row, kind in matches[start:start + page_size]:
                item = {key: row[key] for key in ("id", "name", "category", "group", "quantity", "unit",
                                                "municipality_id", "municipality", "organization", "stock_date")}
                item["availability"] = self._availability(row, today)
                item["availability_label"] = _STATUS_LABELS[item["availability"]]
                item["match_type"] = kind
                items.append(item)
            queries.append({"query": part, "status": "empty" if not total else "clarify" if ambiguous else "ok",
                            "items": items, "pagination": {"page": page, "page_size": page_size, "total": total,
                                                           "has_next": start + page_size < total},
                            "warnings": ["Уточните характеристики и выберите конкретную позицию."] if ambiguous else []})
        warnings = []
        if snapshot["omitted"]:
            warnings.append("Часть строк источника неполна и не включена в выдачу.")
        if len(parts) > 1:
            warnings.append("Результаты показаны отдельно по каждой позиции; наличие всего комплекта в одном месте не подтверждено.")
        return {"query": query.strip(), "page": page, "page_size": page_size, "queries": queries,
                "published_at": snapshot["published_at"], "policy": {"stale_days": self.stale_days}, "warnings": warnings}

    def contacts(self, item_id, municipality_id):
        if (not isinstance(item_id, str) or not re.fullmatch(r"i_[0-9a-f]{32}", item_id)
                or not isinstance(municipality_id, str) or not re.fullmatch(r"m_[0-9a-f]{32}", municipality_id)):
            raise StockError("RESULT_NOT_FOUND", "Позиция больше недоступна. Обновите поиск.", 404)
        snapshot = self._snapshot()
        row = next((item for item in snapshot["rows"] if item["id"] == item_id and item["municipality_id"] == municipality_id), None)
        if row is None or not row["_municipality_key"]:
            raise StockError("RESULT_NOT_FOUND", "Позиция больше недоступна. Обновите поиск.", 404)
        payload = _read_json(self.contacts_path, MAX_CONTACT_BYTES, "CONTACTS")
        source = payload.get("items")
        if not isinstance(source, list) or len(source) > MAX_CONTACTS or any(not isinstance(item, dict) for item in source):
            raise StockError("CONTACTS_UNAVAILABLE", "Справочник контактов временно недоступен. Попробуйте позже.")
        contacts, warnings, seen = [], [], set()
        ambiguous = False
        for item in source:
            organization = _text(item.get("rso"))
            if not organization or _org_key(organization) != row["_org_key"]:
                continue
            municipality = municipality_key(_text(item.get("municipality")))
            if municipality and municipality != row["_municipality_key"]:
                continue
            if not municipality and snapshot["organizations"][row["_org_key"]] != {row["_municipality_key"]}:
                ambiguous = True
                continue
            person, position = _text(item.get("person"), 300), _text(item.get("position"), 300)
            phone = _text(item.get("phone"), 160)
            # Old imports sometimes stored a person's name in the phone column.
            # Never invent or substitute a switchboard number in that case.
            if not re.search(r"\d", phone) or re.search(r"[<>@]", phone):
                phone = None
            identity = (person, position, phone, organization)
            if identity in seen:
                continue
            seen.add(identity)
            contacts.append({"id": _opaque("c_", row["_org_key"], row["_municipality_key"], *identity),
                             "person": person, "position": position, "organization": organization, "phone": phone})
        contacts.sort(key=lambda item: (_plain(item["position"]), _plain(item["person"]), item["id"]))
        if ambiguous:
            warnings.append("Для части контактов не указан муниципалитет; принадлежность требует уточнения у администратора.")
        if not contacts:
            if not source:
                warnings.append("Справочник контактов этой Нейроны пуст. Загрузите контакты в разделе «Остатки ЗиП РСО» → «Контакты ответственных от РСО».")
            else:
                warnings.append("Контакт ответственного для этой организации и муниципалитета не найден.")
        elif any(not item["phone"] for item in contacts):
            warnings.append("Для части ответственных номер телефона не указан в системе.")
        if len(contacts) > 50:
            warnings.append("Показаны первые 50 контактов организации. Уточните ответственного в системе.")
        return {"item_id": item_id, "municipality_id": municipality_id, "municipality": row["municipality"],
                "organization": row["organization"], "item_name": row["name"], "contacts": contacts[:50],
                "updated_at": _stamp(payload.get("updated_at")), "warnings": warnings}
