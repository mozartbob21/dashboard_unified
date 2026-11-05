"""Read-only, bounded MAX view of the web dashboard's contact directory."""
from functools import lru_cache
from threading import Lock
import re

from services import report_municipalities as places
from services.zip_curator.contacts import prepare_stored_contacts
from .stock import (
    MAX_CONTACT_BYTES, MAX_CONTACTS, StockError, _company_identity,
    _opaque, _plain, _read_json, _stamp, _text,
)


_FIELDS = ("organization", "municipality", "person", "position", "phone")
_WORD = re.compile(r"[a-zа-я]+|\d+")
_COMMAND_WORDS = set("""
    контакт контакты контакта контактов контактные телефон телефона телефоны телефонов номер номера
    найти найди найдите покажи покажите нужен нужна нужны нужно ищу ищем подскажи дайте дай
    у в из во по для с со на об от к ко и где кто какой как мне нам или есть можно
    связаться связи позвонить звонить пожалуйста к кому кого человек человека
    рсо организации организация предприятия предприятие муниципалитет муниципалитета муниципалитете омсу
    округ округа округе округу городской городском городского город города городе
    ответственный ответственная ответственные ответственного ответственной ответственных
    подскажите узнать сотрудник сотрудника сотрудники представителя представители представителя
""".split())
_TOO_BROAD = {
    "ооо", "ао", "оао", "пао", "зао", "муп", "гуп", "мкп", "гбу", "г", "о",
    "руководитель", "руководителя", "директор", "директора", "инженер", "инженера",
    "главный", "главного", "главному", "дежурный", "дежурного", "служба", "службы",
}


def _unavailable():
    return StockError("CONTACTS_UNAVAILABLE", "Справочник контактов временно недоступен. Попробуйте позже.")


def _query_terms(query):
    if not isinstance(query, str) or len(query) > 300:
        raise StockError("INVALID_QUERY", "Укажите РСО, муниципалитет или фамилию ответственного.", 400)
    terms = tuple(dict.fromkeys(word for word in _WORD.findall(_plain(query)) if word not in _COMMAND_WORDS))
    if len(terms) > 16 or not any(len(term) >= 2 and term not in _TOO_BROAD for term in terms):
        raise StockError("INVALID_QUERY", "Укажите РСО, муниципалитет или фамилию ответственного.", 400)
    return terms


@lru_cache(maxsize=4096)
def _forms(word):
    return places.word_variants(word)


def _stat_key(path):
    st = path.stat()
    return (st.st_ino, st.st_size, st.st_mtime_ns, st.st_ctime_ns)


def _phone(value):
    phone = _text(value, 160)
    # Imported work numbers may contain extensions or multiple numbers. A name,
    # address, e-mail or arbitrary note in the phone column is not a phone.
    bare = re.sub(r"\b(?:доб(?:авочный)?|вн(?:утренний)?|тел(?:ефон)?|моб(?:ильный)?|раб(?:очий)?)\b\.?",
                  "", phone, flags=re.I)
    if not re.search(r"\d", bare) or not re.fullmatch(r"[\d\s+().,/:;#–—-]+", bare):
        return None
    return phone


class ContactDirectory:
    def __init__(self, path):
        self.path = path
        self._lock = Lock()
        self._key = self._data = None

    def _read(self):
        try:
            key = _stat_key(self.path)
        except FileNotFoundError:
            raise StockError("CONTACTS_NOT_CONFIGURED",
                             "В этой Нейроне справочник контактов ещё не загружен. Откройте Нейрону на том же компьютере или сервере, где работает бот, либо используйте общий каталог. Загрузите справочник: "
                             "«Остатки ЗиП РСО» → «Контакты ответственных от РСО» → «Загрузить контакты».", 409) from None
        except OSError:
            raise _unavailable() from None
        try:
            registry_key = (str(places.REGISTRY), _stat_key(places.REGISTRY))
        except OSError:
            registry_key = None
        key = (key, registry_key)
        with self._lock:
            if self._key != key:
                payload = _read_json(self.path, MAX_CONTACT_BYTES, "CONTACTS")
                raw = payload.get("items")
                if not isinstance(raw, list) or len(raw) > MAX_CONTACTS or any(not isinstance(row, dict) for row in raw):
                    raise _unavailable()
                try:
                    data = self._project(raw, _stamp(payload.get("updated_at")))
                except (AttributeError, TypeError, ValueError):
                    raise _unavailable() from None
                self._data, self._key = data, key
            return self._data

    @staticmethod
    def _project(raw, updated_at):
        aliases, _ = places._maps()
        by_city = {}
        for alias, canonical in aliases.items():
            by_city.setdefault(canonical, []).append(alias)
        entries = []
        seen = set()
        for row in prepare_stored_contacts(raw):
            org = _text(row.get("rso"))
            if not org:
                continue
            city = _text(row.get("municipality"))
            org_key = _company_identity(org)
            city_key = places.key(city)
            person = _text(row.get("person"), 300)
            position = _text(row.get("position"), 300)
            phone = _phone(row.get("phone"))
            identity = (org_key, city_key, _plain(person), _plain(position), phone)
            if identity in seen:
                continue
            seen.add(identity)
            text = ' '.join([org, org_key, person, city, city_key, position, *by_city.get(city_key, [])])
            words = _WORD.findall(_plain(text))
            forms = frozenset(form for word in words for form in _forms(word))
            entries.append(dict(organization=org, municipality=city, person=person, position=position, phone=phone,
                                _org=org_key, _city=city_key, _forms=forms))
        entries.sort(key=lambda c: (c["_city"], c["_org"], _plain(c["position"]), _plain(c["person"]), c["phone"] or ""))
        return entries, updated_at

    def search(self, query, *, limit=10, offset=0):
        terms = _query_terms(query)
        if type(limit) is not int or not 1 <= limit <= 50 or type(offset) is not int or not 0 <= offset <= MAX_CONTACTS:
            raise StockError("INVALID_QUERY", "Некорректная страница контактов.", 400)
        entries, updated_at = self._read()
        exact, partial = [], []
        for entry in entries:
            forms = entry["_forms"]
            missing = [term for term in terms if term not in forms]
            if not missing:
                exact.append(entry)
            elif all(term.isalpha() and len(term) >= 3 and any(w.startswith(term) for w in forms) for term in missing):
                partial.append(entry)
        matches = exact + partial
        items = [{field: entry[field] for field in _FIELDS} for entry in matches[offset:offset + limit]]
        warnings = [] if entries else ["Справочник контактов этой Нейроны пуст. Загрузите контакты в разделе «Остатки ЗиП РСО»."]
        return {"items": items, "total": len(matches), "updated_at": updated_at, "warnings": warnings}

    def for_item(self, item, organization_cities):
        entries, updated_at = self._read()
        contacts, warnings = [], []
        org_key = _company_identity(item["organization"])
        city_key = item["_municipality_key"]
        ambiguous = False
        for entry in entries:
            if entry["_org"] != org_key or (entry["_city"] and entry["_city"] != city_key):
                continue
            if not entry["_city"] and organization_cities != {city_key}:
                ambiguous = True
                continue
            contact = {field: entry[field] for field in _FIELDS if field != "municipality"}
            contact["id"] = _opaque("c_", item["_org_key"], city_key, contact["person"], contact["position"],
                                    contact["phone"], contact["organization"])
            contacts.append(contact)
        if ambiguous:
            warnings.append("Для части контактов не указан муниципалитет; принадлежность требует уточнения у администратора.")
        if not entries:
            warnings.append("Справочник контактов этой Нейроны пуст. Загрузите контакты в разделе «Остатки ЗиП РСО» → «Контакты ответственных от РСО».")
        elif not contacts:
            warnings.append("Контакт ответственного для этой организации и муниципалитета не найден.")
        elif any(not c["phone"] for c in contacts):
            warnings.append("Для части ответственных номер телефона не указан в системе.")
        if len(contacts) > 50:
            warnings.append("Показаны первые 50 контактов организации. Уточните ответственного в системе.")
        return dict(item_id=item["id"], municipality_id=item["municipality_id"], municipality=item["municipality"],
                    organization=item["organization"], item_name=item["name"], contacts=contacts[:50],
                    updated_at=updated_at, warnings=warnings)
