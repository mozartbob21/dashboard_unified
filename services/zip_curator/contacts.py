"""Persistent contacts of responsible representatives for the ZИП dashboard."""
from __future__ import annotations

import csv
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from services.zip_curator.core import DATA, norm, read_xlsx_rows


CONTACTS_FILE = DATA / "contacts.json"
IMPORT_VERSION = 2

FIELD_ALIASES = {
    "rso": (
        "рсо",
        "rso",
        "название рсо",
        "наименование рсо",
        "организация",
        "предприятие",
        "наименование организации",
    ),
    "municipality": ("municipality", "муниципалитет", "омсу", "округ", "городской округ"),
    "person": ("person", "name", "фио", "ф и о", "ответственный", "контактное лицо", "представитель"),
    "position": ("position", "должность", "роль"),
    "phone": ("phone", "mobile", "телефон", "мобильный", "мобильный телефон", "номер телефона", "тел."),
    "email": ("электронная почта", "email", "e-mail", "почта"),
    "note": ("note", "примечание", "комментарий", "заметка"),
}

BASE_FIELDS = {"rso", "municipality", "note"}
CONTACT_FIELDS = {"person", "position", "phone", "email"}
ROLE_PATTERNS = (
    ("chief_engineer", "Главный инженер", re.compile(r"\b(?:главн\w*|гл)\s+инженер\w*\b")),
    ("manager", "Руководитель", re.compile(r"\b(?:руководител\w*|директор\w*|начальник\w*)\b")),
    ("responsible", "Ответственный", re.compile(r"\b(?:ответственн\w*|представител\w*)\b")),
)


def _atomic_write(payload):
    CONTACTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp = CONTACTS_FILE.with_name(CONTACTS_FILE.name + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, CONTACTS_FILE)


def load_contacts():
    if CONTACTS_FILE.exists():
        try:
            payload = json.loads(CONTACTS_FILE.read_text(encoding="utf-8"))
            if isinstance(payload, dict) and isinstance(payload.get("items"), list):
                payload.setdefault("updated_at", None)
                payload.setdefault("source_file", None)
                payload["items"] = _prepare_stored_contacts(payload["items"])
                payload["needs_reimport"] = (
                    payload.get("import_version", 0) < IMPORT_VERSION
                    and any(not item.get("phone") for item in payload["items"])
                )
                return payload
        except (OSError, json.JSONDecodeError):
            pass
    return {"items": [], "updated_at": None, "source_file": None}


def _name_in_phone(value):
    """Recognize the old header collision, without treating short phone numbers as names."""
    text = str(value or "").strip()
    return not re.search(r"[\d@]", text) and len(re.findall(r"[а-яёa-z]+", text, re.I)) >= 2


def _normalize_contact(contact):
    contact = dict(contact)
    person = re.sub(r"\s+", " ", str(contact.get("person") or "")).strip()
    # Only collapse literal repetition; different people's names are never discarded.
    words = person.split()
    for size in range(1, len(words) // 2 + 1):
        if len(words) % size == 0 and words == words[:size] * (len(words) // size):
            person = " ".join(words[:size])
            break
    role_prefix = re.match(
        r"^(директор\s+филиала|генеральный\s+директор|директор|руководитель|главный\s+инженер)\s*[:—–-]?\s+(.+)$",
        person, re.I,
    )
    if role_prefix:
        role, person = role_prefix.groups()
        if not contact.get("position") or norm(contact["position"]) in {"ответственный", "руководитель"}:
            contact["position"] = role[0].upper() + role[1:]
    contact["person"] = person
    return contact


def _prepare_stored_contacts(items):
    prepared = []
    for original in items:
        item = dict(original)
        phone = str(item.get("phone") or "").strip()
        if _name_in_phone(phone):
            item["phone"] = ""
            if not item.get("person"):
                item["person"] = phone
            elif norm(item["person"]) != norm(phone):
                # The old parser could save the engineer as person and the manager
                # as phone. Preserve both names without inventing missing numbers.
                recovered = {**item, "id": uuid.uuid5(uuid.NAMESPACE_URL, str(item.get("id")) + phone).hex,
                             "person": phone, "position": "", "email": ""}
                prepared.append(_normalize_contact(recovered))
        prepared.append(_normalize_contact(item))
    return _deduplicate_contacts(prepared)


def _deduplicate_contacts(items):
    result = []
    for item in items:
        # Deduplicate identical records only; shared switchboard numbers are not identity.
        existing = next((other for other in result if all(
            norm(other.get(key)) == norm(item.get(key))
            for key in FIELD_ALIASES
        )), None)
        if existing is None:
            result.append(item)
    return result


def _header(value):
    value = norm(value)
    return re.sub(r"[^a-zа-яё0-9]+", " ", value).strip()


def _alias_in_header(header, alias):
    """Match whole words/phrases, never fragments inside unrelated words."""
    alias = _header(alias)
    if not alias:
        return False
    return re.search(rf"(?:^|\s){re.escape(alias)}(?:$|\s)", header) is not None


def _field_for_header(value):
    header = _header(value)
    # Prefer exact matches first. The word "РСО" may also occur in a person
    # column such as "Ответственный от РСО".
    for field, aliases in FIELD_ALIASES.items():
        if any(header == _header(alias) for alias in aliases):
            return field

    # Contact types are determined before base fields so that a header such as
    # "ФИО ответственного от РСО" cannot be mistaken for the organization.
    if re.search(r"(?:^|\s)ф\s*и\s*о(?:$|\s)", header):
        return "person"
    if any(_alias_in_header(header, alias) for alias in FIELD_ALIASES["email"]):
        return "email"
    if any(_alias_in_header(header, alias) for alias in FIELD_ALIASES["phone"]):
        return "phone"
    if any(_alias_in_header(header, alias) for alias in FIELD_ALIASES["person"]):
        return "person"
    if any(_alias_in_header(header, alias) for alias in FIELD_ALIASES["position"]):
        return "position"

    matches = []
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            normalized_alias = _header(alias)
            if normalized_alias != "рсо" and _alias_in_header(header, normalized_alias):
                matches.append((len(normalized_alias), field))
    if matches:
        return max(matches, key=lambda match: match[0])[1]
    return None


def _role_for_header(value):
    header = _header(value)
    for key, label, pattern in ROLE_PATTERNS:
        if pattern.search(header):
            return key, label
    number = re.search(r"(?:^|\s)(\d{1,2})(?:$|\s)", header)
    if number:
        return f"contact_{number.group(1)}", f"Контакт {number.group(1)}"
    return "contact", "Ответственный"


def _rows_to_contacts(rows):
    header_index = -1
    base_columns = {}
    contact_columns = {}
    role_labels = {}
    for index, row in enumerate(rows[:15]):
        candidate_base = {}
        candidate_contacts = {}
        candidate_labels = {}
        for column, value in enumerate(row or []):
            field = _field_for_header(value)
            if field in BASE_FIELDS and field not in candidate_base:
                candidate_base[field] = column
            elif field in CONTACT_FIELDS:
                role, label = _role_for_header(value)
                candidate_contacts.setdefault(role, {})[field] = column
                candidate_labels[role] = label
        has_contact = any(
            any(field in group for field in ("person", "phone", "email"))
            for group in candidate_contacts.values()
        )
        if "rso" in candidate_base and has_contact:
            header_index = index
            base_columns = candidate_base
            contact_columns = candidate_contacts
            role_labels = candidate_labels
            break
    if header_index < 0:
        raise ValueError("Не найдены колонки РСО и контакта (ФИО, телефон или почта)")

    # Если в таблице одна именованная роль, а второй столбец назван просто
    # "Телефон" или "ФИО", дополняем этой общей колонкой найденную роль.
    named_roles = [role for role in contact_columns if role != "contact"]
    generic = contact_columns.get("contact")
    if generic and len(named_roles) == 1 and not (set(generic) & set(contact_columns[named_roles[0]])):
        target = contact_columns[named_roles[0]]
        for field, column in generic.items():
            target.setdefault(field, column)
        del contact_columns["contact"]
        role_labels.pop("contact", None)

    contacts = []
    for row in rows[header_index + 1:]:
        row = row or []
        base = {
            field: str(row[column] if column < len(row) and row[column] is not None else "").strip()
            for field, column in base_columns.items()
        }
        if not base.get("rso"):
            continue
        for role, columns in contact_columns.items():
            contact = dict(base)
            for field, column in columns.items():
                contact[field] = str(row[column] if column < len(row) and row[column] is not None else "").strip()
            if any(contact.get(key) for key in ("person", "phone", "email")):
                contact["position"] = contact.get("position") or role_labels.get(role) or "Ответственный"
                contacts.append(_normalize_contact(contact))
    return contacts


def _json_to_contacts(data):
    if isinstance(data, dict):
        data = data.get("items") or data.get("contacts") or data.get("data") or []
    if not isinstance(data, list):
        raise ValueError("JSON должен содержать массив контактов")
    if not data:
        return []
    if not all(isinstance(item, dict) for item in data):
        raise ValueError("Каждый контакт в JSON должен быть объектом")
    headers = list(dict.fromkeys(key for item in data for key in item.keys()))
    rows = [headers] + [[item.get(key) for key in headers] for item in data]
    return _rows_to_contacts(rows)


def _decode_text(content, kind="текстового файла"):
    encodings = ("utf-16", "utf-8-sig", "cp1251") if content.startswith((b"\xff\xfe", b"\xfe\xff")) else ("utf-8-sig", "cp1251")
    for encoding in encodings:
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Не удалось определить кодировку {kind}")


def _text_to_rows(text):
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=";,\t|")
        return list(csv.reader(io.StringIO(text), dialect))
    except csv.Error:
        return [[line.strip()] for line in text.splitlines() if line.strip()]


def _read_docx_rows(content):
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            document = ET.fromstring(archive.read("word/document.xml"))
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as error:
        raise ValueError("Не удалось прочитать документ .docx") from error

    rows = []
    for table in document.findall(".//w:tbl", namespace):
        for row in table.findall("./w:tr", namespace):
            values = []
            for cell in row.findall("./w:tc", namespace):
                values.append(" ".join(
                    value.strip()
                    for value in (node.text or "" for node in cell.findall(".//w:t", namespace))
                    if value.strip()
                ))
            if any(values):
                rows.append(values)
    if rows:
        return rows

    paragraphs = []
    for paragraph in document.findall(".//w:p", namespace):
        value = " ".join(
            text.strip()
            for text in (node.text or "" for node in paragraph.findall(".//w:t", namespace))
            if text.strip()
        )
        if value:
            paragraphs.append(value)
    return _text_to_rows("\n".join(paragraphs))


def _read_legacy_doc(content):
    converters = []
    if shutil.which("textutil"):
        converters.append(lambda path: ["textutil", "-convert", "txt", "-stdout", str(path)])
    if shutil.which("antiword"):
        converters.append(lambda path: ["antiword", str(path)])
    if shutil.which("catdoc"):
        converters.append(lambda path: ["catdoc", str(path)])
    if not converters:
        raise ValueError("Для старого .doc требуется системный конвертер textutil, antiword или catdoc")

    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "contacts.doc"
        path.write_bytes(content)
        for command in converters:
            try:
                result = subprocess.run(command(path), capture_output=True, check=True, timeout=20)
                if result.stdout:
                    return _text_to_rows(_decode_text(result.stdout, "документа .doc"))
            except (OSError, subprocess.SubprocessError, ValueError):
                continue
    raise ValueError("Не удалось прочитать документ .doc")


def parse_contacts(filename, content):
    suffix = Path(filename or "").suffix.lower()
    if suffix == ".json":
        return _json_to_contacts(json.loads(content.decode("utf-8-sig")))
    if suffix in {".csv", ".txt"}:
        return _rows_to_contacts(_text_to_rows(_decode_text(content, suffix.upper()[1:])))
    if suffix == ".xlsx":
        return _rows_to_contacts(read_xlsx_rows(content))
    if suffix == ".docx":
        return _rows_to_contacts(_read_docx_rows(content))
    if suffix == ".doc":
        return _rows_to_contacts(_read_legacy_doc(content))
    raise ValueError("Поддерживаются файлы .xlsx, .csv, .txt, .doc, .docx и .json")


def _phone_key(value):
    return re.sub(r"\D+", "", str(value or ""))


def _same_contact(left, right):
    if norm(left.get("rso")) != norm(right.get("rso")):
        return False
    if (left.get("municipality") and right.get("municipality")
            and norm(left["municipality"]) != norm(right["municipality"])):
        return False
    def position_key(value):
        if not value or norm(value) == "ответственный":
            return ""
        key, _ = _role_for_header(value)
        return norm(value) if key == "contact" else key

    left_position = position_key(left.get("position"))
    right_position = position_key(right.get("position"))
    if left_position and right_position and left_position != right_position:
        return False
    if left.get("person") and right.get("person"):
        return norm(left["person"]) == norm(right["person"])
    pairs = (
        (norm(left.get("person")), norm(right.get("person"))),
        (_phone_key(left.get("phone")), _phone_key(right.get("phone"))),
        (str(left.get("email") or "").strip().casefold(), str(right.get("email") or "").strip().casefold()),
    )
    return any(a and b and a == b for a, b in pairs)


def import_contacts(filename, content):
    incoming = parse_contacts(filename, content)
    if not incoming:
        raise ValueError("В файле не найдены контакты. Проверьте заголовки и строки таблицы.")
    payload = load_contacts()
    items = payload["items"]
    added = updated = 0
    now = datetime.now().isoformat(timespec="seconds")
    for contact in incoming:
        contact = {key: str(contact.get(key) or "").strip() for key in FIELD_ALIASES}
        existing = next((item for item in items if _same_contact(item, contact)), None)
        if existing:
            for key, value in contact.items():
                if value:
                    existing[key] = value
            existing["updated_at"] = now
            updated += 1
        else:
            contact.update({"id": uuid.uuid4().hex, "created_at": now, "updated_at": now})
            items.append(contact)
            added += 1
    items = _deduplicate_contacts(items)
    items.sort(key=lambda item: (norm(item.get("rso")), norm(item.get("position")), norm(item.get("person"))))
    payload.update({"items": items, "updated_at": now, "source_file": Path(filename or "").name,
                    "import_version": IMPORT_VERSION, "needs_reimport": False})
    if CONTACTS_FILE.exists():
        # Keep the original JSON before applying migration/import to legacy data.
        backup = CONTACTS_FILE.with_name(CONTACTS_FILE.name + ".before-import.bak")
        if not backup.exists():
            shutil.copy2(CONTACTS_FILE, backup)
    _atomic_write(payload)
    return {"added": added, "updated": updated, "total": len(items), "payload": payload}


def delete_contact(contact_id):
    payload = load_contacts()
    before = len(payload["items"])
    payload["items"] = [item for item in payload["items"] if item.get("id") != contact_id]
    if len(payload["items"]) == before:
        return False
    payload["updated_at"] = datetime.now().isoformat(timespec="seconds")
    _atomic_write(payload)
    return True
