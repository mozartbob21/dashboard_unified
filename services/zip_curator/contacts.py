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

FIELD_ALIASES = {
    "rso": (
        "рсо",
        "название рсо",
        "наименование рсо",
        "организация",
        "предприятие",
        "наименование организации",
    ),
    "municipality": ("муниципалитет", "омсу", "округ", "городской округ"),
    "person": ("фио", "ответственный", "контактное лицо", "представитель"),
    "position": ("должность", "роль"),
    "phone": ("телефон", "мобильный", "номер телефона", "тел."),
    "email": ("электронная почта", "email", "e-mail", "почта"),
    "note": ("примечание", "комментарий", "заметка"),
}


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
                return payload
        except (OSError, json.JSONDecodeError):
            pass
    return {"items": [], "updated_at": None, "source_file": None}


def _header(value):
    value = norm(value)
    return re.sub(r"[^a-zа-яё0-9]+", " ", value).strip()


def _field_for_header(value):
    header = _header(value)
    # Prefer exact matches first. In particular, the word "РСО" may also occur
    # in a person column such as "Ответственный от РСО".
    for field, aliases in FIELD_ALIASES.items():
        if any(header == _header(alias) for alias in aliases):
            return field
    matches = []
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            normalized_alias = _header(alias)
            if normalized_alias != "рсо" and normalized_alias in header:
                matches.append((len(normalized_alias), field))
    if matches:
        return max(matches)[1]
    return None


def _rows_to_contacts(rows):
    header_index = -1
    columns = {}
    for index, row in enumerate(rows[:15]):
        candidate = {}
        for column, value in enumerate(row or []):
            field = _field_for_header(value)
            if field and field not in candidate:
                candidate[field] = column
        if "rso" in candidate and any(field in candidate for field in ("person", "phone", "email")):
            header_index, columns = index, candidate
            break
    if header_index < 0:
        raise ValueError("Не найдены колонки РСО и контакта (ФИО, телефон или почта)")

    contacts = []
    for row in rows[header_index + 1:]:
        row = row or []
        contact = {}
        for field, column in columns.items():
            contact[field] = str(row[column] if column < len(row) and row[column] is not None else "").strip()
        if contact.get("rso") and any(contact.get(key) for key in ("person", "phone", "email")):
            contacts.append(contact)
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
    pairs = (
        (norm(left.get("person")), norm(right.get("person"))),
        (_phone_key(left.get("phone")), _phone_key(right.get("phone"))),
        (str(left.get("email") or "").strip().casefold(), str(right.get("email") or "").strip().casefold()),
    )
    return any(a and b and a == b for a, b in pairs)


def import_contacts(filename, content):
    incoming = parse_contacts(filename, content)
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
    items.sort(key=lambda item: (norm(item.get("rso")), norm(item.get("person"))))
    payload.update({"items": items, "updated_at": now, "source_file": Path(filename or "").name})
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
