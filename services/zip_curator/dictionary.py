"""Persistent curator additions layered over the bundled classification dictionary."""
import copy
import json
import os
import re
import threading
from functools import lru_cache
from pathlib import Path

BASE_FILE = Path(__file__).with_name("zip_dict.json")
CUSTOM_FILE = Path(__file__).resolve().parents[2] / "data" / "zip_curator" / "dictionary.json"
LOCK = threading.RLock()


def _label(value, required=False):
    if value is not None and not isinstance(value, str):
        raise ValueError("Название должно быть текстом")
    value = re.sub(r"\s+", " ", value or "").strip()
    if required and not value:
        raise ValueError("Введите название категории")
    if len(value) > 120:
        raise ValueError("Название должно быть не длиннее 120 символов")
    return value


def _canonical(value, choices):
    return next((name for name in choices if name.casefold() == value.casefold()), value)


def _custom():
    if not CUSTOM_FILE.exists():
        return {"categories": [], "groupsByCat": {}, "dict": {}}
    return json.loads(CUSTOM_FILE.read_text(encoding="utf-8"))


@lru_cache(maxsize=1)
def _merged(base_path, custom_path, base_version, custom_version):
    result = json.loads(Path(base_path).read_text(encoding="utf-8"))
    custom = _custom()
    result["categories"] = sorted(set(result["categories"]) | set(custom["categories"]), key=str.casefold)
    for category, groups in custom["groupsByCat"].items():
        result["groupsByCat"][category] = sorted(set(result["groupsByCat"].get(category, [])) | set(groups), key=str.casefold)
    result["dict"].update(custom["dict"])
    return result


def _version(path):
    if not path.exists():
        return None
    stat = path.stat()
    return stat.st_mtime_ns, stat.st_size


def _current():
    return _merged(str(BASE_FILE), str(CUSTOM_FILE), _version(BASE_FILE), _version(CUSTOM_FILE))


def load_dictionary():
    with LOCK:
        return copy.deepcopy(_current())


def lookup(name):
    with LOCK:
        dictionary = _current()
        exact = dictionary["dict"].get(name)
        if exact:
            return {"cat": exact[0], "grp": exact[1], "water": bool(exact[2]), "via": "match"}
        for word, category, group in dictionary["kws"]:
            if word in name:
                return {"cat": category, "grp": group, "water": False, "via": "keyword"}
        return {"cat": None, "grp": None, "water": False, "via": None}


def _save(custom):
    CUSTOM_FILE.parent.mkdir(parents=True, exist_ok=True)
    temp = CUSTOM_FILE.with_name(CUSTOM_FILE.name + ".tmp")
    temp.write_text(json.dumps(custom, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, CUSTOM_FILE)
    _merged.cache_clear()


def add_category(category, group=None):
    category, group = _label(category, required=True), _label(group)
    with LOCK:
        current = _current()
        category = _canonical(category, current["categories"])
        group = _canonical(group, current["groupsByCat"].get(category, []))
        custom = _custom()
        created = category not in current["categories"] or bool(group and group not in current["groupsByCat"].get(category, []))
        if category not in current["categories"]:
            custom["categories"].append(category)
        groups = custom["groupsByCat"].setdefault(category, [])
        if group and group not in current["groupsByCat"].get(category, []):
            groups.append(group)
        _save(custom)
        return {"category": category, "group": group, "created": created, "dictionary": load_dictionary()}


def remember(name, category, group, water=False):
    category, group = _label(category, required=True), _label(group)
    with LOCK:
        current = _current()
        category = _canonical(category, current["categories"])
        if category not in current["categories"]:
            raise ValueError("Сначала добавьте категорию в справочник")
        group = _canonical(group, current["groupsByCat"].get(category, []))
        if group and group not in current["groupsByCat"].get(category, []):
            raise ValueError("Группа не принадлежит выбранной категории")
        custom = _custom()
        custom["dict"][name] = [category, group, int(bool(water))]
        _save(custom)
        return category, group
