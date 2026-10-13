"""Persistent curator additions layered over the bundled classification dictionary."""
import copy
import json
import os
import re
import tempfile
import threading
from functools import lru_cache
from pathlib import Path

BASE_FILE = Path(__file__).with_name("zip_dict.json")
CUSTOM_FILE = Path(__file__).resolve().parents[2] / "data" / "zip_curator" / "dictionary.json"
LOCK = threading.RLock()
MAX_IMPORT_BYTES = 10 * 1024 * 1024


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
    temp = None
    try:
        with tempfile.NamedTemporaryFile(dir=CUSTOM_FILE.parent, prefix=".dictionary-", suffix=".tmp",
                                         mode="w", encoding="utf-8", delete=False) as stream:
            temp = Path(stream.name)
            json.dump(custom, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp, CUSTOM_FILE)
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)
    _merged.cache_clear()


def _water_flag(value):
    text = str(value if value is not None else "").strip().casefold()
    if text in {"да", "1", "true", "yes", "+"}:
        return 1
    if text in {"", "нет", "0", "false", "no", "-"}:
        return 0
    raise ValueError('В колонке «Вода» укажите «да», «нет», 1, 0 или оставьте ячейку пустой')


def import_workbook(filename, data):
    """Validate the complete upload, then persist one merged dictionary atomically."""
    # Local import avoids the core -> dictionary dependency during initialization.
    from services.zip_curator.core import norm, read_xlsx_rows
    from services.zip_curator.yandex_disk import validate_workbook

    if Path(str(filename or "")).suffix.lower() != ".xlsx":
        raise ValueError("Выберите справочник в формате XLSX")
    if not data or len(data) > MAX_IMPORT_BYTES:
        raise ValueError("Файл справочника должен быть непустым и не больше 10 МБ")
    try:
        validate_workbook(data, {}, lambda _: None)
        rows = read_xlsx_rows(data)
    except Exception:
        raise ValueError("Не удалось прочитать справочник. Проверьте файл XLSX или скачайте образец кнопкой «Скачать справочник»") from None
    if not rows:
        raise ValueError("Справочник пуст")
    if len(rows) > 50001:
        raise ValueError("В справочнике допускается не больше 50 000 строк")
    columns = {}
    for index, value in enumerate(rows[0]):
        header = str(value or "").strip().casefold()
        key = "name" if header in {"наименование", "наименование (норм.)", "наименование (норм)"} else {
            "категория": "category", "группа": "group", "вода": "water"
        }.get(header)
        if key:
            if key in columns:
                raise ValueError("В справочнике повторяются заголовки колонок")
            columns[key] = index
    if not {"name", "category", "group"}.issubset(columns):
        raise ValueError("Нужны колонки «Наименование (норм.)», «Категория» и «Группа». Скачайте образец кнопкой «Скачать справочник»")

    def cell(row, key):
        index = columns.get(key)
        return row[index] if index is not None and index < len(row) else None

    with LOCK:
        # Build changes in independent objects; a bad last row cannot partially save the file.
        current = copy.deepcopy(_current())
        custom = _custom()
        imported = {}
        for row_number, row in enumerate(rows[1:], start=2):
            if all(value is None or str(value).strip() == "" for value in row):
                continue
            try:
                name = norm(cell(row, "name"))
                if not name or len(name) > 1000:
                    raise ValueError("Наименование должно содержать от 1 до 1000 символов")
                category, group = _label(cell(row, "category")), _label(cell(row, "group"))
                if group and not category:
                    raise ValueError("Для группы необходимо указать категорию")
                category = _canonical(category, current["categories"])
                group = _canonical(group, current["groupsByCat"].get(category, []))
                entry = [category or None, group or None, _water_flag(cell(row, "water"))]
                if name in imported and imported[name] != entry:
                    raise ValueError("Для одного наименования указаны разные категории, группы или признаки воды")
            except ValueError as error:
                raise ValueError(f"Строка {row_number}: {error}") from None
            imported[name] = entry
            if category:
                if category not in current["categories"]:
                    current["categories"].append(category)
                    custom["categories"].append(category)
                if group and group not in current["groupsByCat"].get(category, []):
                    current["groupsByCat"].setdefault(category, []).append(group)
                    custom["groupsByCat"].setdefault(category, []).append(group)
        if not imported:
            raise ValueError("В справочнике нет записей для загрузки")
        custom["dict"].update(imported)
        _save(custom)
        return {"imported": len(imported), "dictionary": load_dictionary()}


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
