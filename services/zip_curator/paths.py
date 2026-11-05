"""Storage locations used by both the ZIP dashboard and its read-only consumers."""
import os
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[2]
DATA = PROJECT / "data/zip_curator"
PUBLISHED_FILE = DATA / "published.json"
CONTACTS_FILE = DATA / "contacts.json"


def contacts_path(env=None, *, default=CONTACTS_FILE):
    """An optional shared directory on the host or a mounted volume, never an HTTP endpoint."""
    setting = (os.environ if env is None else env).get("ZIP_CONTACTS_DIR", "")
    setting = str(setting or "").strip()
    if not setting:
        return default
    if "://" in setting or "\x00" in setting:
        raise ValueError("ZIP_CONTACTS_DIR должен указывать на каталог файловой системы.")
    folder = Path(setting).expanduser()
    if not folder.is_absolute():
        folder = PROJECT / folder
    folder = folder.resolve()
    # No directory intended for public static resources or another block's downloads.
    if (folder.is_relative_to(PROJECT / "static") or folder.is_relative_to(PROJECT / "generated")
            or folder.is_relative_to(PROJECT / "data") and not folder.is_relative_to(DATA)):
        raise ValueError("Выберите закрытый каталог для ZIP_CONTACTS_DIR.")
    return folder / "contacts.json"
