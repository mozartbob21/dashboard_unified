"""Server-only settings; no credentials are sent to the linking page."""
from dataclasses import dataclass, field
from pathlib import Path
import os
import re

ROOT = Path(__file__).resolve().parents[2]
BUNDLED_CA_FILE = ROOT / "deploy/certs/russian-trusted-root-ca.pem"


@dataclass(frozen=True)
class Settings:
    enabled: bool = False
    mode: str = "disabled"
    bot_token: str = field(default="", repr=False)
    bot_public_name: str = ""
    accounts_path: Path = ROOT / "data/dashboard.db"
    db_path: Path = ROOT / "data/max_bot.sqlite"
    published_path: Path = ROOT / "data/zip_curator/published.json"
    contacts_path: Path = ROOT / "data/zip_curator/contacts.json"
    binding_ttl: int = 30 * 86400
    link_ttl: int = 300
    audit_days: int = 30
    ca_file: Path | None = BUNDLED_CA_FILE

    @classmethod
    def from_env(cls, mapping=None):
        env = os.environ if mapping is None else mapping
        enabled = str(env.get("MAX_ZIP_BOT_ENABLED", "false")).strip().lower() in {"1", "true", "yes"}
        mode = str(env.get("MAX_ZIP_BOT_MODE", "disabled")).strip()
        if enabled and mode != "development-polling":
            raise ValueError("Для локальной проверки задайте MAX_ZIP_BOT_MODE=development-polling.")
        name = str(env.get("MAX_ZIP_BOT_PUBLIC_NAME", "")).strip()
        if name and not re.fullmatch(r"@?[A-Za-z0-9_]{1,64}", name):
            raise ValueError("Некорректное публичное имя бота MAX.")
        token = str(env.get("MAX_ZIP_BOT_TOKEN", "") or "").strip()
        if any(ord(c) < 33 or ord(c) > 126 for c in token) or len(token) > 4096:
            raise ValueError("Некорректный формат MAX_ZIP_BOT_TOKEN.")
        db = Path(env.get("MAX_ZIP_BOT_DB") or ROOT / "data/max_bot.sqlite").expanduser()
        if not db.is_absolute():
            db = ROOT / db
        ca_value = str(env.get("MAX_ZIP_BOT_CA_FILE") or "").strip()
        ca = Path(ca_value).expanduser() if ca_value else BUNDLED_CA_FILE
        if not ca.is_absolute():
            ca = ROOT / ca
        return cls(enabled=enabled, mode=mode, bot_token=token, bot_public_name=name,
                   db_path=db, ca_file=ca)

    def validate_polling(self):
        if not self.enabled or self.mode != "development-polling":
            raise ValueError("Локальная проверка MAX отключена в настройках.")
        if not self.bot_token:
            raise ValueError("Добавьте MAX_ZIP_BOT_TOKEN в серверный .env.")


def enabled_for_ui():
    # A typo in optional bot settings must not break the existing ZIP dashboard.
    try:
        return Settings.from_env().enabled
    except ValueError:
        return False


def available_for_user(user):
    """Mirror current local-account eligibility before rendering entry points."""
    return bool(user and user.get("is_active", True) and not user.get("kc_sub")
                and type(user.get("id")) is int and "zips" in user.get("modules", [])
                and enabled_for_ui())
