"""Server-only settings; no credentials are sent to the linking page."""
from dataclasses import dataclass, field
from pathlib import Path
import os
import re
import ipaddress
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parents[2]
BUNDLED_CA_FILE = ROOT / "deploy/certs/russian-trusted-root-ca.pem"
LAN_NETWORKS = tuple(ipaddress.ip_network(value) for value in (
    "10.0.0.0/8", "172.16.0.0/12", "192.168.0.0/16", "fc00::/7",
))


def lan_address(value):
    """RFC1918/ULA only; is_private also accepts unsuitable reserved ranges."""
    try:
        if not isinstance(value, str) or "%" in value:
            return False
        address = ipaddress.ip_address(value)
        return any(address in network for network in LAN_NETWORKS)
    except ValueError:
        return False


def http_origin(value):
    """Canonical internal HTTP origin, without credentials, paths or DNS names."""
    try:
        if not isinstance(value, str) or any(char.isspace() or ord(char) < 32 for char in value):
            raise ValueError
        parts = urlsplit(value)
        if (parts.scheme != "http" or parts.username is not None or parts.password is not None
                or parts.path or parts.query or parts.fragment or "?" in value or "#" in value
                or not lan_address(parts.hostname)):
            raise ValueError
        address = ipaddress.ip_address(parts.hostname)
        host = f"[{address.compressed}]" if address.version == 6 else str(address)
        port = parts.port
        if parts.netloc.endswith(":") or (port is not None and not 1 <= port <= 65535):
            raise ValueError
        return f"http://{host}" + (f":{port}" if port is not None and port != 80 else "")
    except (ValueError, TypeError):
        raise ValueError("MAX_ZIP_BOT_HTTP_ORIGINS: укажите HTTP-адреса внутренней сети с IP и портом, без пути.") from None


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
    http_origins: tuple[str, ...] = ()

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
        origins = str(env.get("MAX_ZIP_BOT_HTTP_ORIGINS") or "").strip()
        if len(origins) > 2048:
            raise ValueError("Слишком длинный список MAX_ZIP_BOT_HTTP_ORIGINS.")
        http_origins = tuple(dict.fromkeys(http_origin(item.strip()) for item in origins.split(","))) if origins else ()
        if len(http_origins) > 16:
            raise ValueError("В MAX_ZIP_BOT_HTTP_ORIGINS допускается не более 16 адресов.")
        return cls(enabled=enabled, mode=mode, bot_token=token, bot_public_name=name,
                   db_path=db, ca_file=ca, http_origins=http_origins)

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
