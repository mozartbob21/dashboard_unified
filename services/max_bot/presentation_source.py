"""Read the designated Redmine document using an isolated authenticated session.

No requests are made on import. Credentials, source responses and download URLs
must never be added to technical logs. The browser collector is not required.
"""
from __future__ import annotations

import os
import re
import tempfile
import unicodedata
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote, unquote, urljoin, urlsplit

import httpx
from bs4 import BeautifulSoup


ORIGIN = "https://digital.rm.mosreg.ru"
DOCUMENT_PATH = "/documents/16"
HTML_MAX_BYTES = 4 * 1024 * 1024
DEFAULT_MAX_BYTES = 50 * 1024 * 1024
_DOWNLOAD_PATH = re.compile(r"^/attachments/download/([1-9][0-9]{0,19})/[^/]+$")
_DATE = re.compile(r"\b(?:\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})(?:[ T]\d{2}:\d{2}(?::\d{2})?)?\b")


class SourceError(Exception):
    """Fixed, safe diagnostics for the administrator, never raw HTTP errors."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


@dataclass(frozen=True)
class Attachment:
    id: str
    name: str
    description: str = ""
    created_at: str = ""
    path: str = ""


def _safe_url(value: str, base: str = ORIGIN) -> str:
    """Resolve a location without ever crossing the fixed HTTPS origin."""
    if not isinstance(value, str) or not value or any(ord(c) < 32 for c in value) or "\\" in value:
        raise SourceError("SOURCE_UNSAFE_URL", "Источник вернул небезопасный адрес файла или перенаправления.")
    try:
        url = urljoin(base, value)
        parsed = urlsplit(url)
        if (parsed.scheme != "https" or parsed.hostname != urlsplit(ORIGIN).hostname
                or parsed.port not in (None, 443) or parsed.username or parsed.password
                or parsed.fragment):
            raise ValueError
    except ValueError:
        raise SourceError("SOURCE_UNSAFE_URL", "Источник вернул небезопасный адрес файла или перенаправления.") from None
    return url


def _download_path(value: str, attachment_id: str | None = None) -> tuple[str, str]:
    parsed = urlsplit(_safe_url(value))
    match = _DOWNLOAD_PATH.fullmatch(parsed.path)
    decoded = unquote(parsed.path)
    if (not match or parsed.query or "\\" in decoded or "/../" in decoded
            or "/./" in decoded or any(ord(c) < 32 for c in decoded)
            or len(decoded.split("/")) != 5
            or attachment_id is not None and match.group(1) != attachment_id):
        raise SourceError("SOURCE_UNSAFE_URL", "Источник вернул некорректный адрес вложения.")
    return parsed.path, match.group(1)


def _safe_filename(name: str) -> str:
    value = unicodedata.normalize("NFC", str(name))
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f\x7f]', "_", value).strip(" .")
    value = "".join(c for c in value if unicodedata.category(c) not in {"Cc", "Cf"})
    if not value:
        value = "presentation"
    # Keep common document extensions when shortening long Unicode names.
    suffix = Path(value).suffix[:15]
    stem = value[:-len(suffix)] if suffix else value
    # Filesystems limit bytes, not Unicode characters. Leave space for the ID.
    budget = 180 - len(suffix.encode("utf-8"))
    while len(stem.encode("utf-8")) > budget:
        stem = stem[:-1]
    value = stem + suffix
    if value.split(".", 1)[0].upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        value = "_" + value
    return value


def _is_login(soup: BeautifulSoup, url: str) -> bool:
    return urlsplit(url).path.rstrip("/") == "/login" or soup.select_one('input[name="password"]') is not None


def parse_attachments(page: str) -> list[Attachment]:
    """Parse metadata only; source-provided strings remain untrusted text."""
    soup = BeautifulSoup(page, "html.parser")
    if _is_login(soup, ORIGIN + DOCUMENT_PATH):
        raise SourceError("SOURCE_AUTH_FAILED", "Не удалось войти в источник. Проверьте сохранённые логин и пароль.")
    anchors = soup.select('a.icon-attachment[href], .attachments a[href*="/attachments/download/"]')
    if not anchors and soup.select_one(".attachments") is None:
        raise SourceError("SOURCE_FORMAT_CHANGED", "Страница источника не содержит ожидаемого списка вложений. Состояние проверки сохранено.")
    result: dict[str, Attachment] = {}
    for anchor in anchors:
        path, attachment_id = _download_path(anchor.get("href", ""))
        name = " ".join(anchor.get_text(" ", strip=True).split())
        if not name or len(name) > 1024:
            raise SourceError("SOURCE_FORMAT_CHANGED", "Источник вернул некорректное имя вложения.")
        container = anchor.find_parent("p") or anchor.find_parent("li") or anchor.parent
        fragment = BeautifulSoup(str(container), "html.parser")
        created_at = ""
        author = fragment.select_one(".author")
        if author:
            time_tag = author.find("time")
            date_candidate = time_tag.get("datetime", "") if time_tag else ""
            if not date_candidate:
                date_candidate = author.get_text(" ", strip=True)
            date_match = _DATE.search(date_candidate)
            created_at = date_match.group(0) if date_match else ""
        for element in fragment.select("a, span, time, script, style"):
            element.decompose()
        description = " ".join(fragment.get_text(" ", strip=True).split()).strip(" -–—")[:2000]
        item = Attachment(attachment_id, name, description, created_at, path)
        if attachment_id in result and result[attachment_id] != item:
            raise SourceError("SOURCE_FORMAT_CHANGED", "Источник вернул противоречивые сведения о вложении.")
        result[attachment_id] = item
    return list(result.values())


class RedmineSource:
    def __init__(self, username: str, password: str, *, transport=None, max_bytes: int = DEFAULT_MAX_BYTES):
        if not isinstance(username, str) or not username.strip() or not isinstance(password, str) or not password:
            raise SourceError("SOURCE_NOT_CONFIGURED", "В панели администратора не сохранены логин и пароль источника презентаций.")
        if type(max_bytes) is not int or max_bytes < 1:
            raise ValueError("max_bytes must be a positive integer")
        self._username = username
        self._password = password
        self.max_bytes = max_bytes
        self._client = httpx.Client(
            verify=True, trust_env=False, follow_redirects=False,
            timeout=httpx.Timeout(60, connect=20), transport=transport,
            headers={"User-Agent": "Neurona-Presentations/1", "Accept": "*/*"},
        )

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()

    def close(self):
        self._client.close()
        self._username = ""
        self._password = ""

    @contextmanager
    def _stream(self, method: str, path: str, *, data=None):
        url = _safe_url(path)
        try:
            for _ in range(6):
                request = self._client.build_request(method, url, data=data)
                response = self._client.send(request, stream=True)
                if response.status_code in {301, 302, 303, 307, 308}:
                    location = response.headers.get("location", "")
                    response.close()
                    url = _safe_url(location, url)
                    if response.status_code == 303 or response.status_code in {301, 302} and method == "POST":
                        method, data = "GET", None
                    continue
                try:
                    if response.status_code in {401, 403}:
                        raise SourceError("SOURCE_AUTH_FAILED", "Источник отклонил доступ. Проверьте учётную запись и её права.")
                    if response.status_code != 200:
                        raise SourceError("SOURCE_HTTP_FAILED", "Источник временно не вернул данные. Проверка будет повторена позже.")
                    yield response
                finally:
                    response.close()
                return
            raise SourceError("SOURCE_REDIRECT_FAILED", "Источник вернул слишком много перенаправлений.")
        except httpx.HTTPError:
            raise SourceError("SOURCE_CONNECTION_FAILED", "Не удалось безопасно подключиться к источнику презентаций. Проверьте сеть и сертификат источника.") from None

    @staticmethod
    def _bounded(response, limit: int):
        try:
            declared = int(response.headers.get("content-length", "0"))
        except ValueError:
            declared = 0
        if declared > limit:
            raise SourceError("SOURCE_TOO_LARGE", "Файл или ответ источника превышает разрешённый размер.")
        total = 0
        for chunk in response.iter_bytes(chunk_size=64 * 1024):
            total += len(chunk)
            if total > limit:
                raise SourceError("SOURCE_TOO_LARGE", "Файл или ответ источника превышает разрешённый размер.")
            yield chunk

    def _html(self, method: str, path: str, *, data=None) -> tuple[BeautifulSoup, str, str]:
        with self._stream(method, path, data=data) as response:
            content = b"".join(self._bounded(response, HTML_MAX_BYTES))
            page = content.decode("utf-8-sig", errors="replace")
            return BeautifulSoup(page, "html.parser"), str(response.url), page

    def _login(self):
        soup, _, _ = self._html("GET", "/login")
        token = soup.select_one('input[name="authenticity_token"]')
        if token is None or not token.get("value"):
            raise SourceError("SOURCE_LOGIN_CHANGED", "Форма входа источника изменилась. Требуется проверка администратором.")
        soup, url, _ = self._html("POST", "/login", data={
            "authenticity_token": token["value"], "username": self._username,
            "password": self._password, "back_url": ORIGIN + DOCUMENT_PATH, "login": "Вход",
        })
        if _is_login(soup, url):
            raise SourceError("SOURCE_AUTH_FAILED", "Не удалось войти в источник. Проверьте сохранённые логин и пароль.")

    def list_attachments(self) -> list[Attachment]:
        soup, url, page = self._html("GET", DOCUMENT_PATH)
        if _is_login(soup, url):
            self._login()
            soup, url, page = self._html("GET", DOCUMENT_PATH)
        if _is_login(soup, url):
            raise SourceError("SOURCE_AUTH_FAILED", "После входа источник снова запросил авторизацию.")
        if urlsplit(url).path.rstrip("/") != DOCUMENT_PATH:
            raise SourceError("SOURCE_FORMAT_CHANGED", "Источник перенаправил запрос на другую страницу. Состояние проверки сохранено.")
        return parse_attachments(page)

    def download(self, attachment: Attachment, directory: Path) -> Path:
        """Download is independent of discovery and may use a fresh session.

        A login-page redirect permits one form login. A genuine 401/403 remains
        a permission error rather than blindly retrying credentials.
        """
        try:
            return self._download_once(attachment, directory)
        except SourceError as exc:
            if exc.code != "SOURCE_SESSION_EXPIRED":
                raise
        self._login()
        try:
            return self._download_once(attachment, directory)
        except SourceError as exc:
            if exc.code == "SOURCE_SESSION_EXPIRED":
                raise SourceError("SOURCE_AUTH_FAILED", "После входа источник снова запросил авторизацию. Файл не отправлен.") from None
            raise

    def _download_once(self, attachment: Attachment, directory: Path) -> Path:
        attachment_id = str(attachment.id)
        if not re.fullmatch(r"[1-9][0-9]{0,19}", attachment_id):
            raise SourceError("SOURCE_INVALID_ATTACHMENT", "Некорректный идентификатор вложения.")
        path, _ = _download_path(attachment.path or f"/attachments/download/{attachment_id}/{quote(attachment.name, safe='')}", attachment_id)
        directory = Path(directory)
        if directory.is_symlink():
            raise SourceError("SOURCE_CACHE_FAILED", "Каталог временных файлов недоступен.")
        temporary: Path | None = None
        try:
            directory.mkdir(parents=True, exist_ok=True, mode=0o700)
            with self._stream("GET", path) as response:
                response_path = urlsplit(str(response.url)).path
                if response_path.rstrip("/") == "/login":
                    raise SourceError("SOURCE_SESSION_EXPIRED", "Источник запросил повторную авторизацию.")
                content_type = response.headers.get("content-type", "").split(";", 1)[0].strip().lower()
                if content_type in {"text/html", "application/xhtml+xml"}:
                    page = b"".join(self._bounded(response, HTML_MAX_BYTES)).decode("utf-8-sig", errors="replace")
                    if _is_login(BeautifulSoup(page, "html.parser"), str(response.url)):
                        raise SourceError("SOURCE_SESSION_EXPIRED", "Источник запросил повторную авторизацию.")
                    raise SourceError("SOURCE_INVALID_FILE", "Источник вернул страницу вместо файла. Вложение не отправлено.")
                fd, temporary_name = tempfile.mkstemp(prefix=".presentation-", suffix=".part", dir=directory)
                temporary = Path(temporary_name)
                total = 0
                with os.fdopen(fd, "wb") as output:
                    for chunk in self._bounded(response, self.max_bytes):
                        if not total:
                            prefix = chunk[:512].lstrip(b"\xef\xbb\xbf \r\n\t").lower()
                            if prefix.startswith((b"<!doctype html", b"<html", b"<form")):
                                raise SourceError("SOURCE_INVALID_FILE", "Источник вернул страницу вместо файла. Вложение не отправлено.")
                        output.write(chunk)
                        total += len(chunk)
                    if not total:
                        raise SourceError("SOURCE_INVALID_FILE", "Источник вернул пустой файл. Вложение не отправлено.")
                    output.flush()
                    os.fsync(output.fileno())
                target = directory / f"{attachment_id}-{_safe_filename(attachment.name)}"
                os.replace(temporary, target)
                temporary = None
                return target
        except OSError:
            raise SourceError("SOURCE_CACHE_FAILED", "Не удалось сохранить временный файл презентации. Проверьте свободное место и права каталога.") from None
        finally:
            if temporary is not None:
                try:
                    temporary.unlink(missing_ok=True)
                except OSError:
                    # A cache failure must not expose the original filesystem
                    # path or obscure the safe diagnostic from the operation.
                    pass
