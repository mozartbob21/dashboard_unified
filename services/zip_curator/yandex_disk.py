"""Bounded Yandex.Disk downloads. Exceptions are safe to show to a curator."""
import hashlib
import io
import time
import zipfile
from urllib.parse import urlsplit

import requests

API = "https://cloud-api.yandex.net/v1/disk/public/resources"
ATTEMPTS = 3
MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_EXPANDED_BYTES = 100 * 1024 * 1024
SYNC_SECONDS = 240
RETRY_STATUS = {408, 429, 500, 502, 503, 504}


class DiskSyncError(Exception):
    pass


class _Retryable(DiskSyncError):
    pass


def validate_workbook(data, item, read_rows):
    """Never replace a good cache with a truncated workbook or an HTTP error page."""
    if len(data) > MAX_FILE_BYTES:
        raise DiskSyncError("Файл превышает допустимый размер 50 МБ.")
    expected_size = item.get("size")
    if isinstance(expected_size, int) and len(data) != expected_size:
        raise _Retryable("Файл скачан не полностью: размер не совпадает с источником.")
    checksum = str(item.get("md5") or "").lower()
    if checksum and hashlib.md5(data).hexdigest() != checksum:
        raise _Retryable("Не удалось проверить целостность скачанного файла.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as workbook:
            members = workbook.infolist()
            if len(members) > 10000 or sum(m.file_size for m in members) > MAX_EXPANDED_BYTES:
                raise DiskSyncError("Слишком большой объём данных внутри Excel-файла.")
            if "[Content_Types].xml" not in workbook.namelist() or not any(
                m.filename.startswith("xl/worksheets/") for m in members
            ):
                raise ValueError("not a workbook")
            if workbook.testzip() is not None:
                raise ValueError("damaged archive")
        read_rows(data)
    except DiskSyncError:
        raise
    except Exception:
        raise _Retryable("Не удалось прочитать Excel-файл. Проверьте формат XLSX и повторите обновление.") from None


class DiskClient:
    def __init__(self, session, public_url, timeout=45):
        self.session = session
        self.public_url = public_url
        self.timeout = timeout
        self.deadline = time.monotonic() + SYNC_SECONDS

    def _timeout(self):
        left = self.deadline - time.monotonic()
        if left <= 0:
            raise DiskSyncError("Время обновления истекло. Повторите обновление оставшихся файлов.")
        return (min(10, left), min(self.timeout, left))

    def _pause(self, attempt):
        time.sleep(min(0.5 * (2 ** attempt), max(0, self.deadline - time.monotonic())))

    @staticmethod
    def _check_status(response, download=False):
        code = response.status_code
        if code == 200:
            return
        message = f"Яндекс Диск вернул HTTP {code}."
        if download and code == 500:
            # Recognise a known storage failure without exposing the response body
            # (or the signed URL) to the user.
            prefix = next(response.iter_content(chunk_size=2048), b"")
            if b"file IO failure: no successful lookup result" in prefix:
                raise _Retryable("Яндекс Диск не может прочитать файл в своём хранилище (HTTP 500). Если ошибка повторяется, загрузите исходный файл в папку А заново.")
        if code in RETRY_STATUS or (download and code in {401, 403, 404, 410}):
            raise _Retryable(message + " Повторите обновление позже.")
        if code in {401, 403, 404}:
            message += " Проверьте доступ по публичной ссылке."
        raise DiskSyncError(message)

    def _json(self, suffix="", **params):
        for attempt in range(ATTEMPTS):
            try:
                with self.session.get(
                    API + suffix, params={"public_key": self.public_url, **params},
                    timeout=self._timeout(),
                ) as response:
                    self._check_status(response)
                    try:
                        payload = response.json()
                    except ValueError:
                        raise _Retryable("Яндекс Диск вернул некорректный ответ.") from None
                    if not isinstance(payload, dict):
                        raise _Retryable("Яндекс Диск вернул некорректный ответ.")
                    return payload
            except (requests.RequestException, _Retryable):
                if attempt == ATTEMPTS - 1:
                    raise DiskSyncError("Не удалось получить ответ Яндекс Диска после трёх попыток. Повторите обновление позже.") from None
                self._pause(attempt)

    def list_files(self):
        # Validate the complete listing before changing or removing local files.
        offset, total, items, paths = 0, None, [], set()
        while True:
            payload = self._json(limit=200, offset=offset)
            embedded = payload.get("_embedded")
            if payload.get("type") != "dir" or not isinstance(embedded, dict):
                raise DiskSyncError("Яндекс Диск не вернул список папки. Проверьте ссылку на папку А.")
            batch, count = embedded.get("items"), embedded.get("total")
            if (not isinstance(batch, list) or type(count) is not int or count < 0
                    or count > 5000 or embedded.get("offset") != offset
                    or (total is not None and total != count)):
                raise DiskSyncError("Получен неполный список папки А. Локальные файлы сохранены; повторите обновление.")
            total = count
            if offset + len(batch) > total or (not batch and offset < total):
                raise DiskSyncError("Получен неполный список папки А. Локальные файлы сохранены; повторите обновление.")
            for item in batch:
                if (not isinstance(item, dict) or not item.get("path") or not item.get("name")
                        or item.get("type") not in {"file", "dir"} or item["path"] in paths):
                    raise DiskSyncError("Список папки А изменился или повреждён. Повторите обновление.")
                paths.add(item["path"])
                items.append(item)
            offset += len(batch)
            if offset == total:
                return items

    def download(self, item, read_rows):
        error = None
        for attempt in range(ATTEMPTS):
            try:
                # A signed link may expire or point to a failing download node.
                # Every retry obtains a fresh link instead of retrying the same URL.
                href = item.get("file") if attempt == 0 else None
                if not href:
                    href = self._json("/download", path=item["path"]).get("href")
                parsed = urlsplit(str(href or ""))
                host = parsed.hostname or ""
                if (parsed.scheme != "https" or parsed.username or parsed.password
                        or not any(host == domain or host.endswith("." + domain) for domain in
                                   ("disk.yandex.ru", "disk.yandex.net", "storage.yandex.net"))):
                    raise DiskSyncError("Яндекс Диск не вернул допустимую ссылку для скачивания.")
                with self.session.get(href, timeout=self._timeout(), stream=True) as response:
                    self._check_status(response, download=True)
                    data = bytearray()
                    for chunk in response.iter_content(chunk_size=65536):
                        self._timeout()
                        data.extend(chunk)
                        if len(data) > MAX_FILE_BYTES:
                            raise DiskSyncError("Файл превышает допустимый размер 50 МБ.")
                result = bytes(data)
                validate_workbook(result, item, read_rows)
                return result
            except requests.RequestException:
                error = "Соединение с Яндекс Диском прервалось или истекло время ожидания."
            except _Retryable as exc:
                error = str(exc)
            if attempt < ATTEMPTS - 1:
                self._pause(attempt)
        raise DiskSyncError(error + " Выполнено три попытки.")
