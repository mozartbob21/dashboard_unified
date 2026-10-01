"""Safe worker diagnostics: raw browser traces belong in the server console."""
import re
import sys

ERROR_PREFIX = "WATER_DASHBOARD_ERROR:"
MESSAGES = {
    "playwright_missing": "В среде Python сервера не установлен Playwright. Установите зависимости проекта в его .venv и перезапустите сервер.",
    "dependency": "В среде Python сервера не хватает зависимости сводного дашборда. Установите зависимости проекта в его .venv. Название пакета указано в журнале сервера.",
    "browser_missing": "Не найден браузер для сбора DataLens. Укажите путь к chrome.exe или msedge.exe в WATER_DASHBOARD_BROWSER_EXECUTABLE либо установите Chromium командой python -m playwright install chromium из .venv проекта.",
    "browser_path": "Не найден файл браузера WATER_DASHBOARD_BROWSER_EXECUTABLE. Укажите полный путь к chrome.exe или msedge.exe, а не к папке, и перезапустите сервер.",
    "browser_locked": "Профиль браузера сводного дашборда занят другим процессом. Завершите предыдущий сбор или закройте его окно браузера и повторите обновление.",
    "browser_permission": "Серверу запрещён запуск браузера или запись в его профиль. Проверьте права учётной записи, запускающей «Нейрону», и разрешение на запуск браузера.",
    "browser_launch": "Не удалось запустить браузер сводного дашборда. Проверьте WATER_DASHBOARD_BROWSER_EXECUTABLE и журнал сервера. Для запуска без рабочего стола используйте WATER_DASHBOARD_HEADLESS=1.",
    "certificate": "DataLens отклонён при проверке сертификата. Проверьте доверенные сертификаты и доступ к DataLens с компьютера-сервера.",
    "network": "Нет соединения с DataLens. Проверьте доступ к datalens.yandex с компьютера-сервера, DNS и настройки прокси.",
    "assets": "Не загрузились скрипты DataLens с yastatic.net. Проверьте доступ компьютера-сервера к этому домену; показатели не получены.",
    "timeout": "DataLens не загрузил данные за отведённое время. Проверьте доступ с компьютера-сервера и повторите обновление.",
    "access": "Источник DataLens требует входа или недоступен этой учётной записи.",
    "source_error": "Источник DataLens возвращает внутреннюю ошибку вместо показателей. Повторите обновление позже; прежние данные сохранены.",
    "captcha": "DataLens запрашивает подтверждение входа. Откройте браузер сборщика на сервере и пройдите проверку вручную.",
    "date": "Не удалось выбрать актуальную дату НВОС. Предыдущие данные этого источника сохранены.",
    "storage": "Не удалось сохранить данные свода. Проверьте свободное место и права на запись в data/water_dashboard у учётной записи сервера.",
    "sources_failed": "Ни один источник DataLens не обновлён. Причины указаны в карточках источников; ранее собранные данные сохранены.",
    "unknown": "Обновление сводного дашборда прервано. Причина записана в журнале сервера в строках [WATER_DASHBOARD] перед завершением процесса.",
}


class SourceReadError(RuntimeError):
    def __init__(self, code):
        self.code = code if code in MESSAGES else "unknown"
        super().__init__(MESSAGES[self.code])


def log_server_line(text):
    # A Windows cp1251 console cannot print every glyph in Playwright's logs.
    # Logging must not abort the worker before it receives the diagnostic code.
    encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
    print(str(text).encode(encoding, errors="replace").decode(encoding))


def error_code(exc):
    code = getattr(exc, "code", "")
    if code in MESSAGES:
        return code
    if isinstance(exc, ModuleNotFoundError):
        return "playwright_missing" if (exc.name or "").startswith("playwright") else "dependency"
    if isinstance(exc, TimeoutError):
        return "timeout"
    if isinstance(exc, (PermissionError, OSError)):
        return "storage"
    # Match known failure signatures; never expose exception text to clients.
    text = str(exc).lower()
    if "err_cert" in text or "certificate verify failed" in text:
        return "certificate"
    if any(s in text for s in ("err_name_not_resolved", "err_connection", "err_proxy", "err_tunnel", "err_internet_disconnected")):
        return "network"
    if "timeout" in text or "timed out" in text:
        return "timeout"
    return "unknown"


def failure_message(output):
    """Accept only allowlisted codes, not arbitrary subprocess output."""
    matches = re.findall(r"^WATER_DASHBOARD_ERROR:([a-z_]+)$", output, re.M)
    return MESSAGES.get(matches[-1], MESSAGES["unknown"]) if matches else MESSAGES["unknown"]
