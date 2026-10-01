"""Launch a browser for the water dashboard without relying on a downloaded build.

Installed browsers are alternatives when Playwright Chromium is unavailable. Each
variant uses a module-owned profile, never a personal or EDDS browser profile.
Only safe, actionable errors leave this module; Playwright logs may contain paths
and command-line values and must not be sent to the dashboard.
"""
import hashlib
import logging
import os
from pathlib import Path


class WaterDashboardBrowserError(RuntimeError):
    """A browser startup failure that is safe to display to the user."""

    def __init__(self, message, *, code="browser_launch"):
        super().__init__(message)
        self.code = code if code in {
            "browser_missing", "browser_path", "browser_locked",
            "browser_permission", "browser_launch",
        } else "browser_launch"


def _failure_kind(error):
    message = str(error).lower()
    if any(text in message for text in (
        "singleton", "profile in use", "user data directory is already in use",
        "profile appears to be in use", "profile is in use",
    )):
        return "locked"
    if isinstance(error, PermissionError) or any(text in message for text in (
        "permission denied", "access is denied", "eacces", "eperm",
    )):
        return "permission"
    if isinstance(error, FileNotFoundError) or any(text in message for text in (
        "executable doesn't exist", "executable does not exist", "executable not found",
        "distribution 'msedge' is not found", "distribution 'chrome' is not found",
        "enoent", "no such file or directory",
    )):
        return "missing"
    return "startup"


def _executable(env_name, *, required=False):
    value = os.getenv(env_name, "").strip().strip('"').strip("'")
    if not value:
        return None
    path = Path(os.path.expandvars(value)).expanduser()
    try:
        if path.is_file():
            return str(path)
        is_directory = path.is_dir()
    except OSError:
        if required:
            raise WaterDashboardBrowserError(
                f"Нет доступа к браузеру из {env_name}. Проверьте права учётной записи, "
                "под которой запущен сервер.", code="browser_permission",
            ) from None
        return None
    if required:
        description = "указана папка" if is_directory else "файл браузера не найден"
        raise WaterDashboardBrowserError(
            f"В {env_name} {description}. Укажите полный путь к существующему "
            "файлу chrome.exe или msedge.exe на компьютере-сервере.", code="browser_path",
        )
    return None


def launch_context(playwright, *, profile_dir, headless=True):
    """Return ``(context, browser_name)``; caller must close the returned context.

    An explicitly configured executable is authoritative. Otherwise try bundled
    Chromium, Edge, Chrome and finally the EDDS executable if it exists. Reusing
    only that executable allows the office's installed Chromium-GOST to collect
    public dashboards without sharing EDDS credentials, cookies or profile locks.
    """
    explicit = _executable("WATER_DASHBOARD_BROWSER_EXECUTABLE", required=True)
    if explicit:
        key = hashlib.sha256(os.path.normcase(explicit).encode()).hexdigest()[:12]
        candidates = [("configured", "configured-" + key, {"executable_path": explicit})]
    else:
        candidates = [
            ("Chromium", "chromium", {}),
            ("Microsoft Edge", "edge", {"channel": "msedge"}),
            ("Google Chrome", "chrome", {"channel": "chrome"}),
        ]
        edds_executable = _executable("EDDS_CHROME_EXECUTABLE")
        if edds_executable:
            candidates.append(("Chromium-GOST", "edds-browser", {"executable_path": edds_executable}))

    failures = []
    for name, profile_name, selection in candidates:
        profile = Path(profile_dir) / profile_name
        try:
            profile.mkdir(parents=True, exist_ok=True, mode=0o700)
        except OSError:
            raise WaterDashboardBrowserError(
                "Нет доступа к рабочему профилю сводного дашборда. Проверьте права "
                "на папку data/water_dashboard/playwright_profile у учётной записи сервера.",
                code="browser_permission",
            ) from None
        try:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(profile), headless=headless,
                viewport={"width": 1440, "height": 1100},
                extra_http_headers={"Cache-Control": "no-cache"},
                ignore_https_errors=False, chromium_sandbox=True,
                accept_downloads=False, timeout=20000, **selection,
            )
            return context, name
        except Exception as error:
            # Server console only; the worker exposes allowlisted diagnostics.
            logging.getLogger(__name__).warning("Браузер свода %s: %s", name, error)
            kind = _failure_kind(error)
            if kind == "locked":
                raise WaterDashboardBrowserError(
                    "Рабочий профиль сводного дашборда уже открыт. Дождитесь завершения "
                    "предыдущего сбора или закройте его окно браузера на сервере и повторите обновление.",
                    code="browser_locked",
                ) from None
            if kind == "permission":
                raise WaterDashboardBrowserError(
                    "Серверу запрещён запуск браузера сводного дашборда. Проверьте права "
                    "учётной записи сервера и разрешение запуска браузера в настройках защиты Windows.",
                    code="browser_permission",
                ) from None
            failures.append(kind)

    if explicit:
        raise WaterDashboardBrowserError(
            "Не удалось запустить браузер из WATER_DASHBOARD_BROWSER_EXECUTABLE. "
            "Проверьте, что он открывается под учётной записью сервера, и повторите обновление."
        ) from None
    if all(kind == "missing" for kind in failures):
        raise WaterDashboardBrowserError(
            "На сервере не найден браузер для сводного дашборда. Установите Chromium командой "
            "python -m playwright install chromium либо укажите путь к установленному Chrome, "
            "Edge или Chromium-GOST в WATER_DASHBOARD_BROWSER_EXECUTABLE в .env и перезапустите сервер.",
            code="browser_missing",
        ) from None
    raise WaterDashboardBrowserError(
        "Не удалось запустить браузер для сводного дашборда. Проверьте запуск Chrome или Edge "
        "под учётной записью сервера; нужный браузер можно задать в WATER_DASHBOARD_BROWSER_EXECUTABLE."
    ) from None
