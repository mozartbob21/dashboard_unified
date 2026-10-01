from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from services.water_dashboard.browser import WaterDashboardBrowserError, launch_context


@pytest.fixture(autouse=True)
def clean_environment(monkeypatch):
    monkeypatch.delenv("WATER_DASHBOARD_BROWSER_EXECUTABLE", raising=False)
    monkeypatch.delenv("EDDS_CHROME_EXECUTABLE", raising=False)


def fake_playwright(*results):
    launch = Mock(side_effect=results)
    return SimpleNamespace(chromium=SimpleNamespace(launch_persistent_context=launch)), launch


def missing():
    return RuntimeError("Executable doesn't exist at /private/user/secret/chromium")


def test_bundled_browser_keeps_tls_sandbox_and_own_profile(tmp_path):
    context = object()
    p, launch = fake_playwright(context)
    result, name = launch_context(p, profile_dir=tmp_path, headless=False)
    assert result is context and name == "Chromium"
    options = launch.call_args.kwargs
    assert options["user_data_dir"] == str(tmp_path / "chromium")
    assert options["ignore_https_errors"] is False
    assert options["chromium_sandbox"] is True
    assert options["accept_downloads"] is False
    assert options["headless"] is False
    assert "--no-sandbox" not in options.get("args", [])
    assert "executable_path" not in options and "channel" not in options


def test_missing_bundled_chromium_falls_back_to_installed_edge(tmp_path):
    context = object()
    p, launch = fake_playwright(missing(), context)
    assert launch_context(p, profile_dir=tmp_path) == (context, "Microsoft Edge")
    assert launch.call_count == 2
    assert launch.call_args.kwargs["channel"] == "msedge"
    assert launch.call_args.kwargs["user_data_dir"] == str(tmp_path / "edge")


def test_installed_chrome_is_used_if_edge_is_also_missing(tmp_path):
    context = object()
    p, launch = fake_playwright(missing(), missing(), context)
    assert launch_context(p, profile_dir=tmp_path) == (context, "Google Chrome")
    assert launch.call_args.kwargs["channel"] == "chrome"


def test_edds_executable_fallback_uses_a_separate_profile(tmp_path, monkeypatch):
    exe = tmp_path / "gost.exe"
    exe.touch()
    monkeypatch.setenv("EDDS_CHROME_EXECUTABLE", str(exe))
    context = object()
    p, launch = fake_playwright(missing(), missing(), missing(), context)
    assert launch_context(p, profile_dir=tmp_path / "water") == (context, "Chromium-GOST")
    assert launch.call_args.kwargs["executable_path"] == str(exe)
    assert launch.call_args.kwargs["user_data_dir"] == str(tmp_path / "water" / "edds-browser")


def test_explicit_browser_takes_precedence_and_retains_spaced_path(tmp_path, monkeypatch):
    exe = tmp_path / "Installed Browser" / "chrome.exe"
    exe.parent.mkdir()
    exe.touch()
    monkeypatch.setenv("WATER_DASHBOARD_BROWSER_EXECUTABLE", '"' + str(exe) + '"')
    context = object()
    p, launch = fake_playwright(context)
    assert launch_context(p, profile_dir=tmp_path / "water") == (context, "configured")
    assert launch.call_count == 1
    assert launch.call_args.kwargs["executable_path"] == str(exe)
    assert "channel" not in launch.call_args.kwargs
    assert launch.call_args.kwargs["user_data_dir"].startswith(str(tmp_path / "water" / "configured-"))


@pytest.mark.parametrize("directory", [False, True])
def test_invalid_explicit_browser_is_not_silently_replaced(tmp_path, monkeypatch, directory):
    exe = tmp_path if directory else tmp_path / "missing.exe"
    monkeypatch.setenv("WATER_DASHBOARD_BROWSER_EXECUTABLE", str(exe))
    p, launch = fake_playwright(object())
    with pytest.raises(WaterDashboardBrowserError, match="папка|не найден") as raised:
        launch_context(p, profile_dir=tmp_path / "water")
    assert raised.value.code == "browser_path"
    launch.assert_not_called()


def test_explicit_browser_startup_failure_does_not_fall_back_or_leak_logs(tmp_path, monkeypatch):
    exe = tmp_path / "chrome.exe"
    exe.touch()
    monkeypatch.setenv("WATER_DASHBOARD_BROWSER_EXECUTABLE", str(exe))
    p, launch = fake_playwright(RuntimeError("raw secret launch argument"))
    with pytest.raises(WaterDashboardBrowserError, match="WATER_DASHBOARD_BROWSER_EXECUTABLE") as raised:
        launch_context(p, profile_dir=tmp_path / "water")
    assert launch.call_count == 1
    assert "raw secret" not in str(raised.value)
    assert raised.value.code == "browser_launch"
    assert raised.value.__suppress_context__


@pytest.mark.parametrize("error", [
    RuntimeError("Failed to create a ProcessSingleton for your profile directory"),
    RuntimeError("user data directory is already in use"),
])
def test_profile_lock_is_actionable_and_not_bypassed_with_another_browser(tmp_path, error):
    p, launch = fake_playwright(error, object())
    with pytest.raises(WaterDashboardBrowserError, match="профиль.*уже открыт") as raised:
        launch_context(p, profile_dir=tmp_path)
    assert raised.value.code == "browser_locked"
    assert launch.call_count == 1


@pytest.mark.parametrize("error", [PermissionError("denied"), RuntimeError("spawn EACCES raw secret")])
def test_browser_permission_failure_stops_before_fallback(tmp_path, error):
    p, launch = fake_playwright(error, object())
    with pytest.raises(WaterDashboardBrowserError, match="запрещён запуск") as raised:
        launch_context(p, profile_dir=tmp_path)
    assert raised.value.code == "browser_permission"
    assert launch.call_count == 1


def test_unwritable_profile_is_actionable_without_browser_launch(tmp_path):
    profile_parent = tmp_path / "not-a-directory"
    profile_parent.touch()
    p, launch = fake_playwright(object())
    with pytest.raises(WaterDashboardBrowserError, match="Нет доступа к рабочему профилю"):
        launch_context(p, profile_dir=profile_parent)
    launch.assert_not_called()


def test_all_browsers_missing_explains_installation_and_hides_raw_paths(tmp_path, monkeypatch):
    monkeypatch.setenv("EDDS_CHROME_EXECUTABLE", str(tmp_path / "missing.exe"))
    p, launch = fake_playwright(missing(), missing(), missing())
    with pytest.raises(WaterDashboardBrowserError, match="python -m playwright install chromium") as raised:
        launch_context(p, profile_dir=tmp_path)
    assert launch.call_count == 3
    assert "secret" not in str(raised.value)
    assert raised.value.code == "browser_missing"


def test_generic_startup_failure_can_recover_with_installed_browser(tmp_path):
    context = object()
    p, launch = fake_playwright(RuntimeError("Browser closed unexpectedly"), context)
    assert launch_context(p, profile_dir=tmp_path) == (context, "Microsoft Edge")


def test_all_launches_failed_do_not_disclose_raw_errors(tmp_path):
    p, launch = fake_playwright(*(RuntimeError("raw secret log") for _ in range(3)))
    with pytest.raises(WaterDashboardBrowserError, match="Не удалось запустить браузер") as raised:
        launch_context(p, profile_dir=tmp_path)
    assert "raw secret" not in str(raised.value)
    assert launch.call_count == 3
