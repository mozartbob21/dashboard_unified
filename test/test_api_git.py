# test/test_api_git.py
"""Тесты API Git (обновления проекта)."""
from unittest.mock import patch


class TestGitCheck:
    def test_requires_token(self, admin_client):
        """Без токена → 403."""
        resp = admin_client.get("/system/git/check")
        assert resp.status_code == 403

    def test_wrong_token(self, admin_client):
        """Неверный токен → 403."""
        resp = admin_client.get(
            "/system/git/check",
            headers={"X-Git-Update-Token": "wrong"},
        )
        assert resp.status_code == 403

    @patch("app.is_git_repository")
    def test_not_a_repo(self, mock_is_repo, admin_client):
        """Папка не является Git-репозиторием."""
        mock_is_repo.return_value = False

        resp = admin_client.get(
            "/system/git/check",
            headers={"X-Git-Update-Token": "12345"},
        )

        assert resp.status_code == 400
        assert resp.json()["ok"] is False

    @patch("app.run_git_command")
    @patch("app.is_git_repository")
    @patch("app.has_local_git_changes")
    @patch("app.get_current_git_branch")
    def test_no_updates(
        self, mock_branch, mock_changes, mock_is_repo, mock_git, admin_client
    ):
        """Всё ок, обновлений нет."""
        mock_is_repo.return_value = True
        mock_branch.return_value = "main"
        mock_changes.return_value = {"ok": True, "has_changes": False}
        mock_git.return_value = {"ok": True, "stdout": "", "stderr": ""}

        resp = admin_client.get(
            "/system/git/check",
            headers={"X-Git-Update-Token": "12345"},
        )

        assert resp.status_code == 200
        data = resp.json()
        assert data["ok"] is True
        assert data["has_updates"] is False


class TestGitPull:
    @patch("app.run_git_command")
    @patch("app.is_git_repository")
    @patch("app.has_local_git_changes")
    @patch("app.get_current_git_branch")
    def test_pull_with_local_changes_blocked(
        self, mock_branch, mock_changes, mock_is_repo, mock_git, admin_client
    ):
        """Есть незакоммиченные правки → 409 Conflict."""
        mock_is_repo.return_value = True
        mock_branch.return_value = "main"
        mock_changes.return_value = {"ok": True, "has_changes": True}

        resp = admin_client.post(
            "/system/git/pull",
            headers={"X-Git-Update-Token": "12345"},
        )

        assert resp.status_code == 409
        assert "локальные изменения" in resp.json()["message"].lower()

    @patch("app.run_git_command")
    @patch("app.is_git_repository")
    @patch("app.has_local_git_changes")
    @patch("app.get_current_git_branch")
    def test_pull_success(
        self, mock_branch, mock_changes, mock_is_repo, mock_git, admin_client
    ):
        """Успешный pull."""
        mock_is_repo.return_value = True
        mock_branch.return_value = "main"
        mock_changes.return_value = {"ok": True, "has_changes": False}

        def git_side_effect(cmd, timeout=30):
            if "fetch" in cmd:
                return {"ok": True, "stdout": "", "stderr": ""}
            if "pull" in cmd:
                return {"ok": True, "stdout": "Already up to date", "stderr": ""}
            return {"ok": True, "stdout": "", "stderr": ""}

        mock_git.side_effect = git_side_effect

        resp = admin_client.post(
            "/system/git/pull",
            headers={"X-Git-Update-Token": "12345"},
        )

        assert resp.status_code == 200
        assert resp.json()["ok"] is True
        