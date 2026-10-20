"""Regression checks for removing prescriptions from the scheduler only."""
from contextlib import contextmanager
import sqlite3
import unittest
from unittest.mock import patch

from fastapi import FastAPI, Request
from fastapi.testclient import TestClient

from services import scheduler


class RetiredSchedulerJobTests(unittest.TestCase):
    def setUp(self):
        self.conn = sqlite3.connect(":memory:", check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("""
            CREATE TABLE scheduler_jobs (
                module_id TEXT PRIMARY KEY, module_label TEXT, module_icon TEXT,
                command TEXT, enabled INTEGER DEFAULT 0, interval_minutes INTEGER,
                last_run_at TEXT, next_run_at TEXT, last_status TEXT,
                last_error TEXT, created_at TEXT, updated_at TEXT
            )
        """)
        self.conn.execute("""
            INSERT INTO scheduler_jobs
                (module_id, enabled, interval_minutes, next_run_at, last_status)
            VALUES ('camera_prescriptions', 1, 360, '2000-01-01T00:00:00+00:00', 'success')
        """)

        @contextmanager
        def connection():
            with self.conn:
                yield self.conn

        self.db_patch = patch.object(scheduler, "get_db_connection", connection)
        self.db_patch.start()
        self.seed_patch = patch.object(scheduler, "_SEEDED", False)
        self.seed_patch.start()
        self.addCleanup(self.db_patch.stop)
        self.addCleanup(self.seed_patch.stop)
        self.addCleanup(self.conn.close)
        app = FastAPI()

        def user(request: Request):
            request.state.user = {"modules": ["cameras", "edo"]}
            return request.state.user

        app.dependency_overrides[scheduler.require_job_access] = user
        app.include_router(scheduler.router)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def test_old_enabled_job_is_disabled_and_hidden_without_losing_history(self):
        response = self.client.get("/api/scheduler/jobs")
        self.assertEqual(response.status_code, 200)
        self.assertNotIn("camera_prescriptions", [
            job["module_id"] for job in response.json()["jobs"]
        ])
        self.assertIn("cameras", [job["module_id"] for job in response.json()["jobs"]])
        row = self.conn.execute(
            "SELECT * FROM scheduler_jobs WHERE module_id='camera_prescriptions'"
        ).fetchone()
        self.assertEqual(row["enabled"], 0)
        self.assertIsNone(row["next_run_at"])
        self.assertEqual(row["last_status"], "success")
        self.assertIsNone(scheduler.get_job("camera_prescriptions"))
        scheduler._seed_jobs()  # Repeated reads do not restore the retired task.
        self.assertNotIn("camera_prescriptions", scheduler.MODULE_COMMANDS)
        self.assertNotIn("camera_prescriptions", scheduler.DEFAULT_INTERVALS)

    def test_retired_api_cannot_be_reenabled_or_launched(self):
        base = "/api/scheduler/jobs/camera_prescriptions"
        self.assertEqual(self.client.get(base).status_code, 404)
        for action in ("enable", "disable", "interval"):
            with self.subTest(action=action):
                response = self.client.post(
                    f"{base}/{action}", json={"interval_minutes": 60},
                )
                self.assertEqual(response.status_code, 400)
        with patch.object(scheduler, "_launch_module") as launch:
            self.assertEqual(self.client.post(base + "/run-now").status_code, 404)
            launch.assert_not_called()

    def test_tick_skips_retired_job_even_before_migration(self):
        with patch.object(scheduler, "_launch_module") as launch:
            scheduler._tick()
            launch.assert_not_called()
        with patch.object(scheduler, "_is_module_running") as running, \
                patch.object(scheduler, "mark_run_started") as started:
            scheduler._launch_module("camera_prescriptions")
            running.assert_not_called()
            started.assert_not_called()

    def test_other_jobs_still_enable_and_run(self):
        response = self.client.post("/api/scheduler/jobs/cameras/enable")
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.json()["job"]["enabled"])
        self.assertIsNotNone(response.json()["job"]["next_run_at"])
        with patch.object(scheduler, "_is_module_running", return_value=False), \
                patch.object(scheduler, "_launch_module") as launch:
            response = self.client.post("/api/scheduler/jobs/cameras/run-now")
            self.assertEqual(response.status_code, 200)
            launch.assert_called_once_with("cameras")


if __name__ == "__main__":
    unittest.main()
