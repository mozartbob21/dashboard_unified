# test/test_api.py
"""Тесты API планировщика (admin-only endpoints)."""


def test_scheduler_jobs_requires_admin(api_client):
    """Без авторизации → редирект на /login или 401."""
    resp = api_client.get("/api/scheduler/jobs", follow_redirects=False)
    assert resp.status_code in (302, 401)


def test_scheduler_jobs_list(admin_client):
    """Авторизованный админ получает список задач."""
    resp = admin_client.get("/api/scheduler/jobs")
    assert resp.status_code == 200
    
    data = resp.json()
    assert data["ok"] is True
    assert isinstance(data["jobs"], list)


def test_scheduler_enable_job(admin_client):
    """Включение автозапуска модуля."""
    resp = admin_client.post("/api/scheduler/jobs/edo/enable")
    assert resp.status_code == 200
    
    data = resp.json()
    assert data["ok"] is True
    assert data["job"]["enabled"] is True


def test_scheduler_disable_job(admin_client):
    """Выключение автозапуска модуля."""
    # Сначала включим
    admin_client.post("/api/scheduler/jobs/edo/enable")
    
    # Теперь выключим
    resp = admin_client.post("/api/scheduler/jobs/edo/disable")
    assert resp.status_code == 200
    
    data = resp.json()
    assert data["ok"] is True
    assert data["job"]["enabled"] is False


def test_scheduler_set_interval(admin_client):
    """Изменение интервала запуска."""
    resp = admin_client.post(
        "/api/scheduler/jobs/edo/interval",
        json={"interval_minutes": 90},
    )
    assert resp.status_code == 200
    
    data = resp.json()
    assert data["ok"] is True
    assert data["job"]["interval_minutes"] == 90
    