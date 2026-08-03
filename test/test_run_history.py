# tests/test_run_history.py
from services import run_history

def test_record_start_returns_run_id():
    run_id = run_history.record_start("edo", user="admin")
    assert isinstance(run_id, str)
    assert len(run_id) == 36  # UUID

def test_record_finish_calculates_duration():
    run_id = run_history.record_start("edo")
    # ... можно time.sleep(0.1) или мокнуть datetime
    run_history.record_finish(run_id, status="success")

    records = run_history.get_all()
    rec = next(r for r in records if r["run_id"] == run_id)
    
    assert rec["status"] == "success"
    assert rec["module_label"] == "Заполненность данных"
    assert rec["module_icon"] == "📋"
    assert rec["duration_seconds"] is not None

def test_get_recent_limits_results():
    for i in range(10):
        run_history.record_start("edo")
    
    recent = run_history.get_recent(limit=3)
    assert len(recent) == 3
    