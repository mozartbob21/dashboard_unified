"""History labels are localized at read time; tests never open the working DB."""
import ast
import importlib.util
import re
import shutil
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture
def history(tmp_path, monkeypatch):
    from utils import db

    monkeypatch.setattr(db, 'DB_FILE', tmp_path / 'history.sqlite')
    monkeypatch.setattr(db, '_SCHEMA_READY', False)
    # Import a copy so the one-time JSON migration cannot touch a live JSON file.
    source = tmp_path / 'services' / 'run_history.py'
    source.parent.mkdir()
    shutil.copyfile(ROOT / 'services' / 'run_history.py', source)
    spec = importlib.util.spec_from_file_location('isolated_history_labels', source)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _assignments(path, names):
    tree = ast.parse(path.read_text())
    return {
        node.targets[0].id: ast.literal_eval(node.value)
        for node in tree.body
        if isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name)
        and node.targets[0].id in names
    }


def test_all_platform_modules_and_scheduler_jobs_have_russian_labels(history):
    modules = _assignments(ROOT / 'core' / 'roles.py', {'ALL_MODULE_IDS'})['ALL_MODULE_IDS']
    scheduler = ast.parse((ROOT / 'services' / 'scheduler.py').read_text())
    commands = next(node.value for node in scheduler.body if isinstance(node, ast.AnnAssign)
                    and isinstance(node.target, ast.Name) and node.target.id == 'MODULE_COMMANDS')
    ids = set(modules) | {ast.literal_eval(key) for key in commands.keys}
    for module_id in ids:
        label, icon = history._module_presentation(module_id)
        assert label != history.DEFAULT_MODULE_LABEL, module_id
        assert not re.search('[A-Za-z_]', label), (module_id, label)
        assert icon != '⚙️', module_id


@pytest.mark.parametrize(('module_id', 'expected'), [
    ('mingkh-water-map', 'МИНЖКХ — карта жалоб'),
    ('water_dashboard', 'Качество водоснабжения'),
    ('water-dashboard', 'Качество водоснабжения'),
    ('mgkh_rm', 'Задачи МИНЖКХ'),
    ('mgkh-rm', 'Задачи МИНЖКХ'),
    ('camera_prescriptions', 'Предписания по камерам'),
    ('edds', 'ЕДДС — контроль заявок'),
])
def test_existing_saved_labels_are_refreshed_without_modifying_history(history, module_id, expected):
    run_id = history.record_start(module_id, user='test-user')
    with history.get_db_connection() as conn:
        conn.execute('UPDATE run_history SET module_label=?,module_icon=? WHERE run_id=?',
                     ('Old untranslated label', '📦', run_id))
    for records in (history.get_recent(), history.get_all()):
        assert records[0]['module_id'] == module_id
        assert records[0]['module_label'] == expected
        assert records[0]['module_icon'] == history._module_presentation(module_id)[1]
        assert records[0]['user'] == 'test-user'
        assert records[0]['status'] == 'running'
    with history.get_db_connection() as conn:
        saved = conn.execute('SELECT module_label,module_icon FROM run_history WHERE run_id=?', (run_id,)).fetchone()
        assert tuple(saved) == ('Old untranslated label', '📦')


def test_new_unknown_module_never_uses_raw_id_as_title(history):
    run_id = history.record_start('internal-unknown-job')
    record = history.get_recent()[0]
    assert record['run_id'] == run_id
    assert record['module_id'] == 'internal-unknown-job'
    assert record['module_label'] == 'Модуль платформы'
    with history.get_db_connection() as conn:
        assert conn.execute('SELECT module_label FROM run_history WHERE run_id=?', (run_id,)).fetchone()[0] == 'Модуль платформы'


def test_old_unknown_module_has_russian_fallback_even_with_saved_english_label(history):
    assert history._enrich({'module_id': 'removed-job', 'module_label': 'Old task'})['module_label'] == 'Модуль платформы'


def _local_access_functions():
    # Execute the real pure access functions without importing authentication or app startup.
    path = ROOT / 'core' / 'roles.py'
    tree = ast.parse(path.read_text())
    names = {'ALL_MODULE_IDS', 'MODULE_ALIASES'}
    function_names = {'allowed_history_modules', 'effective_modules', 'check_module_access'}
    selected = [node for node in tree.body if
                isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name) and node.targets[0].id in names
                or isinstance(node, ast.FunctionDef) and node.name in function_names]
    namespace = {}
    exec(compile(ast.Module(body=selected, type_ignores=[]), str(path), 'exec'), namespace)
    return namespace


def test_water_dashboard_alias_respects_existing_module_access(history):
    access = _local_access_functions()
    history.record_start('water_dashboard')
    history.record_start('edo')
    permitted = {'modules': ['water-dashboard'], 'role': 'Пользователь'}
    denied = {'modules': ['edo'], 'role': 'admin'}
    assert access['check_module_access'](permitted, 'water_dashboard') is True
    assert access['check_module_access'](denied, 'water_dashboard') is False
    permitted_history = history.get_recent(modules=access['allowed_history_modules'](permitted))
    denied_history = history.get_recent(modules=access['allowed_history_modules'](denied))
    assert [r['module_id'] for r in permitted_history] == ['water_dashboard']
    assert permitted_history[0]['module_label'] == 'Качество водоснабжения'
    assert [r['module_id'] for r in denied_history] == ['edo']
    assert history.get_recent(modules=access['allowed_history_modules'](None)) == []
