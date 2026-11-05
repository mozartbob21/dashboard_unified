"""Exercise the actual app executor with bounded synthetic local subprocesses.

AST loading avoids app startup, source connections and runtime account stores.
Only the executor definitions are loaded; the same command/progress/history path
used by normal buttons and chat is exercised. No network or real source data.
"""
import ast
from datetime import datetime
import os
from pathlib import Path
import signal
import subprocess
import sys
import threading
import time
from types import SimpleNamespace
from unittest.mock import Mock

import pytest


@pytest.fixture
def executor(tmp_path):
    source = Path(__file__).resolve().parents[1] / 'app.py'
    tree = ast.parse(source.read_text(encoding='utf-8-sig'))
    wanted = {'_run_subprocess_worker', 'run_subprocess_worker'}
    definitions = [node for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in wanted]
    assert {node.name for node in definitions} == wanted
    processes = []
    def launch(*args, **kwargs):
        process = subprocess.Popen(*args, **kwargs)
        processes.append(process)
        return process
    transport = SimpleNamespace(Popen=launch, PIPE=subprocess.PIPE, STDOUT=subprocess.STDOUT,
        CREATE_NEW_PROCESS_GROUP=getattr(subprocess, 'CREATE_NEW_PROCESS_GROUP', 0),
        run=subprocess.run, SubprocessError=subprocess.SubprocessError, TimeoutExpired=subprocess.TimeoutExpired)
    history = SimpleNamespace(record_start=Mock(return_value='synthetic-new-run'), record_finish=Mock())
    namespace = {'__builtins__': __builtins__, 'Path': Path, 'BASE_DIR': tmp_path,
        'run_status': {'synthetic': {}}, 'save_run_time': Mock(), 'run_history': history,
        'subprocess': transport, 'threading': threading, 'os': os, 'datetime': datetime}
    exec(compile(ast.Module(body=definitions, type_ignores=[]), str(source), 'exec'), namespace)
    def cleanup():
        for process in processes:
            try:
                if os.name == 'nt':
                    if process.poll() is None:
                        subprocess.run(['taskkill', '/PID', str(process.pid), '/T', '/F'], capture_output=True, timeout=5)
                else:
                    # The group may still contain children after its leader exits.
                    os.killpg(process.pid, signal.SIGKILL)
            except (ProcessLookupError, subprocess.SubprocessError):
                pass
            if process.poll() is None:
                process.kill()
            process.wait(timeout=5)
            if process.stdout:
                process.stdout.close()
    def bounded(code, *, timeout=1, callback=None, run_id='synthetic-run-id-12345'):
        box = {}
        def work():
            try:
                box['result'] = namespace['run_subprocess_worker']('synthetic', [sys.executable, '-u', '-c', code],
                    tmp_path, run_id=run_id, timeout_seconds=timeout, progress_callback=callback)
            except BaseException as error:
                box['error'] = error
        thread = threading.Thread(target=work, daemon=True)
        thread.start()
        thread.join(5)
        if thread.is_alive():
            cleanup()
            thread.join(2)
            pytest.fail('Executor did not finish within the bounded local test')
        if 'error' in box:
            raise box['error']
        return box['result']
    harness = SimpleNamespace(namespace=namespace, history=history, processes=processes, run=bounded, cleanup=cleanup)
    yield harness
    cleanup()


def test_real_executor_keeps_history_id_reports_progress_and_confirms_exit(executor):
    progress = []
    result = executor.run("print('STAGE: Synthetic collection'); print('result saved')", callback=progress.append)
    assert result == {'status': 'success'}
    assert progress == ['Исполнитель блока запущен', 'Блок обрабатывает данные']
    executor.history.record_start.assert_not_called()
    executor.history.record_finish.assert_called_once_with('synthetic-run-id-12345', 'success')
    status = executor.namespace['run_status']['synthetic']
    assert status['running'] is False and status['stage'] == 'Готово' and status['finished_at']
    assert executor.processes[0].poll() == 0


def test_real_nonzero_exit_is_failure_and_stops_tracking(executor):
    result = executor.run("print('STAGE: Synthetic check'); raise SystemExit(7)")
    assert result['status'] == 'error' and '7' in result['error']
    assert executor.history.record_finish.call_args.args[1] == 'error'
    assert executor.namespace['run_status']['synthetic']['running'] is False
    assert executor.processes[0].poll() == 7


def test_silent_hung_process_is_killed_on_execution_deadline(executor):
    started = time.monotonic()
    result = executor.run('import time; time.sleep(30)', timeout=1)
    assert time.monotonic() - started < 5
    assert result['status'] == 'error' and 'предел времени' in result['error']
    assert executor.processes[0].poll() is not None
    assert executor.namespace['run_status']['synthetic']['running'] is False


@pytest.mark.skipif(os.name == 'nt', reason='POSIX process-group cleanup; Windows taskkill needs a Windows runner')
def test_exited_parent_does_not_leave_child_holding_stdout_until_forever(executor):
    # The parent exits successfully while its child inherits the output pipe.
    # A timeout must still terminate the group and mark the run as interrupted.
    code = "import subprocess,sys; subprocess.Popen([sys.executable,'-c','import time; time.sleep(30)']); print('parent exited')"
    result = executor.run(code, timeout=1)
    assert result['status'] == 'error'
    assert 'предел времени' in result['error']
    assert executor.namespace['run_status']['synthetic']['running'] is False


def test_progress_failure_cleans_up_the_running_executor(executor):
    def fail(_stage):
        raise RuntimeError('synthetic progress write failure')
    result = executor.run('import time; time.sleep(30)', timeout=20, callback=fail)
    assert result['status'] == 'error'
    assert executor.processes[0].poll() is not None
    assert executor.namespace['run_status']['synthetic']['running'] is False


def test_shared_process_lock_refuses_a_second_executor(executor, tmp_path):
    from services.mingkh.sms.context import process_lock
    path = tmp_path / 'data' / 'execution-locks' / 'synthetic.lock'
    with process_lock(path):
        result = executor.run("raise AssertionError('must not start')")
    assert result['status'] == 'busy'
    assert not executor.processes
    executor.history.record_finish.assert_called_once_with('synthetic-run-id-12345', 'skipped', 'Блок уже выполняется.')
    # Once released, the same regular executor becomes usable again.
    assert executor.run("print('ok')")['status'] == 'success'


def test_rejected_second_launch_keeps_first_collectors_visible_status(executor, tmp_path):
    from services.mingkh.sms.context import process_lock
    from services.process_control import set_active
    set_active('synthetic', True)
    try:
        with process_lock(tmp_path / 'data/execution-locks/synthetic.lock'):
            assert executor.run('raise AssertionError()')['status'] == 'busy'
            assert executor.namespace['run_status']['synthetic']['running'] is True
            assert not executor.processes
    finally:
        set_active('synthetic', False)
