"""Shared appearance preferences apply before render and stay consistent across tabs."""
from pathlib import Path
import json
import shutil
import subprocess
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]


def run_controller(settings=None, event=None, *, now='2026-10-01T12:00:00Z', denied=False):
    node = shutil.which('node')
    if not node:
        import playwright
        node = str(Path(playwright.__file__).parent / 'driver' / ('node.exe' if sys.platform == 'win32' else 'node'))
    script = """
const fs=require('fs'), vm=require('vm');
const options=JSON.parse(process.argv[1]);
const callbacks={}, root={dataset:{}};
class Clock extends Date { constructor(...args) { super(...(args.length?args:[options.now])); } }
const context={Date:Clock, Set, Math, document:{documentElement:root,readyState:'loading',addEventListener(){}},
 localStorage:{getItem(k){if(options.denied)throw Error('denied');return options.settings[k]},setItem(){}},
 window:{addEventListener(k,fn){callbacks[k]=fn}},setInterval(){}};
vm.runInNewContext(fs.readFileSync(process.argv[2],'utf8'),context);
if(options.event)callbacks.storage(options.event);
process.stdout.write(JSON.stringify(root.dataset));
"""
    options = {'settings':settings or {}, 'event':event, 'now':now, 'denied':denied}
    result = subprocess.run([node, '-e', script, json.dumps(options), str(ROOT/'static/appearance-menu.js')],
                            capture_output=True, text=True, check=True)
    return json.loads(result.stdout)


@pytest.mark.parametrize('mode', ['light', 'dark', 'kreml'])
def test_saved_theme_and_glass_are_applied_together(mode):
    assert run_controller({'neurona-theme-mode':mode,'neurona-glass':'on'}) == {'theme':mode,'glass':'on'}


@pytest.mark.parametrize(('instant','expected'), [('2026-10-01T12:00:00Z','light'),('2026-10-01T22:00:00Z','dark')])
def test_default_moscow_sun_uses_moscow_time(instant, expected):
    assert run_controller(now=instant)['theme'] == expected


def test_unavailable_local_storage_does_not_break_page():
    assert run_controller(denied=True) == {'theme':'light','glass':'off'}


def test_unknown_legacy_preference_falls_back_to_supported_theme():
    assert run_controller({'neurona-theme-mode':'invalid'})['theme'] == 'light'


def test_theme_change_in_another_tab_updates_current_page():
    assert run_controller({'neurona-theme-mode':'dark'}, {'key':'neurona-theme-mode','newValue':'kreml'})['theme'] == 'kreml'


def test_glass_change_in_another_tab_does_not_reset_theme():
    assert run_controller({'neurona-theme-mode':'dark'}, {'key':'neurona-glass','newValue':'on'}) == {'theme':'dark','glass':'on'}


def test_legacy_theme_controllers_are_not_copied_into_modules():
    for path in (ROOT/'templates').glob('*.html'):
        if path.name in {'auth_base.html','login.html','register.html'}:
            continue
        text=path.read_text(encoding='utf-8')
        assert 'localStorage.setItem("neurona-theme-mode"' not in text, path.name
        assert 'b.id = "glassToggle"' not in text, path.name
