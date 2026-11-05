"""Bounded numerical projections of known collector outputs. Never persist source rows."""
from __future__ import annotations

import json
import math
import re
from datetime import date, datetime
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
MAX_BYTES = 32 * 1024 * 1024
VERSION = 1

# Key, human label and whether a decrease has an established interpretation.
# Counts of observed activity have no intrinsic good/bad direction.
DEFINITIONS = {
    'items': ('Записей', ''), 'rows': ('Строк в источнике', ''), 'files': ('Файлов', ''),
    'positions': ('Позиций', ''), 'organizations': ('Организаций', ''),
    'municipalities': ('Муниципалитетов', ''), 'appeals': ('Обращений', ''),
    'signatures': ('Подписей', ''), 'tasks': ('Задач', ''), 'sources': ('Источников', ''),
    'failed_sources': ('Источников с ошибкой', 'lower'), 'issues': ('Объектов с незаполненными данными', 'lower'),
    'overdue': ('Просроченных задач', 'lower'), 'violators': ('Объектов с нарушениями', 'lower'),
    'working': ('Работающих камер', 'higher'), 'not_working': ('Неработающих камер', 'lower'),
    'not_connected': ('Неподключённых камер', 'lower'), 'unknown': ('Камер с неизвестным статусом', ''),
    'close': ('К закрытию', ''), 'extend': ('К продлению', ''), 'rework': ('К доработке', ''),
    'problems': ('Проблемных камер', 'lower'), 'documents': ('Индивидуальных документов', ''),
    'water_issues': ('Строк контроля воды с замечаниями', 'lower'), 'incidents': ('Инцидентов', ''), 'active': ('Открытых инцидентов', ''), 'closed': ('Закрытых инцидентов', ''),
    'added_registries': ('Добавленных реестров', ''), 'unclassified_tasks': ('Неклассифицированных задач', ''),
    'facilities': ('Объектов', ''), 'days': ('Дней в сохранённом своде', ''), 'new_files': ('Новых файлов', ''),
}
PATHS = {
    'edo': 'data/edo/result.json', 'overdue': 'data/overdue/final_result.json',
    'watercontrol': 'data/watercontrol/result.json', 'utnkr': 'data/utnkr/violators.json',
    'cds': 'data/cds/result.json', 'cameras': 'data/cameras/state/dashboard_state.json',
    'camera_prescriptions': 'generated/prescriptions/generation_result.json',
    'water_dashboard': 'data/water_dashboard/snapshot.json', 'water-dashboard': 'data/water_dashboard/snapshot.json',
    'mgkh_rm': 'data/mgkh_rm/result.json', 'edds': 'data/edds/water_daily.json',
}
OPERATIONS = {'all','scan','upload','refresh','collect','arm','arm_coordinates','daily','individual','combined','zip','archive','published','range','full','source','ranking','sync','rating','collective','zips_scan','water_rm','water_map','treatment','presentations','sms','local_rows','live','process'}
SOURCE_IDS = {'valves','tasks','sys_vs','edo_rso','nvos','controller_tasks','energy_audit','gkh_kontur','meetings','sys_kr','all','yandex','mosreg','both'}

def iso_date(value):
    if not isinstance(value, str) or len(value) != 10:
        return None
    try:
        return (datetime.strptime(value, '%d.%m.%Y').date() if '.' in value else date.fromisoformat(value)).isoformat()
    except ValueError:
        return None


def safe_params(params):
    if not params:
        return {}
    if not isinstance(params, dict):
        return {'unknown': True}
    result = {}
    # An unsupported parameter makes comparisons fail closed without keeping its value.
    for key, value in params.items():
        if key in ('start', 'end') and iso_date(value):
            result[key] = iso_date(value)
        elif key == 'config_hash' and isinstance(value, str) and re.fullmatch(r'[0-9a-f]{64}',value):
            result[key] = value
        elif key == 'operation' and isinstance(value, str) and value in OPERATIONS:
            result[key] = value
        elif key == 'source' and isinstance(value, str) and value in SOURCE_IDS:
            result[key] = value
        else:
            result['unknown'] = True
    return result


def metric_info(key):
    if key in DEFINITIONS:
        label, polarity = DEFINITIONS[key]
        return label, ('дн.' if key == 'days' else 'шт.'), polarity
    from services.check_water_snapshot import DEFINITIONS as water_metrics
    if key in water_metrics:
        label, unit = water_metrics[key]
        return label, unit, ''
    return None


def safe_metrics(metrics):
    if not isinstance(metrics, dict):
        return {}
    return {key: value for key, value in metrics.items()
            if metric_info(key) and type(value) in (int, float)
            and 0 <= value <= 10**15 and math.isfinite(value)}


def command_params(module, command):
    # Persist only whitelisted inputs, never a shell command, URL or environment.
    if module == 'cds' and len(command) >= 5:
        return {'start':command[-2], 'end':command[-1]}
    if module in ('water_dashboard','water-dashboard'):
        if '--source' in command:
            index = command.index('--source')
            return {'source':command[index+1] if index+1<len(command) else None}
        return {'source':'all'}
    return {}


def signature(module):
    relative = PATHS.get(module)
    if not relative:
        return None
    try:
        stat = (BASE / relative).stat()
        return [stat.st_mtime_ns, stat.st_size, stat.st_ino]
    except OSError:
        return None


def _project(module, value, params):
    if value.get('ok') is False or value.get('success') is False or value.get('status') in ('error','failed'):
        return 'failed', {}, {}
    if 'Проверка не завершена' in str(value.get('summary_message','')):
        return 'failed', {}, {}
    if module == 'overdue' and value.get('extraction_note'):
        return ('partial' if value.get('items') else 'failed'), {}, {}
    if module in ('edo', 'watercontrol') and value.get('rows') == [] and value.get('extraction_note'):
        return 'unconfirmed', {}, {}
    if module in ('edo','watercontrol'):
        metrics = {'rows': len(value['rows'])} if isinstance(value.get('rows'), list) else {}
        if module == 'watercontrol' and isinstance(value.get('rows'), list) and all(
                isinstance(row,dict) and isinstance(row.get('status'),str) for row in value['rows']):
            metrics['water_issues'] = sum(row['status'] != 'ok' for row in value['rows'])
        if module == 'edo' and isinstance(value.get('missing_data_issues'), list):
            metrics['issues'] = len(value['missing_data_issues'])
    elif module == 'overdue':
        rows = value.get('items')
        metrics = {}
        if isinstance(rows, list) and all(isinstance(row,dict) and type(row.get('overdue_count')) is int and row['overdue_count'] >= 0 for row in rows):
            metrics = {'overdue': sum(row['overdue_count'] for row in rows), 'rows': len(rows)}
    elif module == 'cameras':
        summary = value.get('summary') or {}
        metrics = {k: summary[k] for k in ('working','not_working','not_connected','unknown') if k in summary}
    elif module == 'cds':
        metrics = {'appeals': len(value['data'])} if isinstance(value.get('data'), list) else {}
        return 'measured', metrics, safe_params({'start':value.get('date_from'), 'end':value.get('date_to')})
    elif module == 'utnkr':
        if value.get('source_mode') == 'local_rows':
            return 'unconfirmed', {}, {'operation':'local_rows'}
        metrics = {'violators': value.get('violators_count'), 'rows': value.get('rows_count')}
    elif module == 'mgkh_rm':
        buckets = value.get('buckets') or {}
        metrics = {k:len(buckets[k]) for k in ('close','extend','rework') if isinstance(buckets.get(k),list)}
    elif module == 'camera_prescriptions':
        metrics = {'documents':value.get('generated_count'),'problems':value.get('problem_count')}
        return 'measured', metrics, safe_params({'operation': value.get('mode')})
    elif module in ('water_dashboard','water-dashboard'):
        from services.check_water_snapshot import project
        metrics = project(value, params)
        return ('partial' if metrics.get('failed_sources') else 'measured'), metrics, {}
    elif module == 'edds':
        days = value.get('days')
        metrics = {}
        if isinstance(days,dict):
            counts = [counts for cities in days.values() if isinstance(cities,dict) for counts in cities.values()]
            if all(isinstance(cities,dict) for cities in days.values()) and all(
                    isinstance(n,list) and len(n)==3 and all(type(v) is int and v>=0 for v in n) for n in counts):
                metrics = {'days':len(days), 'appeals':sum(sum(n) for n in counts)}
        # Coverage can change; unlike-period cumulative totals are not comparable.
        return 'measured', metrics, safe_params({'operation':'daily','start':value.get('from'),'end':value.get('to')})
    else:
        metrics = {}
    return 'measured', metrics, {}


def captured(module, before, started_at, params):
    after = signature(module)
    if not after or after == before or after[0] / 1e9 < datetime.fromisoformat(started_at).timestamp() - 1:
        return 'unconfirmed', {}, {}
    try:
        if after[1] > MAX_BYTES:
            return 'unconfirmed', {}, {}
        with (BASE / PATHS[module]).open('rb') as stream:
            raw = stream.read(MAX_BYTES + 1)
        if len(raw) > MAX_BYTES or after != signature(module):
            return 'unconfirmed', {}, {}
        value = json.loads(raw.decode('utf-8-sig'))
        if not isinstance(value, dict):
            return 'unconfirmed', {}, {}
        # A manual edit may rewrite an older collector artifact. Reject it when
        # the embedded generation timestamp predates this run (second precision).
        stamp = next((value.get(k) for k in ('created_at','checked_at','updated_at','timestamp','generated_at') if value.get(k)), None)
        if stamp and module not in ('water-dashboard','water_dashboard'):
            dt = datetime.fromisoformat(str(stamp))
            if dt.timestamp() + 1 < datetime.fromisoformat(started_at).timestamp():
                return 'unconfirmed', {}, {}
        code, metrics, derived = _project(module, value, params)
        if any(k in params and params[k] != v for k, v in derived.items()):
            return 'unconfirmed', {}, {}
        clean = safe_metrics(metrics)
        return (code if clean or code in ('failed','partial') else 'unconfirmed'), clean, derived
    except (OSError, ValueError, TypeError, OverflowError, KeyError):
        return 'unconfirmed', {}, {}
