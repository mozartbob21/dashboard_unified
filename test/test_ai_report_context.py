"""Local AI reporting contracts; no AI service or portal calls are made."""
import json
from pathlib import Path
from unittest.mock import patch

import pytest

from services import water_ai_context as water
from services.aichat import report_context as reports


def sample_snapshot(value=31):
    return {'schema_version': 2, 'checked_at': '2026-10-01T18:00:00', 'sources': {
        'tasks': {'ok': True, 'metric_schema': 1, 'updated_at': '2026-10-01T17:59:00',
                  'data_date': '2026-09-30', 'metrics': [{'id': 'overdue_tasks', 'label': 'Просроченные задачи', 'value': 99, 'unit': ''}],
                  'details': {'schema_version': 1, 'basis': 'Больше задач — хуже', 'entities': [{'name': 'Солнечногорск', 'value': value, 'unit': ''}],
                              'groups': [{'title': 'Отстающие', 'items': [{'name': 'Солнечногорск', 'value': value, 'unit': ''}]}]}},
        'flush': {'ok': True, 'metrics': [{'label': 'Retired secret', 'value': 88}]},
    }, 'table': [{'name': 'Солнечногорск', 'tasks': value, 'att': 55.5}, {'name': 'Балашиха', 'tasks': 68, 'att': 80}]}


@pytest.fixture
def local_data(tmp_path, monkeypatch):
    monkeypatch.setattr(reports, 'DATA_DIR', tmp_path)
    monkeypatch.setattr(water, 'load_snapshot', lambda: sample_snapshot())
    return tmp_path


def test_no_access_does_not_read_sources(monkeypatch):
    with patch.object(water, 'load_snapshot') as load, patch.object(reports, '_load_result') as load_other:
        bundle = reports.prepare('Дай сводку', [])
    assert bundle['sources'] == []
    assert bundle['clarification']
    load.assert_not_called()
    load_other.assert_not_called()


def test_removed_source_never_appears_even_in_old_snapshot(local_data):
    bundle = reports.prepare('Сводный дашборд', ['water-dashboard'])
    assert len(bundle['sources']) == 7
    assert 'flush' not in {s['id'] for s in bundle['sources']}
    assert 'Retired secret' not in reports.to_prompt(bundle)


def test_stale_values_keep_real_dates(local_data, monkeypatch):
    snap = sample_snapshot()
    snap['sources']['tasks'].update(ok=False, error='Портал не ответил')
    monkeypatch.setattr(water, 'load_snapshot', lambda: snap)
    source = reports.prepare('РМ МИНЖКХ', ['water-dashboard'])['sources'][0]
    assert source['status'] == 'stale'
    assert source['collected_at'] == '2026-10-01T17:59:00'
    assert source['data_date'] == '2026-09-30'
    assert source['metrics'][0]['value'] == 99
    assert source['warning'] == 'Портал не ответил'


def test_report_municipality_is_not_region_total(local_data):
    bundle = reports.prepare('Сформируй отчёт по Солнечногорску', ['water-dashboard'])
    assert bundle['selection']['municipality'] == 'Солнечногорск'
    source = next(s for s in bundle['sources'] if s['id'] == 'tasks')
    assert source['scope'] == 'municipality'
    assert source['metrics'] == []
    assert source['rows'][0]['metrics'][0]['value'] == 31
    assert source['entities'][0]['value'] == 31
    assert 'Балашиха' not in reports.to_prompt(bundle)


def test_general_module_filters_and_projects_fields(local_data):
    path = local_data / 'overdue/final_result.json'; path.parent.mkdir()
    path.write_text(json.dumps({'created_at': '2026-09-28', 'password': 'DO_NOT_INCLUDE', 'personal_messages': ['private'], 'items': [
        {'municipality': 'Солнечногорск', 'overdue_count': 0, 'responsible_phone': 'SECRET_PHONE', 'token': 'TOKEN', 'status': ''},
        {'municipality': 'Балашиха', 'overdue_count': 9, 'status': 'unknown'},
    ]}))
    bundle = reports.prepare('Просроченные задачи по Солнечногорску', ['overdue'])
    source = bundle['sources'][0]
    assert source['collected_at'] == '2026-09-28'
    assert source['matching_rows'] == 1
    assert source['metrics'][1]['value'] == 0
    assert source['status_counts'] == {'Статус не указан': 1}
    payload = reports.to_prompt(bundle)
    assert all(secret not in payload for secret in ['DO_NOT_INCLUDE', 'SECRET_PHONE', 'TOKEN', 'private', 'Балашиха'])


def test_followup_uses_selection_but_reloads_values(local_data, monkeypatch):
    first = reports.prepare('РМ МИНЖКХ по Солнечногорску', ['water-dashboard'])
    monkeypatch.setattr(water, 'load_snapshot', lambda: sample_snapshot(44))
    second = reports.prepare('Какие выводы?', ['water-dashboard'], previous=first['selection'])
    assert second['selection'] == first['selection']
    assert second['sources'][0]['rows'][0]['metrics'][0]['value'] == 44


def test_permission_is_rechecked_for_shared_followup(monkeypatch):
    with patch.object(water, 'load_snapshot') as load:
        bundle = reports.prepare('Какие выводы?', ['edo'], previous={'module': 'water-dashboard', 'source': 'tasks', 'municipality': 'Солнечногорск'})
    assert bundle['clarification']
    assert not bundle['sources']
    load.assert_not_called()


def test_explicit_new_module_resets_previous_source(local_data):
    bundle = reports.prepare('Посмотри этот блок', ['water-dashboard', 'cameras'], requested={'module': 'cameras'}, previous={'module': 'water-dashboard', 'source': 'tasks'})
    assert bundle['selection']['module'] == 'cameras'
    assert not bundle['selection']['source']
    assert bundle['sources'][0]['id'] == 'cameras'


def test_unknown_municipality_is_clarification_not_region_fallback(local_data):
    bundle = reports.prepare('Отчёт', ['water-dashboard'], requested={'municipality': 'Несуществующий округ'})
    assert bundle['clarification']
    assert not bundle['sources']


def test_unselected_this_module_requests_clarification(local_data):
    assert reports.prepare('Дай информацию об этом блоке', ['water-dashboard'])['clarification']


def test_user_can_clear_municipality_for_region(local_data):
    bundle = reports.prepare('Теперь вся область', ['water-dashboard'], requested={'municipality': '*'}, previous={'module': 'water-dashboard', 'municipality': 'Балашиха'})
    assert bundle['selection']['municipality'] == ''
    assert bundle['sources'][0]['scope'] == 'region'


def test_legacy_global_widgets_are_not_used_as_facts(local_data, monkeypatch):
    snap = sample_snapshot()
    snap['sources']['tasks'].pop('metric_schema')
    snap['sources']['tasks']['widgets'] = [{'value': 123456}]
    monkeypatch.setattr(water, 'load_snapshot', lambda: snap)
    source = reports.prepare('РМ МИНЖКХ', ['water-dashboard'])['sources'][0]
    assert source['status'] == 'missing'
    assert source['metrics'] == []
    assert '123456' not in reports.to_prompt({'sources': [source]})


def test_truncation_is_explicit_and_json_stays_valid(local_data, monkeypatch):
    snap = sample_snapshot()
    for spec in water.source_catalog():
        snap['sources'][spec['id']] = {'ok': True, 'metric_schema': 1,
            'metrics': [{'id': str(i), 'label': 'Показатель ' * 30, 'value': i, 'unit': '%'} for i in range(20)],
            'tables': [{'headers': ['ОМСУ', 'Организация'], 'rows': [['Солнечногорск', 'Организация ' * 40]] * 100}]}
    monkeypatch.setattr(water, 'load_snapshot', lambda: snap)
    bundle = reports.prepare('Сводный дашборд', ['water-dashboard'])
    serialized = reports.to_prompt(bundle)
    assert len(serialized) <= reports.MAX_CONTEXT_CHARS
    assert len(json.loads(serialized)['sources']) == 7
    assert any('сокращена' in x for x in bundle['limitations'])
    assert any(s.get('omitted_metrics') for s in bundle['sources'])


def test_safe_table_projection_excludes_credentials_and_unmatched_municipality():
    source = {'tables': [{'headers': ['ОМСУ', 'Организация', 'Пароль организации', 'Кол-во задач'],
                           'rows': [['А', 'РСО', 'SECRET', 2], ['Б', 'Другая', 'SECRET', 3]]}]}
    result = water.selected_tables(source, 'А')
    assert result[0]['columns'] == ['ОМСУ', 'Организация', 'Кол-во задач']
    assert result[0]['rows'] == [['А', 'РСО', '2']]
    assert 'SECRET' not in json.dumps(result)


def test_zero_is_a_value_and_missing_is_not_zero(local_data, monkeypatch):
    snap = sample_snapshot()
    snap['sources']['tasks']['metrics'][0]['value'] = 0
    monkeypatch.setattr(water, 'load_snapshot', lambda: snap)
    reports_list = reports.prepare('Сводный дашборд', ['water-dashboard'])['sources']
    tasks = next(s for s in reports_list if s['id'] == 'tasks')
    missing = next(s for s in reports_list if s['id'] == 'nvos')
    assert tasks['status'] == 'current'
    assert tasks['metrics'][0]['value'] == 0
    assert missing['status'] == 'missing'
    assert missing['metrics'] == []


def test_invalid_local_file_has_no_raw_exception_or_content(local_data):
    path = local_data / 'edo/result.json'; path.parent.mkdir(); path.write_text('TOKEN=secret')
    bundle = reports.prepare('Заполненность данных', ['edo'])
    assert bundle['sources'][0]['status'] == 'missing'
    assert 'secret' not in reports.to_prompt(bundle)


def test_storage_keeps_shared_dialog_and_only_safe_scope_metadata(tmp_path, monkeypatch):
    from services.aichat import storage
    monkeypatch.setattr(storage, 'DATA_DIR', tmp_path)
    monkeypatch.setattr(storage, 'DIALOGS_FILE', tmp_path / 'dialogs.json')
    dialog = storage.create_dialog()
    storage.append_message(dialog['id'], 'user', 'Отчёт', report_scope={'module': 'water-dashboard', 'source': 'tasks', 'municipality': 'Балашиха', 'token': 'SECRET'})
    saved = storage.get_dialog(dialog['id'])
    assert saved['messages'][0]['report_scope'] == {'module': 'water-dashboard', 'source': 'tasks', 'municipality': 'Балашиха'}
    assert storage.list_dialogs()[0]['id'] == dialog['id']
    assert 'SECRET' not in json.dumps(saved)


def test_engine_receives_factual_context_separately_from_conversation():
    from services.aichat import engine
    with patch.object(engine, '_qwen_chat', return_value='Отчёт') as qwen:
        assert engine.ask([{'role': 'user', 'content': 'Отчёт', 'report_scope': {'module': 'secret'}}], platform_context='{"sources":[]}') == 'Отчёт'
    sent = qwen.call_args.args[0]
    assert sent[1]['role'] == 'system'
    assert '<platform_data>' in sent[2]['content']
    assert sent[2]['role'] == 'user'
    assert 'null/отсутствие строк' in sent[1]['content']
    assert sent[-1] == {'role': 'user', 'content': 'Отчёт'}


def test_actual_local_water_snapshot_is_readable_and_typed():
    # Uses only saved data, and remains valid in a fresh checkout without a snapshot.
    snap = water.load_snapshot()
    sources = water.source_reports(snap)
    assert len(sources) == 7
    for source in sources:
        assert source['id'] != 'flush'
        assert source['status'] in {'current', 'stale', 'missing'}
        assert source['source_url'].startswith('https://datalens.yandex/')
        assert 'text' not in source and 'widgets' not in source


def test_unrecognized_result_does_not_claim_zero_rows(local_data):
    path = local_data / 'edo/result.json'; path.parent.mkdir()
    path.write_text(json.dumps({'error': 'password=DO_NOT_INCLUDE', 'unrecognized': []}))
    source = reports.prepare('Заполненность данных', ['edo'])['sources'][0]
    assert source['status'] == 'missing'
    assert source['metrics'] == []
    assert 'DO_NOT_INCLUDE' not in json.dumps(source)


def test_plain_water_supply_selects_its_source_without_narrowing_full_dashboard(local_data):
    assert reports.prepare('Водоснабжение: кто отстаёт?', ['water-dashboard'])['selection']['source'] == 'sys_vs'
    assert reports.prepare('Сводный дашборд качества водоснабжения', ['water-dashboard'])['selection']['source'] == ''


def test_edo_municipality_filters_real_rso_entities(local_data, monkeypatch):
    snap = sample_snapshot()
    snap['sources']['edo_rso'] = {'ok': True, 'metric_schema': 1, 'metrics': [{'label': 'Областная доля ЭЦП', 'value': 61, 'unit': '%'}],
        'details': {'schema_version': 1, 'entities': [
            {'name': 'РСО Один', 'municipality': 'Балашиха', 'value': 55, 'unit': '%', 'secondary': []},
            {'name': 'РСО Два', 'municipality': 'Солнечногорск', 'value': 77, 'unit': '%', 'secondary': []}]}}
    monkeypatch.setattr(water, 'load_snapshot', lambda: snap)
    source = reports.prepare('Переход РСО на ЭДО по Балашихе', ['water-dashboard'])['sources'][0]
    assert source['metrics'] == []
    assert [e['name'] for e in source['entities']] == ['РСО Один']
    assert source['entities'][0]['value'] == 55
    assert 'РСО Один' not in reports.catalog(['water-dashboard'])['municipalities']


def test_common_plural_municipality_inflections(local_data, monkeypatch):
    snap = sample_snapshot()
    snap['table'].append({'name': 'Химки', 'tasks': 1})
    monkeypatch.setattr(water, 'load_snapshot', lambda: snap)
    assert reports.prepare('Дай информацию о водоснабжении по Химкам', ['water-dashboard'])['selection'] == {
        'module': 'water-dashboard', 'source': 'sys_vs', 'municipality': 'Химки',
        'modules': ['water-dashboard'], 'sources': ['sys_vs'], 'municipalities': ['Химки'], 'group_by': ''}
