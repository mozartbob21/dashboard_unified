import gzip
import json
from pathlib import Path
from unittest.mock import patch

import pytest

from services import water_ai_context as water
from services.aichat import report_context as reports, local_sources as local
from services.aichat import storage


@pytest.fixture
def data(tmp_path, monkeypatch):
    monkeypatch.setattr(reports, 'DATA_DIR', tmp_path)
    registry = {'municipalities': {n: {} for n in ['Власиха', 'Химки', 'Щёлково', 'Балашиха']}}
    (tmp_path / 'municipality_registry.json').write_text(json.dumps(registry))
    monkeypatch.setattr(water, 'load_snapshot', lambda: {'schema_version': 3, 'sources': {}, 'table': [{'name': 'Власиха', 'tasks': 3}, {'name': 'Химки', 'tasks': 9}]})
    return tmp_path


def write(root, path, value):
    p = root / path; p.parent.mkdir(parents=True, exist_ok=True)
    if p.suffix == '.gz':
        p.write_bytes(gzip.compress(json.dumps(value).encode()))
    else:
        p.write_text(json.dumps(value))


@pytest.mark.parametrize('question', ['Дай краткий отчёт по всем блокам', 'Дай отчёт по Власихе', 'Покажи камеры и просрочку по Химкам', 'Что в критичных по технадзору?', 'Дай отчёт по каждому муниципалитету в блоках ЕДДС и МинЖКХ'])
def test_plain_request_activates_reports(data, question):
    assert reports.wants_context(question, list(reports.MODULES))


@pytest.mark.parametrize('question,has_files', [('Привет', False), ('Спасибо', False), ('Напиши сказку о камерах', False), ('Сделай отчёт по файлу', True), ('Переведи этот документ', True), ('Как написать отчёт о работе?', False)])
def test_ordinary_chat_and_files_do_not_load_context(data, question, has_files):
    assert not reports.wants_context(question, list(reports.MODULES), has_files=has_files)


def test_all_modules_clear_sticky_scope(data):
    previous = {'modules': ['water-dashboard'], 'sources': ['tasks'], 'municipalities': ['Власиха']}
    scope = reports.resolve_scope('Дай краткий отчёт по всем блокам', ['cameras', 'edds'], previous=previous)
    assert scope['modules'] == [] and scope['sources'] == [] and scope['municipalities'] == []
    bundle = reports.prepare('Дай краткий отчёт по всем блокам', ['cameras', 'edds'], previous=previous)
    assert {s['module'] for s in bundle['sources']} == {'cameras', 'edds'}


def test_new_municipality_means_all_granted_modules(data):
    previous = {'modules': ['water-dashboard'], 'sources': ['tasks'], 'municipalities': ['Химки']}
    bundle = reports.prepare('Дай отчёт по Власихе', ['water-dashboard', 'cameras', 'edds'], previous=previous)
    assert bundle['selection']['modules'] == []
    assert bundle['selection']['sources'] == []
    assert bundle['selection']['municipalities'] == ['Власиха']
    assert {s['module'] for s in bundle['sources']} == {'water-dashboard', 'cameras', 'edds'}


def test_multiple_blocks_and_municipalities(data):
    b = reports.prepare('Дай отчёт по ЕДДС и МинЖКХ по Власихе и Химкам', ['edds', 'mingkh', 'cameras'])
    assert set(b['selection']['modules']) == {'edds', 'mingkh'}
    assert set(b['selection']['municipalities']) == {'Власиха', 'Химки'}
    assert {s['module'] for s in b['sources']} == {'edds', 'mingkh'}
    assert all('municipality_table' in s for s in b['sources'])


@pytest.mark.parametrize('q', ['Дай отчёт по каждому муниципалитету в блоках ЕДДС и МинЖКХ', 'Дай отчёт по всем блокам с разбивкой по муниципалитетам', 'Сводка по всем блокам в разрезе округов'])
def test_group_by_every_municipality(data, q):
    b = reports.prepare(q, ['edds', 'mingkh'])
    assert b['selection']['group_by'] == 'municipality'
    assert b['coverage']['municipalities_requested'] == 4
    assert all(s['scope'] == 'municipalities' for s in b['sources'])


def test_region_reset_and_neuter_municipality(data):
    previous = {'modules': ['water-dashboard'], 'sources': ['tasks'], 'municipalities': ['Власиха']}
    assert reports.resolve_scope('Теперь вся область', ['water-dashboard'], previous=previous)['municipalities'] == []
    assert reports.resolve_scope('Дай общий отчёт по области', ['water-dashboard'], previous=previous)['modules'] == []
    assert reports.resolve_scope('Дай отчёт по Щёлкову', ['water-dashboard'], previous=previous)['municipalities'] == ['Щёлково']
    assert reports.prepare('Дай отчёт по муниципалитету Бирюлево', ['water-dashboard'], previous=previous)['clarification']


def test_city_identity_deduplicates_and_matches_rows(data):
    write(data, 'cameras/state/dashboard_state.json', {'updated_at': '2026-10-01', 'rows': [
        {'municipality': 'г.о. Химки', 'camera_status': 'offline'}, {'municipality': 'Химки', 'camera_status': 'online'},
        {'municipality': 'Власиха (ЗАТО)', 'camera_status': 'online'}]})
    opts = reports.catalog(['cameras'])
    assert opts['municipalities'] == ['Власиха', 'Химки']
    b = reports.prepare('Покажи камеры по Химкам', ['cameras'])
    assert b['sources'][0]['metrics'][0]['value'] == 2


def test_grouped_camera_data_excludes_region_totals_and_keeps_city_status(data):
    write(data, 'cameras/state/dashboard_state.json', {'updated_at': '2026-10-01', 'rows': [
        {'municipality': 'Химки', 'camera_status': 'offline'}, {'municipality': 'Власиха', 'camera_status': 'online'},
        {'municipality': 'Балашиха', 'camera_status': 'other-city'}]})
    s = reports.prepare('Камеры по Власихе и Химкам', ['cameras'])['sources'][0]
    assert 'status_counts' not in s and 'matching_rows' not in s
    assert 'other-city' not in reports.to_prompt({'source': s})
    assert 'Статус: offline' in s['municipality_table']['columns']


def test_missing_city_counts_are_unknown_for_external_sources(data):
    write(data, 'edds/water_daily.json', {'updated': '2026-10-01', 'days': {'2026-10-01': {'Химки': [1, 2, 3]}}})
    write(data, 'mingkh/report-summary.json', {'schema_version': 1, 'collected_at': '2026-10-01', 'periods': {'curr': {'bounds': ['2026-09-01', '2026-09-30'], 'total': 5, 'municipalities': {'Химки': 5}}}})
    write(data, 'mingkh/water-map.json.gz', {'meta': {'updated': '2026-10-01'}, 'cols': ['omsu', 'kind'], 'omsu': ['Химки'], 'rows': [[0, 0]]})
    b = reports.prepare('ЕДДС и МинЖКХ по Власихе', ['edds', 'mingkh'])
    for s in b['sources']:
        assert all(m['value'] is None for m in s['metrics'])
        assert s['warning']
    grouped = reports.prepare('ЕДДС и МинЖКХ по Власихе и Химкам', ['edds', 'mingkh'])
    assert 'Примечание' in grouped['sources'][0]['municipality_table']['columns']


def test_mingkh_projection_is_safe_persistent_and_specific(data, monkeypatch):
    monkeypatch.setattr(local, 'BASE', data)
    d = {'dims': {'omsu': ['г.о. Химки', 'Власиха'], 'secret': ['TOKEN']}, 'rows': {'curr': [[0, 8], [1, 9]], 'prev': [[0, 3]]}, 'updated': '01.10.2026', 'password': 'SECRET', 'bounds': {'curr': ['2026-09-01', '2026-09-30'], 'prev': ['2026-08-01', '2026-08-31']}}
    assert local.persist_mingkh_dataset(d)
    p = data / 'data/mingkh/report-summary.json'
    saved = json.loads(p.read_text())
    assert saved['periods']['curr']['municipalities']['Химки'] == 1
    assert 'TOKEN' not in p.read_text() and 'SECRET' not in p.read_text()
    report = local.reports_for_module('mingkh', data / 'data', 'Власиха')[0]
    assert report['metrics'][0]['value'] == 1
    assert report['metrics'][1]['value'] is None
    assert report['periods']['curr'] == ['2026-09-01', '2026-09-30']


def test_arm_summary_is_separate_from_complaints(data):
    counts = {'incidents': 2, 'closed': 1, 'active': 1, 'people_known_sum': 30, 'people_missing': 1}
    write(data, 'edds/arm-summary.json', {'schema_version': 1, 'collected_at': '2026-10-01', 'period': {'from': '2026-09-01', 'to': '2026-09-30'}, 'totals': counts, 'municipalities': {'Власиха (ЗАТО)': counts}, 'broken_rows': 1, 'unknown_municipality': 0})
    s = next(s for s in reports.prepare('ЕДДС по Власихе', ['edds'])['sources'] if s['id'] == 'edds-arm')
    assert s['metrics'][0]['value'] == 2
    assert 'повреждённых' in s['warning']
    assert s['period']['from'] == '2026-09-01'


def test_published_zip_counts_do_not_sum_incompatible_units(data):
    write(data, 'zip_curator/published.json', {'published_at': '2026-10-01', 'rows': [['РСО', 'Округ', 'Наименование', 'Количество', 'Ед.'], ['РСО 1', 'Власиха', 'Труба', 100, 'м'], ['РСО 1', 'Власиха', 'Насос', 4, 'шт']]})
    s = reports.prepare('Остатки ЗиП по Власихе', ['zips'])['sources'][0]
    assert s['metrics'][0]['value'] == 2
    assert all(m['value'] != 104 for m in s['metrics'])


def test_storage_preserves_multi_scope_but_not_secrets(data, monkeypatch):
    monkeypatch.setattr(storage, 'DATA_DIR', data)
    monkeypatch.setattr(storage, 'DIALOGS_FILE', data / 'dialogs.json')
    d = storage.create_dialog()
    storage.append_message(d['id'], 'user', 'Отчёт', report_scope={'modules': ['edds', 'mingkh'], 'municipalities': ['Власиха', 'Химки'], 'group_by': 'municipality', 'password': 'SECRET'})
    scope = storage.get_dialog(d['id'])['messages'][0]['report_scope']
    assert scope['modules'] == ['edds', 'mingkh']
    assert scope['municipalities'] == ['Власиха', 'Химки']
    assert 'SECRET' not in json.dumps(scope)


def test_multi_scope_followup_rechecks_grants(data):
    b = reports.prepare('Подробнее', ['cameras'], previous={'modules': ['cameras', 'edds'], 'municipalities': ['Власиха', 'Химки']})
    assert {s['module'] for s in b['sources']} == {'cameras'}
    assert 'edds' in b['selection']['unavailable_modules']


def test_grouping_reads_general_result_once_after_scope_resolution(data):
    source = {'updated_at': '2026-10-01', 'rows': [{'municipality': 'Химки', 'status': 'ok'}]}
    scope = {'modules': ['cameras'], 'sources': [], 'municipalities': ['Химки', 'Власиха', 'Балашиха'], 'group_by': ''}
    with patch.object(reports, 'resolve_scope', return_value=scope), patch.object(reports, '_load_result', return_value=(source, '')) as load:
        reports.prepare('отчёт', ['cameras'])
    load.assert_called_once_with('cameras')


def test_grouped_fallback_contains_actual_city_values(data):
    write(data, 'cameras/state/dashboard_state.json', {'updated_at': '2026-10-01', 'rows': [{'municipality': 'Химки', 'camera_status': 'offline'}, {'municipality': 'Власиха', 'camera_status': 'online'}]})
    answer = reports.fallback(reports.prepare('Камеры по Власихе и Химкам', ['cameras']))
    assert 'Химки' in answer and 'Власиха' in answer and 'offline: 1' in answer


@pytest.mark.parametrize('name', ['Голицыно', 'Можайск', 'Москва'])
def test_municipality_prefix_does_not_cut_real_names(name):
    from services.report_municipalities import key, display
    assert key(name) == name.casefold()
    assert display(name) == name


def test_prefix_without_space_is_recognized():
    from services.report_municipalities import key
    assert key('г.о.Химки') == key('м.о. Химки') == key('Химки')


def test_region_followup_activates_and_source_aliases_are_specific(data):
    assert reports.wants_context('Теперь вся область', ['water-dashboard'], previous={'module': 'water-dashboard'})
    assert reports.resolve_scope('Дай отчёт по заявкам ЕДДС', ['edds', 'water-dashboard'])['modules'] == ['edds']
    assert reports.resolve_scope('Дай отчёт по ЭДО', ['water-dashboard'])['sources'] == ['edo_rso']
    assert reports.resolve_scope('Заполненность данных', ['edo', 'water-dashboard'])['modules'] == ['edo']


def test_grouped_context_bound_preserves_every_selected_module(data):
    cities = ['Муниципалитет' + str(n) for n in range(80)]
    write(data, 'municipality_registry.json', {'municipalities': {n: {} for n in cities}})
    write(data, 'cameras/state/dashboard_state.json', {'updated_at': '2026-10-01', 'rows': [
        {'municipality': name, 'camera_status': 'Очень длинный статус ' * 50} for name in cities]})
    b = reports.prepare('Дай отчёт по всем блокам по каждому муниципалитету', list(reports.MODULES))
    assert len(reports.to_prompt(b)) <= reports.MAX_CONTEXT_CHARS
    assert {s['module'] for s in b['sources']} == set(reports.MODULES)
    assert any(s.get('municipality_table', {}).get('omitted_rows') for s in b['sources'])


@pytest.mark.parametrize('question', ['А теперь по Химкам', 'Теперь по Власихе', 'А по камерам?'])
def test_short_scope_followup_activates_context(data, question):
    assert reports.wants_context(question, ['water-dashboard', 'cameras'], previous={'modules': ['water-dashboard']})


@pytest.mark.parametrize('municipality', ['', 'Власиха'])
def test_arm_floating_population_sum_is_retained(data, municipality):
    counts = {'incidents': 2, 'closed': 1, 'active': 1, 'people_known_sum': 12.5, 'people_missing': 0}
    summary = {'schema_version': 1, 'collected_at': '2026-10-01', 'totals': counts, 'municipalities': {'Власиха': counts}}
    report = local._arm(summary, municipality)
    assert report['metrics'][3]['value'] == 12.5



def test_registry_titles_normalize_case_and_zato_without_losing_identity(tmp_path, monkeypatch):
    from services import report_municipalities as municipalities
    path = tmp_path / 'registry.json'
    path.write_text(json.dumps({'municipalities': {'ХИМКИ': {}, 'Власиха (ЗАТО)': {}}}))
    monkeypatch.setattr(municipalities, 'REGISTRY', path)
    assert municipalities.display('Химки') == 'Химки'
    assert municipalities.display('Власиха') == 'Власиха'
    assert municipalities.key('Власиха (ЗАТО)') == municipalities.key('Власиха')


@pytest.mark.parametrize('name,question', [
    ('Серебряные Пруды', 'Дай отчёт по Серебряным Прудам'),
    ('Павловский Посад', 'Дай отчёт по Павловскому Посаду'),
    ('Шаховская', 'Дай отчёт по Шаховской'),
    ('Электросталь', 'Дай отчёт по Электростали'),
    ('Восход', 'Дай отчёт по Восходу'),
])
def test_compound_feminine_and_consonant_municipality_cases(data, name, question):
    write(data, 'municipality_registry.json', {'municipalities': {name: {}}})
    assert reports.resolve_scope(question, ['edds'])['municipalities'] == [name]


def test_municipality_matching_uses_word_boundaries():
    assert reports._named_municipalities('дай отчет по коломенскому', ['Коломна']) == []
    assert reports._named_municipalities('дай отчет по большому городу', ['Город']) == ['Город']
    assert reports._named_municipalities('дай отчет по пригородам', ['Город']) == []



def test_local_file_cache_reloads_changed_data_and_reports_do_not_mutate_it(data):
    value = {'schema_version': 1, 'collected_at': '2026-10-01', 'periods': {'curr': {'bounds': ['2026-09-01', '2026-09-30'], 'total': 5, 'municipalities': {'Химки': 5}}}}
    write(data, 'mingkh/report-summary.json', value)
    first = local.reports_for_module('mingkh', data, 'Химки')[0]
    first['periods']['curr'][0] = 'changed'
    first['metrics'][0]['value'] = 999
    unchanged = local.reports_for_module('mingkh', data, 'Химки')[0]
    assert unchanged['periods']['curr'][0] == '2026-09-01'
    assert unchanged['metrics'][0]['value'] == 5
    value['periods']['curr']['municipalities']['Химки'] = 27
    write(data, 'mingkh/report-summary.json', value)
    assert local.reports_for_module('mingkh', data, 'Химки')[0]['metrics'][0]['value'] == 27



def test_catalog_excludes_aggregates_and_deduplicates_suffixes(data):
    names = ['', 'null', 'None', 'Итого', 'Москва', 'Московская область', 'Можайский/Рузский',
             'Богородский', 'Богородский г.о.', 'Раменский', 'Раменский м.о.', 'Можайск', 'Можайский']
    write(data, 'cameras/state/dashboard_state.json', {'rows': [{'municipality': n} for n in names]})
    assert reports.catalog(['cameras'])['municipalities'] == ['Богородский', 'Можайск', 'Можайский', 'Раменский']
    from services.report_municipalities import key, display
    assert key('Богородский г.о.') == key('Богородский')
    assert display('Раменский м.о.') == 'Раменский'


@pytest.mark.parametrize('question,city', [('А теперь по Химкам', 'Химки'), ('Теперь по Власихе', 'Власиха')])
def test_short_city_followup_preserves_selected_blocks(data, question, city):
    previous = {'modules': ['edds', 'mingkh'], 'sources': [], 'municipalities': ['Балашиха']}
    scope = reports.resolve_scope(question, ['edds', 'mingkh', 'water-dashboard'], previous=previous)
    assert scope['modules'] == ['edds', 'mingkh']
    assert scope['municipalities'] == [city]
    source_scope = reports.resolve_scope(question, ['water-dashboard'], previous={'modules': ['water-dashboard'], 'sources': ['tasks'], 'municipalities': ['Балашиха']})
    assert source_scope['sources'] == ['tasks']
    fresh = reports.resolve_scope('Дай отчёт по ' + ('Химкам' if city == 'Химки' else 'Власихе'), ['edds', 'mingkh', 'water-dashboard'], previous=previous)
    assert fresh['modules'] == []


def test_region_report_keeps_explicit_block(data):
    selected = reports.resolve_scope('Дай общий отчёт по области по ЕДДС', ['edds', 'mingkh'],
                                     previous={'modules': ['mingkh'], 'municipalities': ['Химки']})
    assert selected['modules'] == ['edds']
    assert selected['municipalities'] == []


@pytest.mark.parametrize('question', [
    'Дай информацию о нейросетях', 'Что такое отчёт?', 'Напиши пример отчёта о командировке',
    'Какие показатели важны для бизнеса?', 'Почему небо синее?',
    'Как работает водоснабжение?', 'Помоги выбрать камеры для офиса',
    'Напиши стихотворение о Власихе', 'Что делать, если компьютер завис?',
    'Дай рекомендации по изучению Python', 'Расскажи подробнее про фотосинтез',
])
@pytest.mark.parametrize('previous', [None, {'modules': ['water-dashboard'], 'municipalities': ['Власиха']}])
def test_general_assistant_requests_do_not_become_reports(data, question, previous):
    assert not reports.wants_context(question, list(reports.MODULES), previous=previous)


@pytest.mark.parametrize('question', ['Почему?', 'Почему так?', 'Подробнее', 'Какие выводы?', 'Что рекомендуешь?'])
def test_short_report_followups_still_read_current_data(data, question):
    assert reports.wants_context(question, ['water-dashboard'], previous={'modules': ['water-dashboard']})
