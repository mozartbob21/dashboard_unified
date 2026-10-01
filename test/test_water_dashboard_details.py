import copy

from services.water_dashboard.builder import build_table, normalize_snapshot
from services.water_dashboard.config import SOURCES
from services.water_dashboard.details import source_details, EDO_CURRENT_HEADER, EDO_PREVIOUS_HEADER


def table(headers, rows):
    return {'headers': headers, 'rows': rows}


def detail(sid, headers, rows):
    return source_details(sid, {'tables': [table(headers, rows)]})


def group(result, key):
    return next(item['items'] for item in result['groups'] if item['id'] == key)


def test_flush_is_absent_even_before_refresh_of_old_snapshot():
    old = {'sources': {
        'flush': {'metric_schema': 1, 'ok': True, 'updated_at': '2099-01-01', 'metrics': []},
        'tasks': {'metric_schema': 1, 'ok': True, 'updated_at': '2026-10-01',
                  'metrics': [{'id': 'overdue_tasks', 'value': 8}],
                  'tables': [table(['ОМСУ', 'Кол-во задач'], [['Округ', '8']])]},
    }, 'sources_updated': {'flush': True}, 'snapshot_date': '01.10.2026'}
    before = copy.deepcopy(old)
    normalized = normalize_snapshot(old)
    assert old == before
    assert 'flush' not in normalized['sources']
    assert 'flush' not in normalized['sources_updated']
    assert len(normalized['sources']) == len(SOURCES) == 7
    assert normalized['updated_at'] == '2026-10-01'
    assert group(normalized['sources']['tasks']['details'], 'worst')[0]['value'] == 8
    assert normalized['schema_version'] == 3
    assert normalize_snapshot(normalized) == normalized


def test_unverified_old_sources_cannot_supply_rankings():
    old = {'sources': {'tasks': {'tables': [table(['ОМСУ', 'Кол-во задач'], [['Округ', '999']])]}}}
    normalized = normalize_snapshot(old)
    assert normalized['sources']['tasks']['details']['entities'] == []
    assert normalized['table'] == []
    assert normalized['kpis']['tasks_total'] is None


def test_stale_verified_rankings_keep_original_date_and_failure():
    old = {'sources': {'meetings': {'metric_schema': 1, 'ok': False, 'error': 'Недоступно',
           'updated_at': '2026-09-01', 'checked_at': '2026-10-01',
           'tables': [table(['ОМСУ', 'Явка, %'], [['Первый', '0']])]}}}
    result = normalize_snapshot(old)['sources']['meetings']
    assert result['updated_at'] == '2026-09-01'
    assert result['checked_at'] == '2026-10-01'
    assert result['error'] == 'Недоступно' and not result['ok']
    assert group(result['details'], 'worst')[0]['value'] == 0


def test_valves_rank_by_real_status_not_plan_as_if_it_were_completion():
    result = detail('valves', ['ОМСУ', 'Кол-во задвижек планируемых к замене, шт', 'Статус занесения в zulugis'], [
        ['Большой план', '99999', 'Полностью'], ['Частично', '100', 'Частично'],
        ['Нет записи', '1', 'Не внесено'], ['Малый план', '2', 'Полностью'], ['Неизвестно', '10', '—'],
    ])
    assert group(result, 'worst')[0]['name'] == 'Нет записи'
    assert group(result, 'best')[0]['name'] == 'Большой план'
    assert group(result, 'best')[1]['name'] == 'Малый план'
    assert group(result, 'worst')[1]['value'] == 'Частично'
    assert all(item['unit'] == '' and '_score' not in item for item in result['entities'])
    assert result['coverage']['excluded'] == 1
    assert 'нет процента' in result['coverage']['note']


def test_count_rankings_limit_five_preserve_zero_and_exclude_unknown():
    result = detail('tasks', ['ОМСУ', 'Население', 'Кол-во задач'],
                    [[f'Округ {i}', '10000000', str(i)] for i in range(8)] + [['Без данных', '999', '—']])
    assert [item['value'] for item in group(result, 'worst')] == [7, 6, 5, 4, 3]
    assert len(result['entities']) == 8
    assert result['entities'][0]['value'] == 0
    assert result['coverage']['excluded'] == 1


def test_duplicate_tables_do_not_duplicate_places_conflicts_are_excluded():
    headers = ['ОМСУ', 'Кол-во задач']
    data = {'tables': [table(headers, [['ОКРУГ', '10'], ['Второй', '5']]),
                       table(headers, [['Округ', '10'], ['Второй', '7']])]}
    result = source_details('tasks', data)
    assert [(item['name'], item['value']) for item in result['entities']] == [('Округ', 10)]
    assert result['coverage']['excluded'] == 1
    built = build_table({'tasks': data})
    assert next(row for row in built if row['name'] == 'Второй')['tasks'] is None
    assert len(built) == 2


def test_system_rankings_do_not_sum_overlapping_categories_or_use_dynamics():
    result = detail('sys_vs', ['ОМСУ', 'Динамика системных за неделю', 'Системных', 'Резонансных'], [
        ['Первый', '999', '10', '50'], ['Второй', '1000', '20', '1'], ['Третий', '0', '0', '0'],
    ])
    assert [row['name'] for row in group(result, 'worst')] == ['Второй', 'Первый', 'Третий']
    assert group(result, 'worst')[0]['value'] == 20
    assert group(result, 'worst')[0]['secondary'][0]['value'] == 1
    assert 'не суммируются' in result['basis']


def test_capital_repairs_use_own_source():
    result = detail('sys_kr', ['ОМСУ', 'Системных', 'Резонансных'], [['Капремонт', '8', '20']])
    assert group(result, 'worst')[0]['value'] == 8


def test_edo_ranks_full_rso_table_by_current_week_not_ecp_or_previous_week():
    headers = ['РСО', 'ОМСУ', EDO_PREVIOUS_HEADER, EDO_CURRENT_HEADER]
    data = {'ranking_tables': [table(headers, [
        ['Водоканал А', 'Первый', '0', '100'], ['Водоканал Б', 'Второй', '100', '20'],
        ['Водоканал В', 'Третий', '40', '50'], ['Водоканал Г', 'Четвёртый', '90', '0'],
        ['Водоканал Д', 'Пятый', '10', '—'], ['Водоканал Е', 'Шестой', '50', '101'],
    ])], 'tables': [table(['ОМСУ', 'Кол-во должностных лиц с правом подписи', 'Наличие ЭЦП у подписывающих, кол-во'],
                         [['Второй', '1', '1']])]}
    result = source_details('edo_rso', data)
    assert group(result, 'best')[0]['name'] == 'Водоканал А'
    assert group(result, 'worst')[0]['name'] == 'Водоканал Г'
    assert group(result, 'worst')[1]['value'] == 20
    assert group(result, 'worst')[0]['municipality'] == 'Четвёртый'
    assert result['coverage']['excluded'] == 2
    assert '4 РСО' in result['coverage']['note']
    assert result['entity_type'] == 'rso'


def test_edo_legacy_ecp_laggards_cannot_produce_best_or_worst_ranking():
    result = detail('edo_rso', ['РСО', 'ОМСУ', 'Кол-во должностных лиц с правом подписи', 'Наличие ЭЦП у подписывающих, кол-во'],
                    [['Заполнено', 'Первый', '10', '2']])
    assert group(result, 'best') == [] and group(result, 'worst') == []
    assert 'Полная таблица РСО ещё не получена' in result['coverage']['note']


def test_edo_deduplicates_77_rows_but_keeps_same_rso_name_in_different_municipalities():
    headers = ['РСО', 'ОМСУ', EDO_PREVIOUS_HEADER, EDO_CURRENT_HEADER]
    rows = [['Водоканал', f'Округ {i}', str(i), str(i)] for i in range(77)]
    result = source_details('edo_rso', {'ranking_tables': [table(headers, rows), table(headers, rows)]})
    assert len(result['entities']) == 77
    assert len(group(result, 'best')) == len(group(result, 'worst')) == 5
    assert group(result, 'best')[0]['value'] == 76 and group(result, 'worst')[0]['value'] == 0


def test_edo_conflicting_rows_excluded_even_when_only_previous_week_differs():
    headers = ['РСО', 'ОМСУ', EDO_PREVIOUS_HEADER, EDO_CURRENT_HEADER]
    rows = [['Водоканал', 'Округ', '10', '20'], ['Водоканал', 'Округ', '30', '20']]
    result = source_details('edo_rso', {'ranking_tables': [table(headers, rows)]})
    assert result['entities'] == [] and result['coverage']['excluded'] == 1


def test_edo_rankings_persist_alongside_independent_overview_kpis(tmp_path, monkeypatch):
    from services.water_dashboard import builder
    monkeypatch.setattr(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json')
    data = {'widgets': [{'label': 'Доля (%) должностных лиц, имеющих право подписи и ЭЦП', 'value': '61 %'}],
            'ranking_tables': [table(['РСО', 'ОМСУ', EDO_CURRENT_HEADER], [['Водоканал', 'Округ', '83']])],
            'ranking_url': 'https://datalens.yandex/f5wqqij889haz?tab=EL'}
    source = builder.build_snapshot({'edo_rso': data}, source_ids=['edo_rso'])['sources']['edo_rso']
    assert source['metrics'][0]['value'] == 61
    assert source['details']['entities'][0]['value'] == 83
    assert source['ranking_tables'] == data['ranking_tables']
    assert source['details']['source_url'] == data['ranking_url']
    assert source['ok'] is True
    data.update(ranking_tables=[], ranking_error='Таблица временно недоступна')
    failed_ranking = builder.build_snapshot({'edo_rso': data}, source_ids=['edo_rso'])['sources']['edo_rso']
    assert failed_ranking['ok'] is True and failed_ranking['metrics'][0]['value'] == 61
    assert failed_ranking['details']['entities'] == []
    assert failed_ranking['details']['coverage']['note'] == 'Таблица временно недоступна'


def test_attendance_lowest_first_and_valid_zero():
    result = detail('meetings', ['ОМСУ', 'Присутствовали на перекличках, %', 'Число перекличек'], [
        ['Высокая', '99,5', '10'], ['Нулевая', '0', '1000'], ['Низкая', '12,3', '9999'],
        ['Некорректная', '101', '10'], ['Нет данных', '—', '12'],
    ])
    assert [row['value'] for row in group(result, 'worst')] == [0, 12.3, 99.5]
    assert result['coverage']['excluded'] == 2


def test_nvos_details_are_typed_mini_summary_not_fake_rankings():
    metrics = [{'id': 'samples_fact_week', 'label': 'Факт за неделю', 'value': 0, 'unit': 'шт.'},
               {'id': 'samples_plan_year', 'label': 'План на год', 'value': None, 'unit': 'шт.'}]
    result = source_details('nvos', {'metrics': metrics})
    assert result['kind'] == 'metrics' and result['groups'] == []
    assert result['metrics'] == metrics
    assert result['metrics'][0]['value'] == 0 and result['metrics'][1]['value'] is None


def test_single_source_refresh_keeps_other_source_values_status_and_dates(tmp_path, monkeypatch):
    import json
    from services.water_dashboard import builder
    monkeypatch.setattr(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json')
    first = builder.build_snapshot({
        'tasks': {'tables': [table(['ОМСУ', 'Кол-во задач'], [['Округ', '5']])]},
        'meetings': {'tables': [table(['ОМСУ', 'Явка, %'], [['Округ', '60']])]}
    })
    first['sources']['meetings'].update(updated_at='2026-01-01', checked_at='2026-02-01')
    builder.SNAPSHOT_FILE.write_text(json.dumps(first))
    old_meetings = copy.deepcopy(first['sources']['meetings'])
    second = builder.build_snapshot({'tasks': {'tables': [table(['ОМСУ', 'Кол-во задач'], [['Округ', '7']])]}}, source_ids=['tasks'])
    assert second['sources']['meetings'] == old_meetings
    assert second['last_checked_sources'] == ['tasks']
    assert second['last_updated_sources'] == ['tasks']
    assert second['kpis']['tasks_total'] == 7 and second['kpis']['att_avg'] == 60
    third = builder.build_snapshot({'tasks': {'error': 'Ошибка'}}, source_ids=['tasks'])
    assert third['last_checked_sources'] == ['tasks'] and third['last_updated_sources'] == []
    assert third['sources']['meetings'] == old_meetings
    assert third['sources']['tasks']['updated_at'] == second['sources']['tasks']['updated_at']


def test_single_source_can_create_first_snapshot_with_other_sources_unknown(tmp_path, monkeypatch):
    from services.water_dashboard import builder
    monkeypatch.setattr(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json')
    result = builder.build_snapshot({'tasks': {'tables': [table(['ОМСУ', 'Кол-во задач'], [['Округ', '7']])]}}, source_ids=['tasks'])
    assert len(result['sources']) == 7
    assert result['kpis']['tasks_total'] == 7
    assert result['sources']['meetings']['metrics'][0]['value'] is None
    assert result['sources']['meetings'].get('checked_at') is None


def test_source_schedule_is_reread_even_if_metrics_fail(tmp_path, monkeypatch):
    from services.water_dashboard import builder
    monkeypatch.setattr(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json')
    data = {'tables': [table(['ОМСУ', 'Кол-во задач'], [['Округ', '7']])],
            'text': 'Дашборд автоматически обновляется каждые 30 мин.'}
    first = builder.build_snapshot({'tasks': data}, source_ids=['tasks'])
    second = builder.build_snapshot({'tasks': {'error': 'Ошибка графика', 'refresh_text': 'Дашборд обновляется каждый час.'}}, source_ids=['tasks'])
    source = second['sources']['tasks']
    assert source['refresh'] == 'Дашборд обновляется каждый час'
    assert source['updated_at'] == first['sources']['tasks']['updated_at']
    assert source['ok'] is False and source['refresh_available'] is True
    third = builder.build_snapshot({'tasks': {'error': 'Нет сети'}}, source_ids=['tasks'])
    assert third['sources']['tasks']['refresh'] == source['refresh']
    assert third['sources']['tasks']['refresh_checked_at'] == source['refresh_checked_at']
    fourth = builder.build_snapshot({'tasks': {'error': 'Нет доступа', 'refresh_text': 'Войдите в аккаунт'}}, source_ids=['tasks'])
    assert fourth['sources']['tasks']['refresh'] == source['refresh']
    assert fourth['sources']['tasks']['refresh_checked_at'] == source['refresh_checked_at']


def test_fresh_source_without_schedule_clears_old_schedule(tmp_path, monkeypatch):
    from services.water_dashboard import builder
    monkeypatch.setattr(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json')
    data = {'tables': [table(['ОМСУ', 'Кол-во задач'], [['Округ', '7']])],
            'text': 'Дашборд обновляется каждый час.'}
    builder.build_snapshot({'tasks': data}, source_ids=['tasks'])
    data['text'] = 'Свод задач без примечания о расписании'
    result = builder.build_snapshot({'tasks': data}, source_ids=['tasks'])
    assert result['sources']['tasks']['refresh'] == ''
    assert result['sources']['tasks']['refresh_available'] is False


def test_schedule_parser_accepts_event_based_update_and_preserves_weekdays():
    from services.water_dashboard.builder import extract_refresh_info
    assert extract_refresh_info('Данные обновляются автоматически после переклички на совещании') == 'Данные обновляются автоматически после переклички на совещании'
    assert extract_refresh_info('Дашборд обновляется по вторникам и четвергам в 12:00') == 'Дашборд обновляется по вторникам и четвергам в 12:00'
    assert extract_refresh_info('Дашборд обновляется ежедневно с 9:00 до 10:30, названия ОМСУ кликабельны') == 'Дашборд обновляется ежедневно с 9:00 до 10:30'
    assert extract_refresh_info('Последнее обновление 01.10.2026') == ''


def test_rejects_unknown_and_empty_selection_before_writing(tmp_path, monkeypatch):
    import pytest
    from services.water_dashboard import builder
    monkeypatch.setattr(builder, 'SNAPSHOT_FILE', tmp_path / 'snapshot.json')
    for selected in ([], ['flush'], ['../../outside']):
        with pytest.raises(ValueError):
            builder.build_snapshot({}, source_ids=selected)
    assert not builder.SNAPSHOT_FILE.exists()
