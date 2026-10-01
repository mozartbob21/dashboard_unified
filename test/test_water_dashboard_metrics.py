import pytest

from services.water_dashboard.metrics import primary_available, source_metrics


def widget(label, value, caption=''):
    return {'label': label, 'value': value, 'caption': caption}


def values(metrics):
    return {metric['id']: metric['value'] for metric in metrics}


def table_data(*headers):
    return {'tables': [{'headers': list(headers), 'rows': []}]}


def test_valves_uses_named_indicator_not_population_or_plan_as_primary():
    data = {'widgets': [
        widget('Население', '8 900 000'), widget('Должно быть внесено', '100'),
        widget('Внесено с корректным адресом', '80'), widget('Внесено всего на 2027 год', '120'),
    ]}
    metrics = source_metrics('valves', data, [])
    assert values(metrics) == {'inserted': 120, 'plan': 100, 'correct_address': 80, 'completion_pct': 120}
    assert metrics[0]['unit'] == 'шт.' and primary_available(metrics)


def test_plan_only_cannot_replace_missing_valves_primary():
    metrics = source_metrics('valves', {'widgets': [widget('Должно быть внесено', '100')]}, [])
    assert metrics[0]['value'] is None and values(metrics)['plan'] == 100
    assert not primary_available(metrics)


def test_conflicting_indicator_values_are_not_silently_chosen():
    data = {'widgets': [widget('Внесено всего на 2026 год', '100'),
                        widget('Внесено всего на 2026 год', '200')]}
    assert not primary_available(source_metrics('valves', data, []))


def test_matching_duplicate_indicators_and_nonbreaking_spaces_are_safe():
    data = {'widgets': [widget('Внесено всего на 2026 год', '1\u00a0000'),
                        widget('  Внесено  всего на 2026 год ', '1000')]}
    assert source_metrics('valves', data, [])[0]['value'] == 1000


def test_edo_signature_percentage_is_not_electronic_document_percentage():
    data = {'widgets': [
        widget('Доля (%) ЭД в общем кол-ве документов', '90'),
        widget('Кол-во должностных лиц с правом подписи', '20'),
        widget('Наличие ЭЦП у подписывающих, кол-во', '12'),
        widget('Доля (%) должностных лиц, имеющих право подписи и ЭЦП', '60 %'),
    ]}
    metrics = source_metrics('edo_rso', data, [])
    assert values(metrics) == {
        'signers_ecp_pct': 60, 'signers_total': 20, 'signers_ecp': 12,
        'electronic_documents_pct': 90,
    }
    assert metrics[0]['unit'] == '%'


def test_edo_cannot_claim_primary_from_document_percentage_only():
    metrics = source_metrics('edo_rso', {'widgets': [
        widget('Доля (%) ЭД в общем кол-ве документов', '90'),
    ]}, [])
    assert metrics[0]['value'] is None and not primary_available(metrics)


def test_tasks_sums_verified_municipal_values_and_preserves_zero():
    metrics = source_metrics('tasks', table_data('ОМСУ', 'Кол-во задач'),
                             [{'name': 'Первый', 'tasks': 0}, {'name': 'Второй', 'tasks': 12}])
    assert metrics == [{'id': 'overdue_tasks', 'label': 'Просроченные задачи', 'value': 12, 'unit': 'шт.'}]
    zero = source_metrics('tasks', table_data('ОМСУ', 'Кол-во задач'), [{'name': 'Первый', 'tasks': 0}])
    assert primary_available(zero) and zero[0]['value'] == 0


def test_population_column_cannot_become_task_count():
    data = table_data('ОМСУ', 'Количество жителей')
    metrics = source_metrics('tasks', data, [{'name': 'Первый', 'tasks': 800000}])
    assert not primary_available(metrics)


@pytest.mark.parametrize('sid,suffix', [('sys_vs', 'VS'), ('sys_kr', 'KR')])
def test_system_and_resonance_counts_are_distinct_and_not_dynamics(sid, suffix):
    data = table_data('ОМСУ', 'Системных', 'Динамика системных за неделю', 'Резонансных')
    metrics = source_metrics(sid, data, [
        {'name': 'Первый', 'sys' + suffix: 12, 'res' + suffix: 5},
        {'name': 'Второй', 'sys' + suffix: 0, 'res' + suffix: 10},
    ])
    assert values(metrics) == {'system_addresses': 12, 'resonant_addresses': 15}
    dynamics = table_data('ОМСУ', 'Динамика системных за неделю')
    assert not primary_available(source_metrics(sid, dynamics, [{'sys' + suffix: 99}]))


def test_attendance_uses_percentage_column_including_zero_without_truncation():
    data = table_data('ОМСУ', 'Присутствовали на перекличках, %', 'Число перекличек')
    data['widgets'] = [widget('Проведено совещаний', '8')]
    metrics = source_metrics('meetings', data, [{'att': 0}, {'att': 66.6}, {'att': None}])
    assert values(metrics) == {'attendance_pct': 33.3, 'meetings': 8}
    assert metrics[0]['label'] == 'Средняя явка по ОМСУ'


def test_attendance_count_without_percentage_is_not_presented_as_percent():
    data = table_data('ОМСУ', 'Присутствовали на перекличках', 'Число перекличек')
    assert not primary_available(source_metrics('meetings', data, [{'att': 50}]))


def test_nvos_distinguishes_collection_from_subscribers_and_year_from_week():
    data = {'widgets': [
        widget('Доля (%) абонентов', '85', 'на коэфф 0.5'),
        widget('Доля (%)', '28', 'на коэфф 2'),
        widget('Доля (%)', '81,25', 'собираемости платы за негативное воздействие'),
        widget('% от НВВ', '8,5'),
        widget('Факт отборов проб', '10', '(неделя)'),
        widget('Факт отборов проб', '100', '(год)'),
        widget('План отборов проб', '15', '(неделя)'),
        widget('План отборов проб', '200', '(год)'),
        widget('Сумма НВВ', '2,5M'),
    ]}
    result = values(source_metrics('nvos', data, []))
    assert result['collection_pct'] == 81.25 and result['nvv_pct'] == 8.5
    assert result['samples_fact_year'] == 100 and result['samples_plan_year'] == 200
    assert result['samples_fact_week'] == 10 and result['samples_plan_week'] == 15
    assert result['nvv_total'] == 2500000


def test_nvos_accepts_explicit_parsed_fields_without_global_text_guessing():
    data = {'text': 'Произвольное число\n1000000', 'nvos': {
        'sbor': '75,5', 'nvv_pct': '8,34', 'fact_week': '0', 'plan_week': '12',
        'sum_nvv': '20\u00a0000M', 'pay_str': '1,0 / 2,0 млн ₽',
    }}
    metrics = source_metrics('nvos', data, [])
    assert primary_available(metrics)
    result = values(metrics)
    assert result['collection_pct'] == 75.5 and result['samples_fact_week'] == 0
    assert result['paid_accrued'] == '1,0 / 2,0 млн ₽'
    assert result['nvv_total'] == 20000000000


def test_flush_ignores_unverified_indicators_and_control_row_counts():
    data = {'widgets': [widget('Всего строк', '123'), widget('Выполнено промывок', '99')],
            'tables': [{'headers': ['ОМСУ', 'Sys'], 'rows': [['Первый', 1]]}]}
    metrics = source_metrics('flush', data, [])
    assert metrics == [{'id': 'completed', 'label': 'Выполнено промывок', 'value': None, 'unit': 'шт.'}]
    assert not primary_available(metrics)


def test_flush_recovers_when_verified_numeric_indicator_becomes_available():
    data = {'widgets': [widget('Кол-во выполненных промывок от общего кол-ва', '1\u00a0250')]}
    metrics = source_metrics('flush', data, [])
    assert metrics[0]['value'] == 1250 and primary_available(metrics)


def test_flush_percentage_or_chart_series_is_not_a_completed_count():
    for value in ('65%', '65 %', '1250\n2000', [1250, 2000]):
        metrics = source_metrics('flush', {'widgets': [
            widget('Кол-во выполненных промывок от общего кол-ва', value),
        ]}, [])
        assert metrics[0]['value'] is None and not primary_available(metrics)


def test_flush_conflicting_indicator_values_do_not_supply_primary():
    data = {'widgets': [
        widget('Кол-во выполненных промывок от общего кол-ва', '1250'),
        widget('Кол-во выполненных промывок от общего кол-ва', '2000'),
    ]}
    metrics = source_metrics('flush', data, [])
    assert metrics[0]['value'] is None and not primary_available(metrics)


@pytest.mark.parametrize('sid', ['valves', 'flush', 'tasks', 'sys_vs', 'edo_rso', 'nvos', 'meetings', 'sys_kr'])
def test_arbitrary_table_widgets_never_supply_main_metric(sid):
    data = {'widgets': [widget('Первый округ', '5000000'), widget('Население', '1000')],
            'text': 'Собираемость\n50\nВнесено\n123'}
    metrics = source_metrics(sid, data, [])
    assert metrics and metrics[0]['value'] is None and not primary_available(metrics)


@pytest.mark.parametrize('invalid', ['—', 'null', '500 error', True, float('nan'), float('inf'), -1, '12%'])
def test_invalid_count_does_not_make_primary_available(invalid):
    metrics = source_metrics('valves', {'widgets': [widget('Внесено всего на 2026 год', invalid)]}, [])
    assert not primary_available(metrics)


def test_unknown_source_has_no_default_metric():
    assert source_metrics('unknown', {'widgets': [widget('Всего', 100)]}, []) == []
    assert not primary_available([])
