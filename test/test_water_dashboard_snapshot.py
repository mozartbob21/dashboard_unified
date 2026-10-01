import json
import pytest
from unittest.mock import patch
from services.water_dashboard import builder


def test_all_sources_keep_own_dates_and_do_not_relabel_old_values(tmp_path):
    path=tmp_path/'snapshot.json'
    with patch.object(builder,'SNAPSHOT_FILE',path):
        one=builder.build_snapshot({'tasks':{'tables':[{'headers':['ОМСУ','Просроченные задачи'], 'rows':[['Тестовый округ','12']]}]}, 'valves':{'text':'Внесено задвижек\nЕще 0\n118\nС корректным адресом\n99'}})
        assert one['kpis']['tasks_total']==12
        assert one['kpis']['sys_vs'] is None
        assert one['sources']['valves']['widgets'][0]['value']=='118'
        # Simulate an earlier successful source timestamp before partial failure.
        one['sources']['valves']['updated_at']='2025-01-01T12:00:00'
        path.write_text(json.dumps(one))
        two=builder.build_snapshot({'tasks':{'tables':[{'headers':['ОМСУ','Просроченные задачи'], 'rows':[['Тестовый округ','0']]}]},'valves':{'error':'Нет доступа'}})
        assert two['kpis']['tasks_total']==0
        assert two['sources']['valves']['updated_at']=='2025-01-01T12:00:00'
        assert not two['sources']['valves']['ok']
        assert two['sources']['valves']['widgets'][0]['value']=='118'
        assert len(two['sources'])==8


def test_legacy_example_values_never_become_live_data(tmp_path):
    path=tmp_path/'snapshot.json';path.write_text(json.dumps({'table':[{'name':'Старые данные','tasks':864}],'kpi_cards':{'valves':1185}}))
    with patch.object(builder,'SNAPSHOT_FILE',path):
        result=builder.build_snapshot({})
    assert result['table']==[] and result['kpis']['tasks_total'] is None
    assert 'kpi_cards' not in result
    assert all(not s['ok'] for s in result['sources'].values())


def test_missing_value_is_not_zero_and_zero_attendance_is_counted():
    table=builder.build_table({'meetings':{'tables':[{'headers':['ОМСУ','Явка'], 'rows':[['Первый','0%'],['Второй','100%'],['Третий','—']]}]}})
    assert builder.derive_kpis(table)['att_avg']==50
    assert next(r for r in table if r['name']=='Третий')['att'] is None


def test_current_meeting_source_header_is_recognized():
    table = builder.build_table({'meetings': {'tables': [{
        'headers': ['ОМСУ', 'Присутствовали на перекличках, %', 'Отсутствовали на перекличках, %', 'Число перекличек'],
        'rows': [['Первый округ', '60', '40', '20']],
    }]}})
    assert table[0]['att'] == 60


@pytest.mark.parametrize('header', ['ОМСУ', 'Муниципальный округ'])
def test_municipality_can_follow_row_number_and_quantity_can_be_zero(header):
    table = builder.build_table({'tasks': {'tables': [{
        'headers': ['№', header, 'Количество'],
        'rows': [['1', 'Первый', '0'], ['2', 'Второй', '15'],
                 ['3', 'Итого', '15'], ['4']],
    }]}})
    assert {row['name']: row['tasks'] for row in table} == {'Первый': 0, 'Второй': 15}
    assert builder.derive_kpis(table)['tasks_total'] == 15


@pytest.mark.parametrize('error_text', [
    'Ошибка подключения\n500',
    'Ошибка загрузки данных\nКод\n503',
    'Подтвердите, что вы не робот\nКод проверки\n12345',
    'Доступ запрещён\n403',
    'Войдите в аккаунт\nКод подтверждения\n123456',
    'Error loading data\n500',
])
def test_error_screen_numbers_do_not_replace_latest_values_or_date(tmp_path, error_text):
    path = tmp_path / 'snapshot.json'
    with patch.object(builder, 'SNAPSHOT_FILE', path):
        first = builder.build_snapshot({'valves': {'text': 'Внесено задвижек\n118'}})
        first['sources']['valves']['updated_at'] = '2025-01-01T12:00:00'
        path.write_text(json.dumps(first), encoding='utf-8')
        failed = builder.build_snapshot({'valves': {'text': error_text}})
    source = failed['sources']['valves']
    assert builder.extract_widgets(error_text) == []
    assert source['ok'] is False
    assert source['updated_at'] == '2025-01-01T12:00:00'
    assert source['widgets'] == [{'label': 'Внесено задвижек', 'value': '118'}]
    assert not any(failed['sources_updated'].values())


def test_zero_value_and_error_count_metric_are_not_mistaken_for_error_screen():
    assert builder.extract_widgets('Количество ошибок в адресах\n0') == [
        {'label': 'Количество ошибок в адресах', 'value': '0'},
    ]


@pytest.mark.parametrize('more', ['Еще', 'Ещё'])
def test_nvos_accepts_both_widget_menu_spellings(more):
    text = '\n'.join([
        'Доля (%)', f'{more} 0', '80,00', 'собираемости',
        '% от НВВ', f'{more} 0', '5,76',
        'Сумма НВВ', f'{more} 0', '24 491M',
        'План отборов проб', f'{more} 0', '500', '(год)',
        'Факт отборов проб', f'{more} 0', '400', '(год)',
        'План отборов проб', f'{more} 0', '20', '(неделя)',
        'Факт отборов проб', f'{more} 0', '0', '(неделя)',
    ])
    assert builder.parse_nvos_kpis(text) == {
        'sbor': '80,00', 'nvv_pct': '5,76', 'sum_nvv': '24 491M',
        'plan_year': '500', 'fact_year': '400', 'plan_week': '20', 'fact_week': '0',
    }
