import json
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
