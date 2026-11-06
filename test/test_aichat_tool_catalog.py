"""Fixed tool dispatch, exact parameters and server-side authorization."""
from datetime import date
import json
from unittest.mock import Mock

import pytest

from services.aichat import tool_catalog as tools


USER = {'id': 7, 'username': 'synthetic', 'is_active': True,
        'modules': ['mingkh', 'edds', 'cds', 'edo']}


@pytest.fixture(autouse=True)
def no_manager_database(monkeypatch):
    from services.auth import accounts
    monkeypatch.setattr(accounts, 'is_account_manager', lambda user: bool(user.get('manager')))
    monkeypatch.setattr(tools.report_periods, 'today', lambda: date(2026, 11, 6))


def test_catalog_is_rights_filtered_and_describes_only_real_operations():
    catalog = {item['id']: item for item in tools.catalog(USER)}
    assert set(catalog) == {'refresh:edds', 'refresh:edo', 'refresh:cds',
                            'check:edds', 'check:edo', 'check:cds',
                            'analytics:mingkh-map', 'analytics:edds-water', 'query:mingkh-complaints'}
    assert tools.catalog({**USER, 'is_active': False}) == []
    assert tools.catalog({**USER, 'modules': []}) == []
    assert tools.catalog(None) == []
    assert 'refresh:mingkh-water-map' in {item['id'] for item in tools.catalog({**USER, 'manager': True})}
    assert catalog['refresh:cds']['parameters']['required'] == ['start', 'end']
    assert catalog['refresh:edo']['parameters']['properties'] == {}
    assert catalog['analytics:mingkh-map']['parameters']['properties']['service']['enum'] == ['all', 'hvs', 'gvs', 'vo']
    assert all(item['parameters']['additionalProperties'] is False for item in catalog.values())
    assert all(item['requires_confirmation'] == (item['operation'] in ('refresh', 'check')) for item in catalog.values())
    serialized = json.dumps(catalog, ensure_ascii=False)
    assert 'data/' not in serialized and 'python' not in serialized and 'password' not in serialized


def test_returned_catalog_cannot_mutate_internal_allowlist():
    item = tools.spec('refresh:edo')
    item['runner_module'] = 'shell'
    item['parameters']['properties']['command'] = {'type': 'string'}
    assert tools.spec('refresh:edo')['runner_module'] == 'edo'
    assert tools.validate(USER, 'refresh:edo', {}) == {}
    with pytest.raises(tools.InvalidParameters):
        tools.validate(USER, 'refresh:edo', {'command': 'anything'})


@pytest.mark.parametrize('ident,params', [
    ('../../shell', {}), ('refresh:camera_prescriptions', {}),
    ('refresh:edo', {'municipality': 'Чехов'}),
    ('refresh:edo', {'command': 'do something'}),
    ('refresh:cds', {'start': '2026-10-01'}),
    ('refresh:cds', {'start': '2026-11-07', 'end': '2026-11-07'}),
    ('refresh:cds', {'start': '2026-11-06', 'end': '2026-11-05'}),
    ('refresh:cds', {'start': True, 'end': '2026-11-05'}),
    ('refresh:cds', {'start': '2025-01-01', 'end': '2026-11-05'}),
    ('analytics:mingkh-map', {'start': '2026-01-01', 'end': '2026-11-05', 'user_id': 99}),
    ('analytics:mingkh-map', {'start': '2026-01-01', 'end': '2026-11-05', 'service': 'execute'}),
    ('analytics:mingkh-map', {'start': 0, 'end': '2026-11-05'}),
    ('analytics:mingkh-map', {'start': '2026-01-01', 'end': '2026-11-05', 'municipality': ['Чехов']}),
    ('analytics:mingkh-map', []),
])
def test_unknown_tools_fields_and_invalid_types_rejected_before_dispatch(ident, params):
    with pytest.raises(tools.InvalidParameters):
        tools.validate(USER, ident, params)


@pytest.mark.parametrize('ident,params', [
    ('refresh:edo', {}),
    ('refresh:mingkh-water-map', {}),
    ('analytics:mingkh-map', {'start': '2026-01-01', 'end': '2026-11-05'}),
])
def test_unauthorized_and_disabled_accounts_cannot_execute(ident, params, monkeypatch):
    invoke = Mock(side_effect=AssertionError('Denied invocation reached a runner'))
    monkeypatch.setattr(tools.block_execution, 'execute', invoke)
    for denied in ({**USER, 'modules': []}, {**USER, 'is_active': False}):
        with pytest.raises(tools.AccessDenied):
            tools.execute(denied, ident, params, confirmed=True)
    invoke.assert_not_called()


def test_inventory_covers_every_platform_block_without_claiming_unsupported_actions():
    from core.roles import ALL_MODULE_IDS
    inventory = {item['module_id']: item for item in tools.inventory({**USER, 'modules': list(ALL_MODULE_IDS)})}
    assert set(inventory) == set(ALL_MODULE_IDS)
    assert inventory['overdue']['tools'] == ['refresh:overdue', 'check:overdue', 'query:overdue']
    assert inventory['water_rm']['status'] == 'saved_read_only'
    assert 'меняет статусы' in inventory['water_rm']['limitations']
    assert inventory['tools']['status'] == 'not_connected'
    assert all(item['limitations'] for item in inventory.values())
    assert all(item['implementation_entry'] and item['source_kind'] and item['result_format'] for item in inventory.values())
    assert all('ui_filters' in item and 'agent_parameters' in item for item in inventory.values())
    assert {item['module_id'] for item in tools.inventory({**USER, 'modules': ['overdue']})} == {'overdue'}
    assert tools.inventory({**USER, 'is_active': False}) == []


def test_live_overdue_catalog_rejects_historical_dates_instead_of_silently_dropping_them():
    user = {**USER, 'modules': ['overdue']}
    with pytest.raises(tools.InvalidParameters):
        tools.validate(user, 'query:overdue', {'municipality': 'Чехов', 'start': '2026-02-01', 'end': '2026-02-28'})
    descriptor = tools.spec('query:overdue')
    assert descriptor['requires_confirmation'] is False
    assert descriptor['parameters']['properties'].keys() == {'municipality'}


def test_ecur_catalog_matches_current_only_adapter_and_inventory(monkeypatch):
    from services.aichat import live_ecur_tool
    user = {**USER, 'modules': ['ecur']}
    descriptor = tools.spec('query:ecur')
    assert descriptor['operation'] == 'query' and descriptor['requires_confirmation'] is False
    assert descriptor['required_modules'] == ['ecur']
    assert descriptor['parameters']['required'] == []
    assert descriptor['parameters']['properties'].keys() == {'municipality'}
    assert descriptor['limits'] == {'historical_period': False, 'max_rows': 10000, 'max_pages': 21}
    assert descriptor['result_schema']['statuses'] == ['ok', 'partial', 'missing']
    assert 'status_counts' in descriptor['result_schema']['extra_fields']
    assert {item['id'] for item in tools.catalog(user)} == {'query:ecur'}
    assert tools.validate(user, 'query:ecur', {}) == {'municipality': ''}
    inventory = tools.inventory(user)
    assert len(inventory) == 1 and inventory[0]['status'] == 'tools_available'
    assert inventory[0]['tools'] == ['query:ecur']
    assert 'текущего оперативного отчёта' in inventory[0]['limitations']
    invoke = Mock(return_value={'status': 'ok', 'metrics': [{'id': 'complaints', 'value': 3}]})
    monkeypatch.setattr(live_ecur_tool, 'execute', invoke)
    progress = Mock()
    assert tools.execute(user, 'query:ecur', {'municipality': ' Чехов '}, on_progress=progress)['status'] == 'ok'
    invoke.assert_called_once_with(user, {'municipality': 'Чехов'}, on_progress=progress)


@pytest.mark.parametrize('params', [
    {'start': '2026-01-01', 'end': '2026-01-31'}, {'status': 'Решено'},
    {'municipality': '<script>'}, {'municipality': ['Чехов']},
    {'municipality': ' ' * 161}, {'user_id': 1}, {'url': 'https://outside.invalid'},
])
def test_ecur_catalog_rejects_unsupported_filters_before_provider_execution(params, monkeypatch):
    from services.aichat import live_ecur_tool
    invoke = Mock()
    monkeypatch.setattr(live_ecur_tool, 'execute', invoke)
    with pytest.raises(tools.InvalidParameters):
        tools.execute({**USER, 'modules': ['ecur']}, 'query:ecur', params)
    invoke.assert_not_called()


def test_ecur_catalog_denies_ungranted_or_disabled_account_before_credentials(monkeypatch):
    from services.aichat import live_ecur_tool
    credentials = Mock(side_effect=AssertionError('Denied credential read'))
    monkeypatch.setattr(live_ecur_tool, 'credentials', credentials)
    for user in (USER, {**USER, 'modules': ['ecur'], 'is_active': False}):
        with pytest.raises(tools.AccessDenied):
            tools.execute(user, 'query:ecur', {})
    credentials.assert_not_called()


def test_live_rating_and_collective_descriptors_delegate_only_to_fixed_providers(monkeypatch):
    from services.aichat import live_rating_tool, live_collective_tool
    rating_user = {**USER, 'modules': ['omsu-rating']}
    rating_params = {'indicator': 7, 'period': 2, 'municipality': ''}
    rating = Mock(return_value={'status': 'ok', 'rating_rows': []})
    monkeypatch.setattr(live_rating_tool, 'execute', rating)
    assert tools.execute(rating_user, 'query:omsu-rating', rating_params)['status'] == 'ok'
    rating.assert_called_once_with(rating_user, 'query:omsu-rating', rating_params, on_progress=None)
    for params in ({'indicator': True, 'period': 2}, {'indicator': 7, 'period': 0}, {'indicator': 5}):
        with pytest.raises(tools.InvalidParameters):
            tools.validate(rating_user, 'query:omsu-rating', params)
    collective = Mock(return_value={'status': 'ok', 'metrics': []})
    monkeypatch.setattr(live_collective_tool, 'execute', collective)
    params = {'start': '2026-10-01', 'end': '2026-10-31', 'municipality': '', 'status': ''}
    user = {**USER, 'modules': ['collective']}
    assert tools.execute(user, 'query:collective', params)['status'] == 'ok'
    collective.assert_called_once_with(user, params, on_progress=None)
    with pytest.raises(tools.AccessDenied):
        tools.validate({**USER, 'modules': ['mingkh']}, 'query:collective', params)


def test_map_refresh_manager_requirement_not_confused_with_read_permission():
    assert tools.permitted(USER, 'analytics:mingkh-map')
    assert not tools.permitted(USER, 'refresh:mingkh-water-map')
    with pytest.raises(tools.AccessDenied):
        tools.validate(USER, 'refresh:mingkh-water-map', {})
    assert tools.validate({**USER, 'manager': True}, 'refresh:mingkh-water-map', {}) == {}


def test_refresh_requires_explicit_server_confirmation_then_reuses_original_runner(monkeypatch):
    invoke = Mock(return_value={'status': 'success', 'fresh_result': True})
    monkeypatch.setattr(tools.block_execution, 'execute', invoke)
    params = {'start': '2026-11-01', 'end': '2026-11-05'}
    for value in (False, None, 1, 'true'):
        with pytest.raises(tools.InvalidParameters):
            tools.execute(USER, 'refresh:cds', params, confirmed=value)
    invoke.assert_not_called()
    runner, progress = Mock(), Mock()
    result = tools.execute(USER, 'refresh:cds', params, confirmed=True,
                           run_id='server_generated_identifier', run_subprocess=runner, on_progress=progress)
    assert result['fresh_result'] is True
    invoke.assert_called_once_with(USER, 'cds', params, run_id='server_generated_identifier',
                                    run_subprocess=runner, on_progress=progress)


def test_analytics_uses_provider_validation_and_fixed_dispatch(monkeypatch):
    from services.aichat import analytical_tools
    raw = {'start': '2026-01-01', 'end': '2026-12-31'}
    clean = {**raw, 'service': 'all', 'municipality': ''}
    check = Mock(return_value=clean)
    invoke = Mock(return_value={'status': 'partial', 'total': 17})
    monkeypatch.setattr(analytical_tools, 'validate', check)
    monkeypatch.setattr(analytical_tools, 'execute', invoke)
    progress = Mock()
    result = tools.execute(USER, 'analytics:mingkh-map', raw, on_progress=progress)
    assert result == {'status': 'partial', 'total': 17}
    check.assert_called_once_with(USER, 'analytics:mingkh-map', raw)
    invoke.assert_called_once_with(USER, 'analytics:mingkh-map', clean, on_progress=progress)


def test_analytics_denied_before_reading_provider(monkeypatch):
    from services.aichat import analytical_tools
    check, invoke = Mock(), Mock()
    monkeypatch.setattr(analytical_tools, 'validate', check)
    monkeypatch.setattr(analytical_tools, 'execute', invoke)
    with pytest.raises(tools.AccessDenied):
        tools.execute({**USER, 'modules': ['collective']}, 'analytics:edds-water',
                      {'start': '2026-01-01', 'end': '2026-12-31'})
    check.assert_not_called()
    invoke.assert_not_called()
