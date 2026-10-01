"""No portal calls or application database fixtures: snapshots use tmp_path."""
import asyncio
import json
from datetime import date

import pytest
from starlette.requests import Request

from services.aichat import remaining_sources as reports


HEADER = ['Номер', 'Адрес', 'Район', 'Категория ЕЦУР', 'Подкатегория ЕЦУР',
          'ЕЦУР факт', 'Исполнитель', 'Текст', 'Дата создания', 'Срок', 'Статус']


def complaint(municipality='Власиха', deadline='2026-10-02', status='В работе'):
    return [123, 'PRIVATE_ADDRESS', municipality, 'Вода', 'SUBCATEGORY', 'FACT',
            'PRIVATE_PERSON', 'PRIVATE_MESSAGE', '2026-09-01', deadline, status]


def water_result(rows=(), total=None):
    return {'rez': {'close': list(rows), 'rw': []}, 'sys': {'close': [], 'ext': [], 'rw': []},
            'total_tasks': len(rows) if total is None else total}


@pytest.fixture
def registry(tmp_path):
    (tmp_path / 'municipality_registry.json').write_text(json.dumps({'municipalities': {
        'Власиха': {'aliases': ['Власиха (ЗАТО)']}, 'Химки': {'aliases': ['г.о. Химки']},
    }}, ensure_ascii=False))
    return tmp_path


def test_ecur_projection_has_no_personal_text_credentials_or_metadata(tmp_path):
    assert reports.persist_ecur_snapshot([HEADER, complaint()], {'email': 'SECRET_EMAIL', 'password': 'SECRET_PASSWORD'}, data_dir=tmp_path)
    raw = (tmp_path / 'ecur/report-summary.json').read_text()
    assert not any(secret in raw for secret in ('PRIVATE_', 'SECRET_', 'SUBCATEGORY', 'FACT'))
    source = reports.reports_for_module('ecur', tmp_path, 'Власиха')[0]
    assert source['metrics'][0]['value'] == 1
    assert source['collected_at'] and source['data_date']
    assert source['status_counts'] == {'В работе': 1}
    assert reports.names_for_module('ecur', tmp_path) == {'Власиха'}


def test_ecur_zero_is_valid_but_absent_municipality_is_unknown(tmp_path):
    assert reports.persist_ecur_snapshot([HEADER], data_dir=tmp_path)
    assert reports.reports_for_module('ecur', tmp_path)[0]['metrics'][0]['value'] == 0
    municipal = reports.reports_for_module('ecur', tmp_path, 'Власиха')[0]
    assert all(metric['value'] is None for metric in municipal['metrics'])
    assert 'Отдельных строк' in municipal['warning']


def test_ecur_deadline_buckets_are_bound_to_collection_date():
    # Friday: today, Sunday and later in the month are separate buckets.
    result = reports.project_ecur([HEADER, complaint(deadline='2026-10-01'),
                                  complaint(deadline='2026-10-02'), complaint(deadline='2026-10-04'),
                                  complaint(deadline='2026-10-12'), complaint(deadline='2026-11-01'),
                                  complaint(deadline='unknown')], today=date(2026, 10, 2))
    assert result['as_of'] == '2026-10-02'
    assert {k: result['totals'][k] for k in reports.DEADLINE_LABELS} == {
        'total': 6, 'overdue': 1, 'today': 1, 'week': 1, 'month': 1, 'later': 1, 'unknown_deadline': 1}


def test_ecur_alias_scope_never_returns_other_city_totals(tmp_path):
    assert reports.persist_ecur_snapshot([HEADER, complaint('Химки'), complaint('г.о. Химки'), complaint('Власиха')], data_dir=tmp_path)
    source = reports.reports_for_module('ecur', tmp_path, 'Химки')[0]
    assert source['metrics'][0]['value'] == 2
    assert source['status_counts'] == {'В работе': 2}


def test_invalid_ecur_does_not_replace_last_success(tmp_path):
    reports.persist_ecur_snapshot([HEADER, complaint()], data_dir=tmp_path)
    before = (tmp_path / 'ecur/report-summary.json').read_bytes()
    assert not reports.persist_ecur_snapshot([HEADER, ['broken']], data_dir=tmp_path)
    assert (tmp_path / 'ecur/report-summary.json').read_bytes() == before


def test_water_whitelist_rejects_geographic_guessing(registry):
    payload = water_result([
        {'id': 1, 'project': 'г.о. Химки', 'subject': 'PRIVATE_SUBJECT', 'why': 'PRIVATE_REASON'},
        {'id': 2, 'project': 'Водоканал Химки', 'subject': 'Власиха', 'token': 'SECRET_TOKEN'},
    ], total=3)
    payload.update(password='SECRET_PASSWORD', collected_at='1900-01-01')
    saved = reports.save_water_check(payload, data_dir=registry)
    raw = json.dumps(saved, ensure_ascii=False)
    assert not any(secret in raw for secret in ('PRIVATE_', 'SECRET_', '1900-01-01', 'Водоканал'))
    assert saved['rez']['close'] == [{'id': 1, 'municipality': 'Химки'}, {'id': 2, 'municipality': ''}]
    assert saved['unclassified_tasks'] == 1 and saved['unknown_municipality_rows'] == 1
    assert reports.names_for_module('water_rm', registry) == {'Химки'}
    municipal = reports.reports_for_module('water_rm', registry, 'Химки')[0]
    assert municipal['metrics'][0]['value'] == 1
    assert 'рекомендации' in municipal['warning']
    assert all(m['value'] is None for m in reports.reports_for_module('water_rm', registry, 'Власиха')[0]['metrics'])


def test_water_valid_zero_and_missing_are_different(registry):
    assert reports.reports_for_module('water_rm', registry)[0]['status'] == 'missing'
    reports.save_water_check(water_result(), data_dir=registry)
    source = reports.reports_for_module('water_rm', registry)[0]
    assert source['status'] == 'saved'
    assert all(m['value'] == 0 for m in source['metrics'])
    assert all(m['value'] is None for m in reports.reports_for_module('water_rm', registry, 'Химки')[0]['metrics'])


def test_invalid_water_preserves_existing_snapshot(registry):
    reports.save_water_check(water_result([{'id': 1, 'project': 'Власиха'}]), data_dir=registry)
    path = registry / 'water_rm/last_check.json'
    before = path.read_bytes()
    for bad in ({}, water_result([{'id': True}]), water_result([{'id': 1}, {'id': 1}]), water_result([{'id': 1}], total=0)):
        with pytest.raises(ValueError):
            reports.save_water_check(bad, data_dir=registry)
        assert path.read_bytes() == before


def test_reader_reloads_after_new_success(registry):
    reports.save_water_check(water_result(), data_dir=registry)
    assert reports.reports_for_module('water_rm', registry)[0]['metrics'][0]['value'] == 0
    reports.save_water_check(water_result([{'id': 8, 'project': 'Власиха'}]), data_dir=registry)
    assert reports.reports_for_module('water_rm', registry)[0]['metrics'][0]['value'] == 1


def test_corrupt_snapshots_are_missing_and_not_zero(tmp_path):
    for module, relative in [('ecur', 'ecur/report-summary.json'), ('water_rm', 'water_rm/last_check.json')]:
        path = tmp_path / relative; path.parent.mkdir(parents=True, exist_ok=True)
        for raw in ('broken', '{}', '{"schema_version":1,"collected_at":"2026-10-01"}'):
            path.write_text(raw)
            source = reports.reports_for_module(module, tmp_path)[0]
            assert source['status'] == 'missing'
            assert not source['metrics']


def test_ecur_success_hooks_do_not_persist_credentials(monkeypatch):
    from services.ecur import client
    state = {'session': None, 'email': None, 'password': None, 'rows': None, 'meta': None}
    monkeypatch.setattr(client, 'STATE', state)
    session = object(); persisted = []
    monkeypatch.setattr(client, 'dobrodel_login', lambda *a: (session, None))
    monkeypatch.setattr(client, 'fetch_report', lambda *a: ([HEADER, complaint()], None))
    monkeypatch.setattr(reports, 'persist_ecur_snapshot', lambda rows, meta: persisted.append((rows, dict(meta))) or True)
    assert client.authenticate_user('qa@example.invalid', 'TEST_PASSWORD')[0]
    assert client.refresh_data()[0]
    assert len(persisted) == 2
    assert 'TEST_PASSWORD' not in json.dumps(persisted)
    monkeypatch.setattr(client, 'fetch_report', lambda *a: (None, 'Failed'))
    assert not client.refresh_data()[0]
    assert len(persisted) == 2


@pytest.mark.parametrize('payload', [[], {'rows': []}, {'content': []}])
def test_ecur_confirmed_empty_response_is_a_successful_zero(payload):
    from services.ecur import client
    class Response:
        status_code = 200
        text = '[]'
        url = client.REPORT_API
        def json(self):
            return payload
    class Session:
        def get(self, *args, **kwargs):
            return Response()
    rows, error = client.fetch_report(Session())
    assert error is None and rows == [HEADER]


@pytest.mark.parametrize('payload', [{}, {'error': 'no access'}, {'error': 'no access', 'rows': []}, {'rows': 'wrong'}, [None]])
def test_ecur_unknown_schema_is_not_a_zero(payload):
    from services.ecur import client
    class Response:
        status_code = 200
        text = '{}'
        url = client.REPORT_API
        def json(self):
            return payload
    class Session:
        def get(self, *args, **kwargs):
            return Response()
    rows, error = client.fetch_report(Session())
    assert rows is None and error


def test_history_endpoint_validates_and_keeps_get_contract(tmp_path, monkeypatch):
    from services.water_rm import proxy
    monkeypatch.setattr(proxy, 'DATA_DIR', tmp_path / 'water_rm')
    monkeypatch.setattr(proxy, 'HISTORY_FILE', tmp_path / 'water_rm/last_check.json')
    def request(payload):
        async def receive():
            return {'type': 'http.request', 'body': json.dumps(payload).encode(), 'more_body': False}
        return Request({'type': 'http', 'method': 'POST', 'path': '/water-rm/history', 'headers': []}, receive)
    result = asyncio.run(proxy.save_history(request(water_result())))
    assert result['ok'] and result['collected_at']
    saved = json.loads(asyncio.run(proxy.get_history()).body)
    assert saved['rez'] == {'close': [], 'rw': []} and saved['total_tasks'] == 0
    with pytest.raises(proxy.HTTPException) as error:
        asyncio.run(proxy.save_history(request({'token': 'SECRET'})))
    assert error.value.status_code == 422
    assert json.loads(asyncio.run(proxy.get_history()).body) == saved
