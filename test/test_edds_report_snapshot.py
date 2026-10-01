import json
from datetime import date

from services.edds import report_snapshot as snapshot

HEADER = ['id_cds_claim', 'name_mr', 'ispolnitel', 'd_create', 'd_doklad', 'type_otkl', 'obj_vs', 'obj_vo', 'd_close', 'cnt_people', 'text_message']
START, END = date(2026, 10, 1), date(2026, 10, 2)


def row(id='1', city='Власиха', water='1', closed='', people='12', created='01.10.2026 08:00'):
    return [id, city, 'РСО', created, '01.10.2026 09:00', 'Аварийная заявка', water, '0', closed, people, 'PRIVATE TEXT secret@example.ru']


def test_projection_matches_water_filter_and_does_not_save_claim_text():
    result = snapshot.project([HEADER, row(), row('2', closed='02.10.26 09:00', people='-1'), row('3', water='0')], START, END)
    assert result['source_rows'] == 3
    assert result['totals'] == {'incidents': 2, 'closed': 1, 'active': 1, 'people_known_sum': 12, 'people_missing': 1}
    assert result['municipalities']['Власиха'] == result['totals']
    assert result['period'] == {'from': '2026-10-01', 'to': '2026-10-02'}
    assert 'PRIVATE' not in json.dumps(result) and 'secret@' not in json.dumps(result)


def test_broken_rows_and_unknown_city_are_not_assigned_to_a_municipality():
    result = snapshot.project([['title'], HEADER, row(city=''), row('piece'), row('3', created='31.02.26 08:00')], START, END)
    assert result['broken_rows'] == 2
    assert result['unknown_municipality'] == 1
    assert result['totals']['incidents'] == 1
    assert result['municipalities'] == {}


def test_unknown_coverage_is_not_a_negative_or_nonfinite_population():
    result = snapshot.project([HEADER, row(people='NaN'), row('2', people=''), row('3', people='0')], START, END)
    assert result['totals']['people_missing'] == 2
    assert result['totals']['people_known_sum'] == 0


def test_invalid_export_preserves_previous_snapshot(tmp_path, monkeypatch):
    target = tmp_path / 'saved.json'
    monkeypatch.setattr(snapshot, 'TARGET', target)
    assert snapshot.persist([HEADER, row()], START, END)
    before = target.read_bytes()
    assert not snapshot.persist([HEADER, row('broken')], START, END)
    assert not snapshot.persist([['other', 'schema']], START, END)
    assert target.read_bytes() == before
    assert list(tmp_path.iterdir()) == [target]


def test_empty_valid_water_selection_is_a_known_zero():
    result = snapshot.project([HEADER, row(water='0')], START, END)
    assert result['totals']['incidents'] == 0
    assert result['broken_rows'] == 0
