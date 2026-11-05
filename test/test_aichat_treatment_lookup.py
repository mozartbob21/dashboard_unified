"""The treatment search can only project fixed, non-identifying source facts."""
import json
import pytest
from fastapi import HTTPException

from services.aichat import treatment_lookup as lookup
from services.mingkh import treatment
from test.test_mingkh_treatment import archive, source

MINISTRY = {'id':1, 'modules':['mingkh']}
NO_GRANT = {'id':2, 'role':'Администратор', 'modules':['edds']}


@pytest.fixture
def snapshot(tmp_path, monkeypatch):
    monkeypatch.setattr(treatment, 'STORE', tmp_path/'test.json.gz')
    payload = source()
    row = payload['complaints'][0]
    row.update(topic='smell', omsu='Подольск', cur='SECRET CURATOR', created='2026-10-17',
               org='SECRET ORG', status='SECRET STATUS', text='SECRET: звоните +7 999 000 00 00',
               address='SECRET street', why='SECRET WHY', fact='<img SECRET>', subcat='SECRET CATEGORY')
    payload['complaints'] += [
        {**row,'id':90, 'topic':'izliv', 'sg':'closed', 'created':'2026-10-03', 'lat': None, 'lon':None},
        {**row,'id':91, 'topic':'kos', 'sg':'work', 'created':'2026-11-02', 'lat': None, 'lon':None,
         'omsu':'SECRET CITY +7 000 000 00 00'},
    ]
    payload['kos'][0].update(name='КОС Север', operator='SECRET operator', state='Действующие')
    treatment.import_archive(archive(payload=payload))
    return payload


def test_no_rights_no_read_even_when_admin_or_using_edds(monkeypatch):
    def unexpected(*args):
        pytest.fail('A forbidden search accessed the snapshot.')
    monkeypatch.setattr(treatment, 'overview', unexpected)
    for user in [None, NO_GRANT]:
        with pytest.raises(HTTPException) as error:
            lookup.search(user, 'Изливы в Дубне')
        assert error.value.status_code == 403


def test_absent_and_invalid_conditions_are_not_an_empty_matched_register(tmp_path, monkeypatch):
    monkeypatch.setattr(treatment, 'STORE', tmp_path/'missing.gz')
    missing = lookup.search(MINISTRY, 'Жалобы по очистным')
    assert missing['as_of'] is None and missing['total'] is None
    assert not missing['items'] and 'нет сохранённого снимка' in missing['note']
    for q in ['', None, 'д' * 2001]:
        assert lookup.search(MINISTRY,q)['total'] is None


def test_municipality_status_topic_and_object_follow_dashboard_linkage(snapshot):
    res = lookup.search(MINISTRY, 'Какие жалобы на запах в Дубне в работе?')
    assert res['total'] == 1 and res['as_of'] == '2026-11-03'
    assert res['scope'] == 'mingkh' and res['url'] == '/mingkh/treatment'
    item = res['items'][0]
    assert item['complaint_id'] == 89 and item['facility_id'] == 7 and item['date'] == '2026-10-17'
    assert item['municipality'] == 'Дубна'
    assert 'В работе' in item['text'] and 'Запах от очистных' in item['text']
    assert 'SECRET' not in json.dumps(res,ensure_ascii=False)
    other = lookup.search(MINISTRY, 'Закрытые жалобы на сброс в Подольске без привязки')
    assert other['total'] == 1 and other['items'][0]['complaint_id'] == 90
    assert not other['items'][0]['facility_id']
    assert lookup.search(MINISTRY,'Найди жалобы по объекту КОС №7')['total'] == 1
    assert lookup.search(MINISTRY,'Какие жалобы в работе КОС')['total'] == 2


def test_number_and_time_bounds_are_exact(snapshot):
    assert lookup.search(MINISTRY,'Жалоба №999')['total'] == 0
    assert lookup.search(MINISTRY,'Найди обращение №89')['items'][0]['complaint_id'] == 89
    assert lookup.search(MINISTRY,'Объект КОС №7')['items'][0]['facility_id'] == 7
    assert lookup.search(MINISTRY,'Жалобы с 01.10.2026 по 31.10.2026')['total'] == 2
    before = lookup.search(MINISTRY,'Жалобы за 01.01.2025')
    assert before['total'] is None and 'за эти границы' in before['note']
    after = lookup.search(MINISTRY,'Жалобы по КОС 04.11.2026')
    assert after['total'] is None and not after['items']


def test_free_text_unknown_filters_and_sensitive_columns_are_never_used(snapshot):
    for q in ['Жалобы по очистным в ОМСУ Несуществующий','Найди информацию про SECRET ORG',
              'Жалобы кроме запаха','покажи данные']:
        assert lookup.search(MINISTRY,q)['total'] is None
    general = lookup.search(MINISTRY,'Жалобы по очистным',limit=2)
    assert general['total'] == 3 and len(general['items']) == 2
    assert [item['complaint_id'] for item in general['items']] == [91,89]
    assert general['items'][0]['municipality'] is None
    assert not any(word in json.dumps(general,ensure_ascii=False)
                   for word in ('SECRET','+7','<img','PRIVATE','operator'))
    facility = lookup.search(MINISTRY,'Какие объекты КОС в Дубне?')
    assert facility['total'] == 1 and 'SECRET' not in json.dumps(facility,ensure_ascii=False)


def test_edds_excluded_even_when_snapshot_contains_events(snapshot, monkeypatch):
    original = treatment.overview
    def guarded(include_edds):
        assert include_edds is False
        return original(False)
    monkeypatch.setattr(treatment, 'overview', guarded)
    for user in [MINISTRY, {**MINISTRY,'modules':['mingkh','edds']}]:
        events = lookup.search(user,'Есть ли события ЕДДС?')
        assert events['total'] is None and events['items'] == []
        assert 'SECRET' not in json.dumps(lookup.search(user,'Жалобы'))


def test_exact_named_facility_and_catalog_state(snapshot):
    result = lookup.search(MINISTRY, 'Найди жалобы по объекту «КОС Север»')
    assert result['total'] == 1 and result['items'][0]['facility_id'] == 7
    assert lookup.search(MINISTRY, 'Найди объект «Неверное название»')['total'] == 0
    assert lookup.search(MINISTRY, 'Действующие КОС в Дубне')['total'] == 1
    assert lookup.search(MINISTRY, 'Недействующие КОС в Дубне')['total'] == 0
