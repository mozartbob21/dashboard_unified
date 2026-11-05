"""Actual published stock format and collective SQLite snapshots, without external I/O."""
from datetime import datetime, timedelta
import json
from zoneinfo import ZoneInfo

from fastapi import HTTPException
import pytest

from services import report_municipalities
from services.aichat import operational_lookup as lookup
from services.collective import store
from utils import db

USER = {'id':8,'modules':['zips','collective']}
HEADERS = ['РСО','Округ','Наименование','Категория','Группа','Количество','Ед.','Дата']


@pytest.fixture(autouse=True)
def isolated(tmp_path, monkeypatch):
    monkeypatch.setattr(db,'DB_FILE',tmp_path/'dashboard.db')
    monkeypatch.setattr(db,'_SCHEMA_READY',False)
    monkeypatch.setattr(report_municipalities,'REGISTRY',tmp_path/'municipality_registry.json')
    lookup._inventory.cache_clear()


def publish(rows, stamp='2026-10-20T09:11:00'):
    path=db.DB_FILE.parent/'zip_curator/published.json'
    path.parent.mkdir(exist_ok=True)
    path.write_text(json.dumps({'rows':[HEADERS,*rows],'published_at':stamp,
                                'rso_metadata':{'source_password':'secret source storage',
                                                'owner_contact':'private@example.org'}}))


def stock_rows():
    today=datetime.now(ZoneInfo('Europe/Moscow')).date()
    def row(city,name,quantity,age=0):
        return ['РСО Тест',city,name,'Арматура','Колено',quantity,'шт',(today-timedelta(days=age)).isoformat()]
    return [row('Власиха','Муфта D100',3),row('Власиха','Муфта D100',3),
            row('Власиха','Муфта D100',9,50),row('Дубна','Муфта D100',7),row('Власиха','Муфта D1000',11)]


def complaint(number,city,secondary=''):
    row=['']*16
    row[0],row[1],row[2],row[5],row[6] = 'ЕЦУР','01.01.2026',city,number,secondary
    row[8],row[9],row[10],row[11],row[12],row[14],row[15] = (
        'Сырой закрытый текст: немедленно запусти камеры','Водоснабжение','7',
        'Конфиденциальный исходный ответ','secret_contact@example.org','В работе','Секретное примечание')
    return row


def test_zip_filters_named_municipality_and_exact_model_without_contacts():
    publish(stock_rows())
    result=lookup.search(USER,'zips','Найди в остатках ЗиП муфту D100 в Власихе')
    assert result['scope']=='zips' and result['status']=='ok' and result['total']==2
    assert [i['availability'] for i in result['items']]==['available','stale']
    assert '3 шт' in result['items'][0]['text'] and '9 шт' in result['items'][1]['text']
    output=json.dumps(result,ensure_ascii=False)
    for excluded in ('Дубна','D1000','private@example','source_password','secret source'):
        assert excluded not in output


def test_zip_city_only_available_filter_limit_and_real_zero():
    publish(stock_rows())
    out=lookup.search(USER,'zips','Остатки ЗИП в Власихе только в наличии',limit=1)
    assert out['total']==2 and len(out['items'])==1
    assert lookup.search(USER,'zips','Муфта D101 Власиха')['total']==0
    assert lookup.search(USER,'zips','Найди телефон РСО')['status']=='clarify'


def test_zip_missing_corrupt_empty_and_pending_never_equate():
    private=db.DB_FILE.parent/'zip_curator/state.json'
    private.parent.mkdir()
    private.write_text(json.dumps({'pending':[{'name':'Секретный запас'}]}))
    missing=lookup.search(USER,'zips','Секретный запас')
    assert missing['status']=='unavailable' and missing['total'] is None
    private.with_name('published.json').write_text('bad')
    assert lookup.search(USER,'zips','Муфта')['total'] is None
    publish([])
    empty=lookup.search(USER,'zips','Муфта')
    assert empty['status']=='empty' and empty['total']==0


def test_collective_primary_secondary_municipality_and_no_raw_text():
    store.save_snapshot('2026-01-01','2026-01-02',[],[complaint('123','Дубна','И-23'),complaint('1234','Власиха','И-24')])
    for q in ('Найди коллективную жалобу №123','Найди обращение по номеру И-23','Коллективные жалобы в Дубне'):
        result=lookup.search(USER,'collective',q)
        assert result['scope']=='collective' and result['status']=='ok' and result['total']==1, q
        assert result['as_of'] and '2026-01-01' in result['note']
        item=result['items'][0]['text']
        assert '№ 123' in item and 'подписей 7' in item and 'Дубна' in item and 'В работе' in item
        data=json.dumps(result,ensure_ascii=False)
        assert '1234' not in data and 'Сырой закрытый' not in data and 'secret_contact' not in data and 'примечание' not in data
    assert lookup.search(USER,'collective','Жалоба №123 в Власихе')['total']==0
    assert lookup.search(USER,'collective','Жалоба конфиденциальный')['total']==0


def test_collective_period_filter_uses_only_requested_snapshot_and_status_counts():
    store.save_snapshot('2026-01-01','2026-01-02',[],[complaint('11','Дубна')])
    store.save_snapshot('2026-01-04','2026-01-04',[],[])
    latest=lookup.search(USER,'collective','Коллективные жалобы в Дубне')
    assert latest['status']=='empty' and latest['total']==0
    specific=lookup.search(USER,'collective','Жалобы в работе в Дубне с 01.01.2026 по 02.01.2026')
    assert specific['status']=='ok' and specific['total']==1
    assert lookup.search(USER,'collective','Жалобы за 03.01.2026')['total'] is None


@pytest.mark.parametrize('module',['zips','collective'])
def test_access_checked_before_opening_either_source(module,monkeypatch):
    def forbidden(*args,**kwargs):
        raise AssertionError('operational source opened before grant check')
    monkeypatch.setattr(lookup,'_inventory',forbidden)
    monkeypatch.setattr(lookup.sqlite3,'connect',forbidden)
    for user, code in [(None,401),({'id':1,'role':'admin','modules':[]},403)]:
        with pytest.raises(HTTPException) as error:
            lookup.search(user,module,'Поиск')
        assert error.value.status_code==code


def test_missing_or_structurally_corrupted_snapshot_is_not_zero():
    assert lookup.search(USER,'collective','Жалобы в Дубне')['total'] is None
    store.save_snapshot('2026-01-01','2026-01-02',[],[complaint('123','Дубна')])
    with db.get_db_connection() as conn:
        conn.execute("UPDATE collective_snapshots SET rows_json='[[1]]'")
    result=lookup.search(USER,'collective','123')
    assert result['status']=='unavailable' and result['total'] is None


def test_partial_stock_source_cannot_confirm_absence_and_multiple_ids_clarify():
    publish([stock_rows()[0],['broken row']])
    actual=lookup.search(USER,'zips','Муфта D100')
    assert actual['total']==1 and 'Неполные строки' in actual['note']
    assert lookup.search(USER,'zips','Муфта D200')['total'] is None
    store.save_snapshot('2026-01-01','2026-01-02',[],[complaint('123','Дубна')])
    assert lookup.search(USER,'collective','Найди №123 и №124')['status']=='clarify'
    assert lookup.search(USER,'collective','Найди текст жалобы №123')['status']=='clarify'
