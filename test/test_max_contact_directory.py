"""The importer and polling worker share one directory. All contacts below are synthetic."""
import io
import json
from concurrent.futures import ThreadPoolExecutor

import pytest
from openpyxl import Workbook

from services.max_bot.config import Settings
from services.max_bot.stock import StockError, StockService
from services.zip_curator import contacts as web
from services.zip_curator import paths


@pytest.fixture
def directory(tmp_path, monkeypatch):
    shared = tmp_path / "mounted"
    monkeypatch.setenv("ZIP_CONTACTS_DIR", str(shared))
    return StockService(tmp_path / "never-published.json", Settings.from_env().contacts_path)


def workbook(rows):
    book = Workbook()
    for row in rows:
        book.active.append(row)
    out = io.BytesIO()
    book.save(out)
    return out.getvalue()


def import_contacts(rows):
    return web.import_contacts("Справочник.xlsx",workbook(rows))


def test_direct_search_uses_imported_workbook_without_stocks_or_notes(directory):
    import_contacts([
        ["Справочник сотрудников"],
        ["Служебный телефон", "Фамилия, имя, отчество", "Наименование ресурсоснабжающей организации", "Муниципальное образование", "Должность", "Примечание", "Почта"],
        ["8 000 111-11-11", "Иванов Иван Иванович", "АО «Водоканал»", "Чехов", "Главный инженер", "PRIVATE-NOTE", "example@invalid.test"],
        ["8 000 222-22-22", "Иванов Пётр", "АО Водоканал", "Химки", "Диспетчер", "", ""],
    ])
    assert not directory.published_path.exists()
    assert len(web.load_contacts()["items"])==2
    for phrase in ["контакт Иванова в Чехове","как связаться с водоканалом в Чехове","кто ответственный в Чехове","найди главного инженера Чехова"]:
        reply=directory.search_contacts(phrase)
        assert reply["total"]==1, phrase
        contact=reply["items"][0]
        assert contact == {"organization":"АО «Водоканал»","municipality":"Чехов","person":"Иванов Иван Иванович",
                           "position":"Главный инженер","phone":"8 000 111-11-11"}
    assert 'invalid.test' not in str(reply) and 'PRIVATE-NOTE' not in str(reply)
    assert directory.search_contacts('нужен телефон водоканала в Химках')['items'][0]['person']=='Иванов Пётр'
    assert directory.search_contacts('в Щёлково')['total']==0


def test_directory_pagination_and_refresh_on_import_and_delete(directory):
    rows=[["ФИО ответственного", "ОМСУ", "Название РСО", "Номер телефона"]] + [
        [f"Инженер {i:02}", "Истра","МУП Теплосеть",f"1234{i:02}"] for i in range(11)]
    import_contacts(rows)
    found=directory.search_contacts('контакты Теплосети в Истре',limit=5)
    next_page=directory.search_contacts('Теплосети в Истре',limit=5, offset=5)
    assert found['total']==next_page['total']==11
    assert len(found['items'])==len(next_page['items'])==5
    assert {c['phone'] for c in found['items']}.isdisjoint(c['phone'] for c in next_page['items'])
    removed=web.load_contacts()['items'][0]
    assert web.delete_contact(removed['id'])
    assert directory.search_contacts("Теплосеть")['total']==10
    rows[-1][3]="123499"
    import_contacts(rows[:1]+rows[-1:])
    assert directory.search_contacts("Инженер 10")['items'][0]['phone']=="123499"
    directory.contacts_path.unlink()
    with pytest.raises(StockError,match="ещё не загружен"):
        directory.search_contacts("Истра")


@pytest.mark.parametrize("query",["","как связаться","телефон РСО","контакты АО","контакты главного инженера","*","в"])
def test_broad_queries_do_not_reveal_directory(directory, query):
    with pytest.raises(StockError) as exc:
        directory.search_contacts(query)
    assert exc.value.code=='INVALID_QUERY'


def test_legacy_phone_correction_and_scoped_match_use_dashboard_rules(directory):
    directory.contacts_path.parent.mkdir()
    directory.contacts_path.write_text(json.dumps({"items":[
        {"rso":"Акционерное общество «Водоканал»", "municipality":"Чехов", "person":"Инженер Тестовый", "phone":"Директор Тестовый"},
        {"rso":"ООО «Водоканал»", "municipality":"Чехов", "person":"Другая Организация", "phone":"123456"},
        {"rso":"АО Водоканал","municipality":"Химки","person":"Другой Город","phone":"123456"},
    ]}))
    assert directory.search_contacts("Инженер Чехов")['items'][0]['phone'] is None
    assert directory.search_contacts("Тестовый Чехов")['total']==2
    directory.published_path.write_text(json.dumps({"rows":[["РСО","Округ","Наименование","Количество","Ед."],["АО Водоканал","Чехов","Муфта",1,"шт"]]}))
    stock=directory.search("Муфта")['queries'][0]['items'][0]
    scoped=directory.contacts(stock['id'],stock['municipality_id'])
    assert len(scoped['contacts'])==2
    assert all(entry['phone'] is None for entry in scoped['contacts'])
    assert "Другая Организация" not in str(scoped) and "Другой Город" not in str(scoped)


def test_shared_location_defaults_and_optional_override(tmp_path, monkeypatch):
    monkeypatch.delenv("ZIP_CONTACTS_DIR",raising=False)
    assert Settings.from_env({}).contacts_path==web.CONTACTS_FILE==paths.CONTACTS_FILE
    assert Settings.from_env({"ZIP_CONTACTS_DIR":str(tmp_path)}).contacts_path==tmp_path/"contacts.json"
    assert Settings.from_env({}).published_path==paths.PUBLISHED_FILE
    with pytest.raises(ValueError):
        Settings.from_env({"ZIP_CONTACTS_DIR":"https://example.invalid/"})
    with pytest.raises(ValueError):
        Settings.from_env({"ZIP_CONTACTS_DIR":"static/contacts"})


def test_two_uploads_keep_both_contacts_and_corrupt_file_is_not_replaced(directory):
    header="РСО;ОМСУ;ФИО;Телефон\n"
    def upload(i):
        return web.import_contacts('a.csv',(header+f"РСО {i};Химки;Тестов {i};123456").encode())["added"]
    with ThreadPoolExecutor(max_workers=2) as pool:
        assert list(pool.map(upload,range(2)))==[1,1]
    assert directory.search_contacts('Химки')['total']==2
    directory.contacts_path.write_bytes(b"{broken")
    with pytest.raises(StockError) as exc:
        directory.search_contacts('Химки')
    assert exc.value.code=='CONTACTS_UNAVAILABLE'
    with pytest.raises(ValueError):
        upload(3)
    assert directory.contacts_path.read_bytes()==b"{broken"


def test_text_or_address_in_phone_field_is_not_sent_as_work_number(directory):
    directory.contacts_path.parent.mkdir()
    directory.contacts_path.write_text(json.dumps({"items":[
        {"rso":"Теплосеть","person":"Первый", "phone":"ул. Наборная 123"},
        {"rso":"Теплосеть","person":"Второй", "phone":"+7 (000) 222-11-11, доб. 9"},
    ]}))
    items = {row["person"]: row["phone"] for row in directory.search_contacts("Теплосети")["items"]}
    assert items == {"Первый":None, "Второй":"+7 (000) 222-11-11, доб. 9"}
