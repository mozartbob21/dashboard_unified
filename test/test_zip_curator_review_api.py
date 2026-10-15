"""Exercise the real curator handlers against isolated server files."""
import ast
from pathlib import Path

import pytest
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
from fastapi.testclient import TestClient
from services.zip_curator import core, dictionary


@pytest.fixture
def client(monkeypatch, tmp_path):
    for key, value in {
        'DATA': tmp_path, 'STATE_FILE': tmp_path / 'state.json',
        'PUBLISHED_FILE': tmp_path / 'published.json', 'INPUT_DIR': tmp_path / 'input',
        'FOLDER_A_DIR': tmp_path / 'input/folder_a',
        'FOLDER_A_MANIFEST_FILE': tmp_path / 'manifest.json',
        'MUNICIPALITY_OVERRIDES_FILE': tmp_path / 'municipalities.json',
    }.items():
        monkeypatch.setattr(core, key, value)
    monkeypatch.setattr(dictionary, 'CUSTOM_FILE', tmp_path / 'dictionary.json')
    dictionary._merged.cache_clear()
    core.save_state({'pending':[{'rso':'Тестовое РСО', 'okrug':'', 'items':[
        {'name':'Труба', 'cat':'Трубы', 'grp':'Труба электросварная', 'qty':None,
         'unit':None, 'unitUnknown':True}]}], 'clean':{}})
    tree = ast.parse((Path(__file__).parents[1] / 'app.py').read_text(encoding='utf-8-sig'))
    selected = ast.Module(body=[node for node in tree.body if isinstance(node, ast.AsyncFunctionDef)
                              and node.name in {'zc_approve','zc_edit','zc_publish'}], type_ignores=[])
    api = FastAPI()
    exec(compile(selected, 'app.py', 'exec'), {'app':api,'zc':core,'JSONResponse':JSONResponse,'Request':Request})
    with TestClient(api) as result:
        yield result
    dictionary._merged.cache_clear()


def test_missing_omsu_and_units_have_actionable_422_and_no_publication(client):
    response = client.post('/zip_curator/api/approve',json={'idx':0})
    assert response.status_code == 422
    body = response.json()
    assert {issue['field'] for issue in body['issues']} >= {'okrug','unit','qty'}
    assert len(body['state']['pending']) == 1
    assert not core.PUBLISHED_FILE.exists()


def test_edit_then_approve_persists_running_metres_zero_and_guards_stale_tabs(client):
    core.save_municipality_override('Тестовое РСО','Королёв')
    original = core.load_state()['pending'][0]
    edit = client.post('/zip_curator/api/edit',json={'pi':0,'ii':0,'cat':'Трубы','grp':'Труба электросварная',
                                                 'unit':'погонный метр','qty':0,'revision':original['revision']})
    assert edit.status_code == 200
    record = edit.json()['state']['pending'][0]
    assert record['items'][0]['unit'] == 'пог. м'
    assert record['items'][0]['qty'] == 0
    assert record['approval_issues'] == []
    assert client.post('/zip_curator/api/approve',json={'idx':0,'revision':original['revision']}).status_code == 409
    assert client.post('/zip_curator/api/approve',json={'idx':0,'revision':record['revision']}).status_code == 200
    assert client.post('/zip_curator/api/approve',json={'idx':0,'revision':record['revision']}).status_code == 409
    assert len(core.load_state()['clean']) == 1


@pytest.mark.parametrize('idx',[None,'0',True,1.5,{},[]])
def test_bad_indices_are_client_errors(client,idx):
    assert client.post('/zip_curator/api/approve',json={'idx':idx}).status_code == 400
    assert client.post('/zip_curator/api/edit',json={'pi':idx,'ii':0}).status_code == 400


def test_invalid_unit_is_not_saved(client):
    before = core.STATE_FILE.read_bytes()
    response = client.post('/zip_curator/api/edit',json={'pi':0,'ii':0,'cat':'Трубы','grp':'Труба электросварная','unit':'ведро'})
    assert response.status_code == 400
    assert core.STATE_FILE.read_bytes() == before


def test_direct_publication_cannot_bypass_missing_municipality(client):
    rows = [['РСО','Округ','Наименование','Категория','Группа','Количество','Ед.','Водоподготовка','Дата'],
            ['Тестовое РСО','','Труба','Трубы','Труба электросварная',90,'пог. м','нет','']]
    response = client.post('/zip_curator/api/publish',json={'rows':rows})
    assert response.status_code == 422
    assert not core.PUBLISHED_FILE.exists()
