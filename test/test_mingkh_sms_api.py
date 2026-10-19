"""No live application, credentials or portal calls in these router checks."""
from pathlib import Path
from datetime import date
import pytest
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.testclient import TestClient
from routers import mingkh_sms as api
from services.mingkh.sms import jobs, operations, rez_data, sys_data


@pytest.fixture
def client(monkeypatch):
    app = FastAPI()
    @app.middleware('http')
    async def identity(request, call_next):
        who = request.headers.get('X-Test-Account', 'reader')
        request.state.user = None if who == 'guest' else {
            'id': 2 if who == 'other' else 1, 'username': who,
            'modules': [] if who == 'denied' else ['mingkh'],
        }
        return await call_next(request)
    monkeypatch.setattr(api, 'is_account_manager', lambda u: u and u['username'] == 'manager')
    monkeypatch.setattr(operations, 'options', lambda: {'date': '2026-10-14'})
    app.mount('/static', StaticFiles(directory=Path(__file__).resolve().parents[1] / 'static'), name='static')
    app.include_router(api.router)
    return TestClient(app)


@pytest.mark.parametrize('path', ['', '/summary', '/resonance', '/system', '/solutions', '/api/options', '/api/rez/found', '/api/sys/state', '/api/jobs/'+'a'*32])
def test_all_pages_and_reads_require_module(client, path):
    for who, code in [('guest',401),('denied',403)]:
        assert client.get('/mingkh/sms'+path, headers={'X-Test-Account':who}).status_code == code


@pytest.mark.parametrize('path', ['build', 'rez/found', 'import', 'sys/run', 'rm/measures'])
def test_posts_require_module(client, path):
    assert client.post('/mingkh/sms/api/'+path, json={}, headers={'X-Test-Account':'denied'}).status_code == 403


def test_reader_cannot_import(client):
    assert client.post('/mingkh/sms/api/import', content=b'anything').status_code == 403


def test_manager_can_import_raw_zip(client, monkeypatch):
    from services.mingkh.sms import imports
    seen=[]
    monkeypatch.setattr(imports, 'import_archive', lambda payload: seen.append(payload) or {'ok':True,'stats':{}})
    r=client.post('/mingkh/sms/api/import', content=b'ZIP payload', headers={'X-Test-Account':'manager','Content-Type':'application/zip'})
    assert r.status_code==200 and r.json()['ok']
    assert seen==[b'ZIP payload']


@pytest.mark.parametrize('section', ['', '/summary', '/resonance', '/system', '/solutions'])
def test_real_templates_render_and_admin_controls_hidden(client, section):
    r = client.get('/mingkh/sms'+section)
    assert r.status_code == 200
    assert 'Дашборд по жалобам МинЖКХ' in r.text
    assert 'id="smsImport"' not in r.text
    r = client.get('/mingkh/sms'+section, headers={'X-Test-Account':'manager'})
    assert 'id="smsImport"' in r.text


def test_embed_omits_shell_and_release_popup(client):
    r = client.get('/mingkh/sms/summary?embed=1')
    assert r.status_code == 200
    assert 'release-notes.js' not in r.text
    assert 'sms-hero' not in r.text


def test_start_validated_and_uses_session_owner(client, monkeypatch):
    seen=[]
    monkeypatch.setattr(jobs, 'submit', lambda *args: seen.append(args) or {'job_id':'a'*32,'status':'queued'})
    r=client.post('/mingkh/sms/api/rm/measures', json={'from':'2026-10-01','to':'2026-10-14','owner':'local:999'}, headers={'Idempotency-Key':'a'*32})
    assert r.status_code == 202
    assert seen[0] == ('local:1','rm/measures',{'from':'2026-10-01','to':'2026-10-14'},'a'*32)
    assert 'no-store' in r.headers['cache-control']


def test_invalid_request_cannot_enqueue(client, monkeypatch):
    monkeypatch.setattr(jobs, 'submit', lambda *a: pytest.fail('invalid request enqueued'))
    for data in ({'from':'bad','to':'2026-10-14'},{'from':'2026-10-15','to':'2026-10-14'},{}):
        assert client.post('/mingkh/sms/api/rm/measures', json=data).status_code == 422
    assert client.post('/mingkh/sms/api/rm/measures', content='{"date": NaN}').status_code == 422
    assert client.post('/mingkh/sms/api/build', content=b'x'*(256*1024+1)).status_code == 413


def test_job_lookup_scoped_to_account(client, monkeypatch):
    monkeypatch.setattr(jobs, 'status', lambda owner, ident: {'result':{'text':'Private'}} if owner=='local:1' else None)
    assert client.get('/mingkh/sms/api/jobs/'+'a'*32).json()['result']['text']=='Private'
    assert client.get('/mingkh/sms/api/jobs/'+'a'*32, headers={'X-Test-Account':'other'}).status_code==404


def test_found_validation_and_state(client, monkeypatch):
    monkeypatch.setattr(rez_data, 'found_upsert', lambda f,t,c: {'added':len(c)})
    monkeypatch.setattr(sys_data, 'registries', lambda: [date(2026,10,14)])
    assert client.get('/mingkh/sms/api/sys/state').json()=={'registries':['2026-10-14']}
    assert client.post('/mingkh/sms/api/rez/found', json={'from':'2026-10-01','to':'2026-10-14','cases':[]}).json()=={'added':0}
    assert client.post('/mingkh/sms/api/rez/found', json={'from':'bad','to':'2026-10-14','cases':[]}).status_code==422


def test_saved_registry_dates_serialize(client, monkeypatch, tmp_path):
    monkeypatch.setattr(sys_data, 'REG_DIR', tmp_path)
    (tmp_path/'2026-10-14.xlsx').touch()
    (tmp_path/'2026-10-07.xlsx').touch()
    assert client.get('/mingkh/sms/api/sys/state').json() == {'registries':['2026-10-07','2026-10-14']}


def test_imported_baseline_has_a_usable_default_date(monkeypatch, tmp_path):
    from services.mingkh.sms import context
    monkeypatch.setattr(sys_data, 'REG_DIR', tmp_path)
    monkeypatch.setattr(context, 'today', lambda: date(2026,10,19))
    monkeypatch.setattr(operations.svod, 'omsu_list', lambda: [])
    (tmp_path/'2026-10-14.xlsx').touch()
    assert operations.options()['system_date'] == '2026-10-19'
    assert operations.options()['date'] == '2026-10-14'
    (tmp_path/'2026-10-07.xlsx').touch()
    assert operations.options()['system_date'] == '2026-10-14'


@pytest.mark.parametrize('path,query,allowed', [
    ('/mingkh/sms/summary','embed=1',True),('/mingkh/sms/resonance','embed=1',True),
    ('/mingkh/sms/system','embed=1',True),('/mingkh/sms/solutions','embed=1',True),
    ('/mingkh/sms/summary','',False),('/mingkh/sms/summary','embed=0',False),
    ('/users','embed=1',False),('/mingkh/sms/api/options','embed=1',False),
])
def test_actual_app_frame_policy_is_narrow(path, query, allowed):
    import ast
    import asyncio
    from starlette.requests import Request
    from starlette.responses import Response
    tree=ast.parse((Path(__file__).resolve().parents[1]/'app.py').read_text(encoding='utf-8-sig'))
    node=next(n for n in tree.body if isinstance(n,ast.AsyncFunctionDef) and n.name=='security_headers')
    node.decorator_list=[]
    ns={'Request':Request}
    exec(compile(ast.Module(body=[node],type_ignores=[]),'security_headers','exec'),ns)
    request=Request({'type':'http','path':path,'query_string':query.encode(),'headers':[], 'scheme':'http','server':('localhost',8000)})
    async def next_response(_): return Response('test')
    response=asyncio.run(ns['security_headers'](request,next_response))
    assert response.headers['X-Frame-Options']==('SAMEORIGIN' if allowed else 'DENY')
    assert 'frame-ancestors '+("'self'" if allowed else "'none'") in response.headers['Content-Security-Policy']
