import io
import json
from pathlib import Path
import stat
import zipfile
from urllib.parse import urlsplit

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.testclient import TestClient
import pytest

from routers import treatment as routes
from services.mingkh import treatment
from utils import db

BASE = Path(__file__).resolve().parents[1]
PATH = '/mingkh/treatment'
SECRET = 'SECRET complaint text <img src=x onerror="window.pwned=1">'
EVENT = 'SECRET incident notes'


def source():
    return {'built': '03.11.2026 12:38', 'from': '2026-01-01',
            'kos': [{'n': 7, 'omsu': 'Дубна', 'name': 'КОС', 'place': 'Дубна',
                     'lat': 55.5, 'lon': 37.5, 'operator': 'Организация', 'okey': 'operator'}],
            'complaints': [{'id': 89, 'created': '2026-11-02', 'sg': 'work', 'topic': 'kos',
                            'omsu': 'Дубна', 'cur': 'МинЖКХ', 'org': 'Организация', 'text': SECRET,
                            'address': 'PRIVATE street', 'status': 'В работе', 'lat':55.5, 'lon':37.501}],
            'edds': [{'id':'incident-1', 'omsu':'Дубна', 'start':'2026-10-30', 'text':EVENT}],
            'pay': {'operator': {'org':'Организация','recv':0,'accr':10}}, 'extra': {'shared':'PRIVATE_PATH'}}


def archive(name='снимок.html', *, payload=None, html=None, symlink=False, additional=False):
    data = source() if payload is None else payload
    if html is None:
        html = ('<script src="https://invalid.example/source.js"></script><script>window.pwned=1</script>'
                + '<script>const DATA = ' + json.dumps(data,ensure_ascii=False) + '; window.pwned=2;</script>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        member = zipfile.ZipInfo(name)
        if symlink: member.external_attr = (stat.S_IFLNK | 0o777) << 16
        z.writestr(member,html)
        if additional: z.writestr('another.html',html)
    return buf.getvalue()


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(treatment,'STORE',tmp_path/'snapshot.json.gz')
    monkeypatch.setattr(db,'DB_FILE',tmp_path/'temp.db')
    monkeypatch.setattr(db,'_SCHEMA_READY',False)
    monkeypatch.setattr(routes,'is_account_manager',lambda u:u.get('username') == 'manager')
    app = FastAPI()
    current = {'id':1,'username':'manager','modules':['mingkh','edds']}
    @app.middleware('http')
    async def identity(request, call_next):
        request.state.user = dict(current)
        return await call_next(request)
    app.include_router(routes.router)
    app.mount('/static',StaticFiles(directory=BASE/'static'),name='static')
    with TestClient(app) as http:
        yield http,current


@pytest.mark.parametrize('name', ['../outside.html','folder/../evil.html','C:\\windows\\evil.html','/evil.html','evil.py'])
def test_path_and_type_rejected(client, name):
    http, _ = client
    res = http.post(PATH+'/api/import',content=archive(name))
    assert res.status_code == 400
    assert not treatment.STORE.exists()
    assert not (treatment.STORE.parent.parent / 'outside.html').exists()


def test_malformed_archives_cannot_replace_saved_snapshot(client, monkeypatch):
    http,_ = client
    assert http.post(PATH+'/api/import',content=archive()).status_code == 200
    saved = treatment.STORE.read_bytes()
    for bad in [b'not a zip', archive(symlink=True), archive(additional=True),
                archive(html='<script>const DATA = getExternalData();</script>'),
                archive(payload={'built':'2026-11-03','from':'2026-01-01','kos':[], 'complaints':[]})]:
        assert http.post(PATH+'/api/import',content=bad).status_code == 400
        assert treatment.STORE.read_bytes() == saved
    monkeypatch.setattr(treatment,'MAX_HTML',10)
    assert http.post(PATH+'/api/import',content=archive()).status_code == 400
    assert treatment.STORE.read_bytes() == saved


def test_import_rights_history_projections_and_details(client):
    http,who = client
    assert http.get(PATH+'/api/data').status_code == 404
    assert http.post(PATH+'/api/import',content=archive()).json() == {
        'built':'03.11.2026 12:38', 'complaints':1, 'objects':1}
    assert http.post(PATH+'/api/import',content=b'wrong').status_code == 400
    with db.get_db_connection() as conn:
        journal = conn.execute('SELECT result_code,metrics,params FROM check_results ORDER BY rowid').fetchall()
    assert len(journal) == 2 and [r['result_code'] for r in journal] == ['measured','failed']
    assert json.loads(journal[0]['metrics']) == {'appeals':1,'facilities':1}
    assert json.loads(journal[0]['params']) == {'operation':'treatment'}
    assert SECRET not in str([tuple(row) for row in journal])
    who.update(username='reader',modules=['mingkh'])
    assert http.post(PATH+'/api/import',content=archive()).status_code == 403
    page = http.get(PATH)
    assert page.status_code == 200 and 'type="file"' not in page.text
    assert 'data-view="edds"' not in page.text and 'data-check-operation="treatment"' in page.text
    data = http.get(PATH+'/api/data')
    assert data.status_code == 200 and 'no-store' in data.headers['cache-control']
    assert all(s not in data.text for s in (SECRET, EVENT, 'PRIVATE', 'window.pwned','edds_built','"edds"'))
    assert data.json()['complaints'][0]['near'] == 7
    assert 0 < data.json()['complaints'][0]['distance'] < 0.5
    assert data.json()['payments'][0]['recv'] == 0
    assert http.get(PATH+'/api/complaints/89').json()['text'] == SECRET
    assert http.get(PATH+'/api/events/incident-1').status_code == 403
    who.update(modules=['mingkh','edds'])
    assert http.get(PATH+'/api/events/incident-1').json()['text'] == EVENT
    who.update(modules=['edds'])
    for path in [PATH, PATH+'/api/data', PATH+'/api/complaints/89', PATH+'/api/events/incident-1']:
        assert http.get(path).status_code == 403


def test_snapshot_settings_apply_without_executing_source(client):
    snapshot = {'inputs':{'t:r:kos':'2'}, 'overrides':{'7':[55.501,37.502]}}
    html = '<script id="snap">window.KOS_SNAPSHOT = '+json.dumps(snapshot)+';</script>'
    html += '<script>const DATA = '+json.dumps(source())+'; arbitraryExecutableCode()</script>'
    parsed = treatment.parse_archive(archive(html=html))
    assert parsed['radii']['kos'] == 2
    assert parsed['kos'][0]['lat'] == 55.501
    assert 'arbitraryExecutableCode' not in json.dumps(parsed)


@pytest.mark.parametrize('width',[390,1280])
def test_browser_snapshot_filters_detail_and_map(client, width):
    playwright = pytest.importorskip('playwright.sync_api')
    http,_ = client
    assert http.post(PATH+'/api/import', content=archive()).status_code == 200
    with playwright.sync_playwright() as pw:
        try:
            browser = pw.chromium.launch()
        except playwright.Error as exc:
            pytest.skip('Локальный Chromium недоступен: ' + str(exc).splitlines()[0])
        context = browser.new_context(viewport={'width':width,'height':880})
        def safe_route(route):
            url = urlsplit(route.request.url)
            if url.netloc != 'testserver':
                return route.abort()
            if url.path.startswith((PATH,'/static/')):
                response = http.request(route.request.method,url.path,content=route.request.post_data_buffer)
                return route.fulfill(status=response.status_code,body=response.content,content_type=response.headers.get('content-type'))
            return route.fulfill(status=200,json={})
        context.route('**/*',safe_route)
        page = context.new_page()
        page.goto('http://testserver'+PATH)
        page.wait_for_selector('#trWorkspace:not([hidden])')
        assert page.locator('#trMetrics strong').first.inner_text() == '1'
        page.locator('[data-view="appeals"]').click()
        page.get_by_role('button',name='№89').click()
        page.wait_for_function("document.querySelector('#trDetailBody').textContent.includes('SECRET complaint')")
        assert SECRET in page.locator('#trDetailBody').inner_text()
        assert not page.evaluate('Boolean(window.pwned)')
        page.locator('#trDetail button').click()
        page.locator('select[name=sg]').select_option('closed')
        assert page.locator('#trMetrics strong').first.inner_text() == '0'
        page.locator('select[name=sg]').select_option('')
        page.locator('[data-view="map"]').click()
        assert page.locator('#trCanvas.leaflet-container').is_visible()
        page.locator('[data-view="edds"]').click()
        assert 'incident-1' in page.locator('#trTable').inner_text()
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth + 1')
        browser.close()
