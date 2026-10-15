"""Real browser interaction with synthetic API data; no live accounts or network."""
import copy
import json
import os
from pathlib import Path

import pytest

ROOT = Path(__file__).parents[1]
playwright = pytest.importorskip('playwright.sync_api')


@pytest.fixture(scope='module')
def browser():
    with playwright.sync_playwright() as runtime:
        try:
            instance = runtime.chromium.launch(headless=True)
        except playwright.Error as error:
            pytest.skip(f'Local Chromium unavailable: {error}')
        yield instance
        instance.close()


@pytest.fixture
def ui(browser):
    context = browser.new_context(viewport={'width':1280,'height':900}, locale='ru-RU')
    page = context.new_page()
    state = {'pending':[{'rso':'Тестовый водоканал', 'okrug':'Королёв','revision':'v1',
              'items':[{'name':f'Труба {i+1} ППУ', 'cat':'Трубы','grp':'Труба электросварная',
                        'qty':i,'unit':'пог. м','unitUnknown':False,'via':'match'} for i in range(606)]}],
             'clean':{'старое рсо':{'rso':'Старое РСО','okrug':'','items':[]}}}
    state['pending'][0]['items'][605].update(qty=None,unit=None,unitUnknown=True)
    release = {'id':'test-20261015','version':'15.10.2026','title':'Обновление сервиса','read_at':None,
               'entries':[{'id':'zip','module_name':'Остатки ЗиП РСО','title':'Согласование и единицы измерения',
                           'text':'Перед публикацией проверяем ОМСУ, категорию, количество и единицы. Все позиции доступны по страницам.'}]}
    seen, edits, errors = [], [], []
    options = {'release':False,'fail':False}
    page.on('pageerror', lambda error: errors.append(str(error)))
    def route(request):
        url = request.request.url.split('http://ui.test',1)[-1].split('?',1)[0]
        def reply(data,status=200):
            request.fulfill(status=status,content_type='application/json',body=json.dumps(data,ensure_ascii=False))
        if url.startswith('/static/'):
            path = ROOT / url.lstrip('/')
            if path.is_file():
                kind = 'text/javascript' if path.suffix == '.js' else 'text/css' if path.suffix == '.css' else 'image/svg+xml'
                request.fulfill(path=path,content_type=kind)
            else: request.fulfill(status=404)
        elif url == '/zip_curator': request.fulfill(path=ROOT/'templates/zip_curator_original.html',content_type='text/html')
        elif url == '/':
            request.fulfill(content_type='text/html',body='''<!doctype html><html lang="ru"><head><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/static/release-notes.css"><script src="/static/release-notes.js" defer></script></head><body><button data-release-notes-open>Что нового · Версия 3</button><input aria-label="Рабочая форма"></body></html>''')
        elif url.startswith('/api/releases/'):
            if url.endswith('/read'):
                if options['fail']: reply({'ok':False},503)
                else: seen.append(release['id']); reply({'ok':True})
            else: reply({'releases':[release] if options['release'] and (url.endswith('/history') or not seen) else []})
        elif url.endswith('/dictionary'):
            reply({'dict':{},'categories':['Трубы'],'groupsByCat':{'Трубы':['Труба электросварная']},
                   'units':['шт','м','пог. м'], 'umap':{'пог. м':'пог. м','шт':'шт','м':'м'},'keywords':[],'rso2omsu':{}})
        elif url.endswith('/municipalities'): reply({'items':[]})
        elif url.endswith('/source'): reply({'provider':'yandex','label':'Яндекс Диск','configured':True})
        elif url.endswith('/state'): reply(state)
        elif url.endswith('/edit'):
            if options['fail']: request.abort(); return
            patch = request.request.post_data_json; edits.append(patch)
            record = state['pending'][patch['pi']]
            assert patch['revision'] == record['revision']
            record['items'][patch['ii']].update({k:patch[k] for k in ['cat','grp','unit','qty']})
            record['items'][patch['ii']]['unitUnknown'] = not patch['unit']
            record['revision'] = 'v'+str(len(edits)+1)
            reply({'ok':True,'state':state})
        elif url.endswith('/approve'):
            record = state['pending'].pop(0); state['clean'][record['rso']] = record
            reply({'ok':True,'state':state})
        else: request.fulfill(status=404)
    context.route('**/*',route)
    yield page,state,options,seen,edits,errors
    context.close()


def snapshot(page,name):
    directory = os.getenv('NEURONA_UI_SCREENSHOTS')
    if directory:
        Path(directory).mkdir(parents=True,exist_ok=True)
        page.screenshot(path=str(Path(directory)/name),full_page=True)


def test_last_page_edit_save_and_approval(ui):
    page,state,options,seen,edits,errors = ui
    page.goto('http://ui.test/zip_curator')
    page.locator('#reviewToggle0').click()
    assert page.locator('#approve0').is_disabled()
    page.locator('#pv0 input[aria-label="Номер страницы"]').first.fill('13')
    page.locator('#pv0 input[aria-label="Номер страницы"]').first.press('Tab')
    row = page.locator('#pv0 tr[data-item-index="605"]')
    row.wait_for(); row.get_by_role('button',name='Исправить',exact=True).click()
    row.locator('[data-edit="unit"]').select_option('пог. м')
    row.locator('[data-edit="qty"]').fill('90,5')
    snapshot(page,'zip-review-light.png')
    row.get_by_role('button',name='Сохранить',exact=True).click()
    page.wait_for_function('!document.querySelector("#approve0").disabled')
    assert edits[0]['ii'] == 605 and edits[0]['qty'] == 90.5 and edits[0]['unit'] == 'пог. м'
    page.reload(); page.locator('#reviewToggle0').click()
    page.locator('#pv0 input[aria-label="Номер страницы"]').first.fill('13')
    page.locator('#pv0 input[aria-label="Номер страницы"]').first.press('Tab')
    assert '90,5' in page.locator('tr[data-item-index="605"]').inner_text()
    assert page.get_by_role('button',name='Указать ОМСУ',exact=True).count() == 1
    page.locator('#approve0').click()
    page.wait_for_function('document.querySelector("#pending").textContent.includes("Пусто")')
    assert len(state['pending']) == 0
    assert not errors


def test_failed_save_keeps_editor_and_page(ui):
    page,state,options,seen,edits,errors = ui
    page.goto('http://ui.test/zip_curator'); page.locator('#reviewToggle0').click()
    row = page.locator('tr[data-item-index="0"]'); row.get_by_role('button',name='Исправить',exact=True).click()
    row.locator('[data-edit="qty"]').fill('123')
    options['fail'] = True
    page.get_by_role('button',name='Следующая страница',exact=True).first.click()
    page.wait_for_function('document.querySelector(".review-save-note").textContent.includes("не сохранена")')
    assert row.locator('[data-edit="qty"]').input_value() == '123'
    assert page.locator('#pv0 input[aria-label="Номер страницы"]').first.input_value() == '1'
    options['fail'] = False
    page.get_by_role('button',name='Следующая страница',exact=True).first.click()
    page.locator('tr[data-item-index="50"]').wait_for()
    assert edits[0]['ii'] == 0
    assert not errors


def test_release_auto_show_read_receipt_manual_reopen_and_safe_text(ui):
    page,state,options,seen,edits,errors = ui
    options['release'] = True
    page.goto('http://ui.test/')
    dialog = page.locator('#releaseNotesDialog'); dialog.wait_for(state='visible',timeout=12000)
    assert not seen
    snapshot(page,'release-light.png')
    dialog.get_by_role('button',name='Понятно',exact=True).click(); dialog.wait_for(state='detached')
    assert seen == ['test-20261015']
    page.reload(); page.wait_for_timeout(6500)
    assert not page.locator('#releaseNotesDialog').count()
    page.get_by_role('button',name='Что нового · Версия 3').click(); dialog.wait_for(state='visible')
    page.set_viewport_size({'width':390,'height':844})
    page.evaluate('document.documentElement.dataset.theme="dark"')
    snapshot(page,'release-mobile-dark.png')
    assert dialog.bounding_box()['width'] <= 390
    assert dialog.locator('.release-notes-body').evaluate('(e)=>e.scrollWidth<=e.clientWidth')
    page.keyboard.press('Escape'); dialog.wait_for(state='detached')
    assert not errors


def test_release_failure_does_not_mark_read_and_form_input_defers_popup(ui):
    page,state,options,seen,edits,errors = ui
    options['release'] = True
    page.goto('http://ui.test/')
    page.get_by_role('textbox',name='Рабочая форма').fill('Не прерывать')
    page.wait_for_timeout(6500)
    assert not page.locator('#releaseNotesDialog').count()
    page.get_by_role('button',name='Что нового · Версия 3').click()
    dialog = page.locator('#releaseNotesDialog'); dialog.wait_for(state='visible')
    options['fail'] = True
    dialog.get_by_role('button',name='Закрыть',exact=True).click()
    dialog.get_by_role('button',name='Повторить',exact=True).wait_for()
    assert not seen
    options['fail'] = False
    dialog.get_by_role('button',name='Повторить',exact=True).click(); dialog.wait_for(state='detached')
    assert seen
    assert not errors
