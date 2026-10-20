"""Synthetic mobile regression suite. No app imports, accounts or production data.
Run: python -m unittest test.test_mobile_system_browser
Browser requests are all intercepted; no external service is contacted.
Optional screenshots: NEURONA_UI_SCREENSHOTS=/path/to/output.
"""
from pathlib import Path
from urllib.parse import urlsplit
import json
import os
import mimetypes
import unittest
import jinja2

ROOT = Path(__file__).resolve().parents[1]
PAGES=['edds','home','aichat','summarizer','telegram','tools','users','ux_survey','ux_survey_admin','login','register','mingkh','mingkh-water-map','collective','water_dashboard','zip_curator_original','mingkh-sms-overview','mingkh-sms-system','mingkh-sms-solutions','mingkh-sms-resonance','mingkh-sms-summary','edo','overdue','watercontrol','water_rm','mgkh_rm','utnkr','ecur','cds','cameras','camera_prescriptions','appeals','appeal_detail','scheduler','municipality_report','prompthub','zips']
MODULES=['edo','overdue','watercontrol','utnkr','cameras','cds','mgkh_rm','ecur','appeals','municipality-report','mingkh','water-dashboard','zip_curator','tools','telegram']
ctx=dict(url_for=lambda route,**kw:'/static/'+kw.get('path','') if route=='static' else '/'+route,user_username='Тестовый пользователь',user_role='admin',user_modules=MODULES,can_manage_users=True,home_preferences={},can_manage=True,can_view_mingkh=True,can_view_collective=True,can_import=True,can_build=True,can_manage_import=True,owner_key='synthetic',module_names={x:'Длинное название тестового блока '+x for x in MODULES},initial_admin_section='users',initial_dates={'start':'2026-10-01','end':'2026-10-20'},presets=[('week','Последние семь дней'),('month','Последние тридцать дней')],config={},metrics={},result=None,items=[],results=[],filtered_results=[],grouped_results={},group_by='none',status_filter='all',table_info={},check_state={},appeals=[],stats={},item={},final_index={},safe_final_index={},run_status={},prescriptions=[],generated_items=[],report={},all_blocks=[],municipalities=['Красногорск','Богородский городской округ'],selected_blocks=[],snapshot_date='20.10.2026',map_config={'provider':'osm'},logo_url='/static/logo.svg',is_authed=False,user_email='',sms_labels={'overview':'Обзор','system':'Системные','solutions':'Решения','resonance':'Резонансные','summary':'Свод'},sms_descriptions={},sms_page='overview',embed=False)
ctx['snapshot']={'metric_schema':1,'updated_at':'2026-10-20T12:00:00','checked_at':'2026-10-20T12:00:00','sources':{k:{'metric_schema':1,'ok':True,'updated_at':'2026-10-20T12:00:00','refresh':'Ежедневно в рабочее время','metrics':[{'label':'Доля выполненных работ','value':78,'unit':'%'}],'details':{'basis':'Тестовая детализация по муниципалитетам','groups':[{'id':'worst','title':'Топ-5 отстающих','items':[{'name':'Богородский городской округ с длинным названием','value':20,'unit':'%'}]}]}} for k in ['valves','tasks','sys_vs','edo_rso','nvos','meetings','sys_kr']},'table':[{'name':'Богородский городской округ','resVS':12,'sysVS':7,'tasks':99,'sysKR':18,'resKR':3,'att':87}],'kpis':{}}

# Semantic IDs exercise the same presentation contract as production collectors.
for source_id in ('sys_vs', 'sys_kr'):
    ctx['snapshot']['sources'][source_id]['metrics'] = [
        {'id': 'system_addresses', 'label': 'Системные адреса', 'value': 17, 'unit': ''},
        {'id': 'resonant_addresses', 'label': 'Резонансные адреса', 'value': 8, 'unit': ''},
    ]
ctx['snapshot']['sources']['edo_rso']['metrics'] = [
    {'id': 'electronic_documents_pct', 'label': 'Электронные документы в общем количестве', 'value': 78, 'unit': '%'},
]

def template_environment():
    env = jinja2.Environment(loader=jinja2.FileSystemLoader(str(ROOT / 'templates')),
                             undefined=jinja2.ChainableUndefined, autoescape=True)
    env.policies['json.dumps_kwargs'] = {'default': lambda value: None}
    return env


def rendered_page(env, name):
    if name == 'edds':
        return (ROOT / 'services/edds/dashboard.html').read_text().replace('__NEURONA_EDDS_OWNER__', 'synthetic').replace('__NEURONA_MAP_CONFIG__', '{}')
    return env.get_template(name + '.html').render(**ctx)


class MobileTemplates(unittest.TestCase):
    def test_every_user_page_renders_without_app_or_data(self):
        env = template_environment()
        for name in PAGES:
            with self.subTest(page=name):
                page = rendered_page(env, name)
                self.assertIn('<html', page.lower())
                self.assertIn('viewport', page)
                self.assertNotIn('{%', page)


class MobileBrowser(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            from playwright.sync_api import sync_playwright, Error
        except ImportError as error:
            raise unittest.SkipTest('Playwright is not installed') from error
        cls.runtime = sync_playwright().start()
        try:
            cls.browser = cls.runtime.chromium.launch(headless=True)
        except Error as error:
            cls.runtime.stop()
            raise unittest.SkipTest('Local Chromium cannot launch in this environment') from error
        cls.templates = template_environment()

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.runtime.stop()

    def open_page(self, name, width, height, theme):
        context = self.browser.new_context(viewport={'width': width, 'height': height},
                     locale='ru-RU', has_touch=width < 800, reduced_motion='reduce', service_workers='block')
        context.add_init_script("localStorage.setItem('neurona-theme-mode', %s);"
                               "localStorage.setItem('neurona-glass', %s);" %
                               (json.dumps('dark' if theme == 'dark' else 'light'),
                                json.dumps('on' if theme == 'glass' else 'off')))
        context.add_init_script("window.WebSocket = class {}; window.EventSource = class {}; navigator.sendBeacon = () => true;")
        document = rendered_page(self.templates, name)
        def route(request):
            url = urlsplit(request.request.url)
            if url.hostname == 'ui.test' and url.path == '/qa':
                request.fulfill(content_type='text/html', body=document)
                return
            if url.hostname == 'ui.test' and url.path.startswith('/static/'):
                asset = (ROOT / url.path.lstrip('/')).resolve()
                if asset.is_relative_to(ROOT / 'static') and asset.is_file():
                    request.fulfill(path=asset, content_type=mimetypes.guess_type(str(asset))[0] or 'application/octet-stream')
                else:
                    request.fulfill(status=404)
                return
            if url.hostname != 'ui.test':
                # Only local library copies. No remote request is ever allowed through.
                local = {'chart.umd.min.js': 'chart.umd.min.js', 'leaflet.js': 'leaflet.min.js'}.get(url.path.rsplit('/', 1)[-1])
                asset = ROOT / 'static/vendor' / str(local)
                if local and asset.is_file():
                    request.fulfill(path=asset, content_type='text/javascript')
                else:
                    request.abort()
                return
            result = {'ok': True, 'running': False, 'status': 'idle', 'items': [],
                'rows': [], 'data': [], 'results': [], 'users': [], 'is_parent_manager': False,
                'events': [], 'dialogs': [], 'notifications': [], 'releases': [],
                'pending': [], 'clean': {}, 'configured': False, 'sources': [], 'modules': [],
                'completed': True, 'eligible': False, 'should_invite': False, 'questions': []}
            if url.path in ('/api/ux-survey/state', '/api/ux-survey/open', '/api/ux-survey/draft'):
                result = {'completed': False, 'eligible': False, 'draft_version': 0, 'draft': {},
                    'survey': {'questions': [{'id': 'q'+str(i), 'label': 'Насколько удобно пользоваться функциями и читать текст сервиса?', 'low': 'совсем неудобно', 'high': 'очень удобно'} for i in range(1,7)]}}
            elif url.path == '/water-dashboard/snapshot':
                result = ctx['snapshot']
            elif url.path == '/water-dashboard/run-status':
                result = {'running': False, 'auto_refresh_needed': False, 'snapshot_revision': 'qa'}
            elif url.path.endswith('/dictionary'):
                result = {'dict': {}, 'categories': ['Трубы'], 'groupsByCat': {}, 'units': ['шт', 'пог. м'], 'umap': {}, 'keywords': [], 'rso2omsu': {}}
            elif url.path == '/api/history/recent':
                result = []
            request.fulfill(content_type='application/json', body=json.dumps(result, ensure_ascii=False))
        context.route('**/*', route)
        page = context.new_page()
        page.goto('http://ui.test/qa', wait_until='load')
        page.wait_for_timeout(100)
        return context, page

    def check_layout(self, name, size, theme):
        context, page = self.open_page(name, *size, theme)
        try:
            sizes = page.evaluate('({viewport:document.documentElement.clientWidth,content:document.documentElement.scrollWidth})')
            self.assertLessEqual(sizes['content'], sizes['viewport'] + 2, (name, size, theme, sizes))
            if name == 'water_dashboard':
                for metric in ('system_addresses', 'resonant_addresses', 'electronic_documents_pct'):
                    self.assertGreater(page.locator('[data-metric=\"' + metric + '\"] .wd-num').count(), 0)
                    self.assertNotEqual(page.locator('[data-metric=\"' + metric + '\"] .wd-num').first.inner_text(), '—')
            output = os.getenv('NEURONA_UI_SCREENSHOTS')
            if output:
                folder = Path(output); folder.mkdir(parents=True, exist_ok=True)
                page.screenshot(path=str(folder / (name + '-' + str(size[0]) + '-' + theme + '.png')), full_page=True)
            if name == 'ux_survey':
                choice = page.locator('#uxSurvey-q1 label.ux-survey-choice').first
                choice.click()
                self.assertTrue(page.locator('input[name=q1][value="0"]').is_checked())
            trigger = page.locator('.appearance-trigger').first
            if trigger.count() and trigger.is_visible():
                trigger.click()
                self.assertEqual(trigger.get_attribute('aria-expanded'), 'true')
                panel = page.locator('.appearance-panel').first
                if not panel.count():
                    panel = page.locator('.appearance-menu').first
                if panel.count() and panel.is_visible():
                    box = panel.bounding_box()
                    self.assertGreaterEqual(box['x'], -2)
                    self.assertGreaterEqual(box['y'], -2)
                    self.assertLessEqual(box['y'] + box['height'], size[1] + 2)
                    self.assertLessEqual(box['x'] + box['width'], size[0] + 2)
                glass = page.locator('#glassMenuToggle')
                if glass.count():
                    old = glass.get_attribute('aria-checked')
                    glass.click()
                    self.assertNotEqual(glass.get_attribute('aria-checked'), old)
                page.keyboard.press('Escape')
                self.assertEqual(trigger.get_attribute('aria-expanded'), 'false')
                self.assertTrue(trigger.evaluate('(element) => document.activeElement === element'))
            profile = page.locator('#profileMenuButton')
            if profile.count() and profile.is_visible():
                profile.click()
                self.assertEqual(profile.get_attribute('aria-expanded'), 'true')
                box = page.locator('#profileMenuDropdown').bounding_box()
                self.assertIsNotNone(box)
                self.assertGreaterEqual(box['x'], -2)
                self.assertLessEqual(box['x'] + box['width'], size[0] + 2)
                self.assertGreaterEqual(box['y'], -2)
                self.assertLessEqual(box['y'] + box['height'], size[1] + 2)
                page.keyboard.press('Escape')
                self.assertEqual(profile.get_attribute('aria-expanded'), 'false')
                self.assertTrue(profile.evaluate('(element) => document.activeElement === element'))
        finally:
            context.close()

    def test_all_pages_at_phone_width(self):
        for name in PAGES:
            with self.subTest(page=name):
                self.check_layout(name, (360, 800), 'light')

    def test_key_pages_phone_tablet_desktop_and_landscape(self):
        pages = ['home', 'aichat', 'tools', 'users', 'water_dashboard', 'collective', 'mingkh-water-map', 'ux_survey', 'login']
        for name in pages:
            for size, theme in [((320,740),'light'), ((390,844),'dark'), ((768,1024),'glass'), ((1440,900),'light'), ((740,375),'glass')]:
                with self.subTest(page=name, size=size, theme=theme):
                    self.check_layout(name, size, theme)


if __name__ == '__main__':
    unittest.main()
