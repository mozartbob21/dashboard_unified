"""Desktop/mobile interactions with the real directory API on an isolated SQLite DB."""
from urllib.parse import urlsplit
import pytest
from test.test_admin_directory import app, client


@pytest.fixture(scope='module')
def browser():
    playwright=pytest.importorskip('playwright.sync_api')
    with playwright.sync_playwright() as driver:
        try:
            browser=driver.chromium.launch()
        except playwright.Error as exc:
            pytest.skip('Локальный Chromium недоступен: '+str(exc).splitlines()[0])
        yield browser
        browser.close()


@pytest.mark.parametrize('width',[390,1280])
def test_user_directory_and_service_center(app,browser,width):
    context=browser.new_context(viewport={'width':width,'height':830})
    with client(app,'parent') as api:
        def serve(route):
            url=urlsplit(route.request.url)
            if url.netloc!='testserver':
                return route.abort()
            path=url.path+('?' + url.query if url.query else '')
            if (url.path in ['/users','/api/users/admin-directory','/api/users/service-center','/api/me/presence','/api/users/record']
                    or url.path.startswith('/static/')):
                response=api.request(route.request.method,path)
                route.fulfill(status=response.status_code,body=response.content,
                              content_type=response.headers.get('content-type','application/json'))
            else:
                # No network, integrations, production preferences or account secrets.
                route.fulfill(status=200,json={})
        context.route('**/*',serve)
        page=context.new_page()
        page.goto('http://testserver/users')
        page.wait_for_function("document.querySelectorAll('#userList button').length === 25")
        assert page.locator('.user-presence').count()==25
        directory=page.locator('.user-directory')
        workspace=page.locator('#userWorkspace')
        assert abs(directory.bounding_box()['width']-workspace.bounding_box()['width'])<2
        assert not page.locator('#userEditor').is_visible()
        page.locator('#directoryNext').click()
        page.wait_for_function("document.querySelector('#directoryPage').textContent.startsWith('2 из')")
        page.locator('#userSearch').fill('сотрудник21')
        page.wait_for_function("document.querySelectorAll('#userList button').length === 1")
        assert 'сотрудник21' in page.locator('#userList button').inner_text()
        page.locator('#userList button').press('Enter')
        page.wait_for_function("document.querySelector('#userEditor').hidden === false")
        assert directory.is_visible()==(width>780)
        if width>780: assert directory.bounding_box()['width']>320
        page.locator('#closeUserEditor').click()
        assert directory.is_visible() and page.locator('#userSearch').input_value()=='сотрудник21'
        page.locator('#userSearch').fill('неизвестный')
        page.wait_for_function("document.querySelector('.user-list-empty')")
        page.locator('#clearUserSearch').click()
        page.wait_for_function("document.querySelectorAll('#userList button').length === 25")
        page.locator('#userPresenceFilter').select_option('offline')
        page.wait_for_function("document.querySelectorAll('.user-presence.online').length === 0")
        page.locator('#myModules').click()
        page.wait_for_function("document.querySelector('#userForm input[name=username]').value === 'parent'")
        page.locator('[data-admin-link=service]').click()
        page.wait_for_function("document.querySelector('#serviceStamp').textContent.includes('Обновлено')")
        assert page.locator('#serviceMetrics dd').count()==7
        assert page.locator('#serviceScope').is_visible()
        assert page.evaluate("document.documentElement.scrollWidth <= innerWidth + 1")
    context.close()
