"""Render the real admin template offline: no app startup or database access."""
from collections import Counter
from html.parser import HTMLParser
import os
from pathlib import Path
import re

from jinja2 import ChainableUndefined, Environment, FileSystemLoader, select_autoescape
import pytest


ROOT = Path(os.environ.get('ADMIN_REPO_ROOT', Path(__file__).resolve().parents[1]))


class Document(HTMLParser):
    def __init__(self):
        super().__init__()
        self.elements = []

    def handle_starttag(self, tag, attrs):
        self.elements.append((tag, dict(attrs)))


def render(section):
    env = Environment(loader=FileSystemLoader(ROOT / 'templates'),
                      undefined=ChainableUndefined, autoescape=select_autoescape())
    env.globals['url_for'] = lambda route, **kwargs: '/static/' + kwargs['path']
    markup = env.get_template('users.html').render(
        module_names={'mingkh': 'МИНЖКХ', 'zips': 'Остатки ЗиП'}, initial_admin_section=section)
    document = Document()
    document.feed(markup)
    return document


@pytest.mark.parametrize('section', ['users', 'statistics', 'feedback', 'integrations'])
def test_rendered_admin_has_one_active_panel_and_unique_script_targets(section):
    document = render(section)
    ids = Counter(attrs['id'] for _, attrs in document.elements if 'id' in attrs)
    assert all(count == 1 for count in ids.values()), {name: count for name, count in ids.items() if count != 1}
    panels = [attrs for _, attrs in document.elements if 'data-admin-panel' in attrs]
    assert {attrs['data-admin-panel'] for attrs in panels} == {'users', 'statistics', 'feedback', 'integrations'}
    assert [attrs['data-admin-panel'] for attrs in panels if 'hidden' not in attrs] == [section]
    links = [attrs for _, attrs in document.elements if 'data-admin-link' in attrs]
    assert [attrs['data-admin-link'] for attrs in links if attrs.get('aria-current') == 'page'] == [section]
    for script in ['users.js', 'admin-navigation.js', 'module-popularity.js', 'ux-survey-admin.js']:
        source = (ROOT / 'static' / script).read_text()
        literal_targets = re.findall(r"(?:getElementById|node)\(['\"]([^'\"]+)['\"]\)", source)
        assert set(literal_targets) <= ids.keys(), (script, sorted(set(literal_targets) - ids.keys()))
    for _, attrs in document.elements:
        for reference in attrs.get('aria-labelledby', '').split():
            assert reference in ids, ('missing heading', reference)
        if 'for' in attrs:
            assert attrs['for'] in ids, ('missing input', attrs['for'])
    assert sum(tag == 'main' for tag, _ in document.elements) == 1
    assert sum(tag == 'h1' for tag, _ in document.elements) == 1
    assert {attrs.get('id') for tag, attrs in document.elements if tag == 'form'} == {'userForm', 'integrationForm', 'aiTokenForm', 'ratingKeyForm'}
    scripts = [attrs['src'].split('?')[0] for tag, attrs in document.elements if tag == 'script' and 'src' in attrs]
    assert scripts.index('/static/admin-navigation.js') < scripts.index('/static/users.js')
    assert scripts.index('/static/admin-navigation.js') < scripts.index('/static/module-popularity.js')
    assert scripts.index('/static/admin-navigation.js') < scripts.index('/static/ux-survey-admin.js')
