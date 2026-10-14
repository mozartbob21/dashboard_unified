"""Role-specific map rendering; no app startup, DB or portal access."""
import os
from html.parser import HTMLParser
from pathlib import Path
import unittest
from jinja2 import Environment, FileSystemLoader, select_autoescape

ROOT = Path(os.environ.get('MINGKH_UI_ROOT', Path(__file__).resolve().parents[1]))
BASE = Path(os.environ.get('MINGKH_BASE_ROOT', ROOT))

class Elements(HTMLParser):
    def __init__(self, html):
        super().__init__()
        self.ids = {}
        self.feed(html)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if attrs.get('id'):
            self.ids[attrs['id']] = (tag, attrs)

class WaterMapViewTests(unittest.TestCase):
    def render(self, **context):
        env = Environment(loader=FileSystemLoader([ROOT / 'templates', BASE / 'templates']), autoescape=select_autoescape())
        html = env.get_template('mingkh-water-map.html').render(
            url_for=lambda name, **values: '/static/' + values['path'], map_config={}, **context)
        return Elements(html), html

    def test_readers_get_map_filters_and_status_without_maintenance(self):
        for context in ({'can_import': False}, {}):
            with self.subTest(context=context):
                elements, html = self.render(**context)
                for control in ['map-maintenance', 'archive-file', 'import-form', 'import-submit', 'live-refresh']:
                    self.assertNotIn(control, elements.ids)
                for control in ['dataset-update-status', 'map', 'per', 'kpis', 'ot', 'xls']:
                    self.assertIn(control, elements.ids)
                self.assertNotIn('сохранённый доступ', html)
                self.assertNotIn('water_points.json', html)

    def test_editors_get_closed_maintenance_and_the_same_passive_status(self):
        elements, _ = self.render(can_import=True)
        tag, attrs = elements.ids['map-maintenance']
        self.assertEqual(tag, 'details')
        self.assertNotIn('open', attrs)
        for control in ['archive-file', 'import-form', 'live-refresh', 'dataset-update-status', 'map']:
            self.assertIn(control, elements.ids)
        self.assertEqual(elements.ids['dataset-update-status'][1]['role'], 'status')

if __name__ == '__main__':
    unittest.main()
