import ast
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi import FastAPI
from fastapi.testclient import TestClient
from services.zip_curator import core, dictionary


class DictionaryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for module, key, value in (
            (dictionary, 'CUSTOM_FILE', self.root / 'dictionary.json'),
            (core, 'STATE_FILE', self.root / 'state.json'),
            (core, 'PUBLISHED_FILE', self.root / 'published.json'),
        ):
            patcher = patch.object(module, key, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        dictionary._merged.cache_clear()
        self.addCleanup(dictionary._merged.cache_clear)

    def test_category_group_persist_and_normalized_duplicates_are_reused(self):
        original = dictionary.BASE_FILE.read_bytes()
        result = dictionary.add_category('  Новые   приборы ', ' Анализаторы ')
        self.assertTrue(result['created'])
        again = dictionary.add_category('НОВЫЕ ПРИБОРЫ', 'анализаторы')
        self.assertFalse(again['created'])
        self.assertEqual(again['category'], 'Новые приборы')
        dictionary._merged.cache_clear()  # Simulate fresh process cache.
        loaded = dictionary.load_dictionary()
        self.assertIn('Новые приборы', loaded['categories'])
        self.assertEqual(loaded['groupsByCat']['Новые приборы'], ['Анализаторы'])
        self.assertEqual(dictionary.BASE_FILE.read_bytes(), original)

    def test_add_group_to_existing_category(self):
        category = dictionary.load_dictionary()['categories'][0]
        original = dictionary.load_dictionary()['groupsByCat'][category]
        dictionary.add_category(category, 'Особая группа')
        loaded = dictionary.load_dictionary()
        self.assertTrue(set(original).issubset(loaded['groupsByCat'][category]))
        self.assertIn('Особая группа', loaded['groupsByCat'][category])

    def test_edit_teaches_future_imports_after_reload(self):
        dictionary.add_category('Измерительные системы', 'Анализаторы')
        core.save_state({'pending': [{'rso': 'РСО', 'items': [
            {'name': 'Новый анализатор XYZ', 'water': True}
        ]}], 'clean': {}})
        self.assertTrue(core.edit_item(0, 0, 'Измерительные системы', 'Анализаторы'))
        dictionary._merged.cache_clear()
        result = core.classify('  НОВЫЙ АНАЛИЗАТОР XYZ ')
        self.assertEqual(result, {'cat': 'Измерительные системы', 'grp': 'Анализаторы',
                                  'water': True, 'via': 'match'})
        saved = core.load_state()['pending'][0]['items'][0]
        self.assertEqual(saved['cat'], 'Измерительные системы')

    def test_validation_does_not_save_invalid_categories_or_groups(self):
        for value in (' ', 'a' * 121, ['category']):
            with self.assertRaises(ValueError):
                dictionary.add_category(value)
        self.assertFalse(dictionary.CUSTOM_FILE.exists())
        dictionary.add_category('Новая категория')
        with self.assertRaises(ValueError):
            dictionary.remember('позиция', 'Новая категория', 'Чужая группа')
        self.assertNotIn('позиция', dictionary.load_dictionary()['dict'])

    def test_category_and_edit_api(self):
        # Exercise the real handlers without starting the unrelated app scheduler.
        tree = ast.parse((core.BASE / 'app.py').read_text(encoding='utf-8-sig'))
        names = {'zc_dictionary', 'zc_add_category', 'zc_edit'}
        handlers = ast.Module(body=[node for node in tree.body
            if isinstance(node, ast.AsyncFunctionDef) and node.name in names], type_ignores=[])
        api = FastAPI()
        exec(compile(handlers, 'app.py', 'exec'), {'app': api, 'zc': core})
        client = TestClient(api)
        bad = client.post('/zip_curator/api/categories', json={'category': ''})
        self.assertEqual(bad.status_code, 400)
        response = client.post('/zip_curator/api/categories', json={'category': 'Новый раздел', 'group': 'Новая группа'})
        self.assertTrue(response.json()['ok'])
        self.assertIn('Новый раздел', client.get('/zip_curator/api/dictionary').json()['categories'])
        core.save_state({'pending': [{'rso': 'РСО', 'items': [{'name': 'Новая позиция'}]}], 'clean': {}})
        response = client.post('/zip_curator/api/edit', json={'pi': 0, 'ii': 0, 'cat': 'Новый раздел', 'grp': 'Новая группа'})
        self.assertTrue(response.json()['ok'])
        self.assertEqual(core.classify('Новая позиция')['cat'], 'Новый раздел')


if __name__ == '__main__':
    unittest.main()
