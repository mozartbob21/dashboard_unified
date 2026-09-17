import copy
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from services.zip_curator import core


class UploadDateTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        root = Path(temp.name)
        for name, path in {'STATE_FILE': root / 'state.json', 'PUBLISHED_FILE': root / 'published.json',
                           'MUNICIPALITY_OVERRIDES_FILE': root / 'municipalities.json'}.items():
            mocked = patch.object(core, name, path)
            mocked.start()
            self.addCleanup(mocked.stop)

    def registry(self, name='РСО А'):
        return {'rso': name, 'date': '01.09.2026', 'fname': 'source.xlsx',
                'items': [{'name': 'Труба', 'qty': 3, 'unitRaw': 'шт'}]}

    def test_upload_date_survives_approval_republication_and_restore(self):
        core.ingest([self.registry()])
        loaded = core.load_state()['pending'][0]['loaded_at']
        self.assertTrue(core.approve(0))
        core.publish_clean()
        approved = core.load_state()['clean'][core.norm('РСО А')]
        self.assertEqual(approved['loaded_at'], loaded)
        self.assertEqual(approved['date'], '01.09.2026')
        core.STATE_FILE.unlink()
        self.assertEqual(core.load_state()['clean'][core.norm('РСО А')]['loaded_at'], loaded)

    def test_rescan_keeps_original_upload_date_for_same_source(self):
        source = self.registry()
        source.update(local_file='folder_a/source.xlsx', uploaded=1234567890000)
        core.ingest([copy.deepcopy(source)])
        state = core.load_state()
        state['pending'][0]['loaded_at'] = '2026-09-01T10:00:00+03:00'
        core.save_state(state)
        core.ingest([copy.deepcopy(source)])
        self.assertEqual(core.load_state()['pending'][0]['loaded_at'], '2026-09-01T10:00:00+03:00')
        source['uploaded'] += 1000
        core.ingest([source])
        self.assertNotEqual(core.load_state()['pending'][0]['loaded_at'], '2026-09-01T10:00:00+03:00')

    def test_other_rso_upload_date_survives_update_and_delete(self):
        for name in ['РСО А', 'РСО Б']:
            core.ingest([self.registry(name)])
            core.approve(0)
        key = core.norm('РСО Б')
        before = core.load_state()['clean'][key]['loaded_at']
        core.ingest([self.registry()])
        core.approve(0)
        self.assertEqual(core.load_state()['clean'][key]['loaded_at'], before)
        core.delete_rso('РСО А')
        core.STATE_FILE.unlink()
        clean = core.load_state()['clean']
        self.assertEqual(list(clean), [key])
        self.assertEqual(clean[key]['loaded_at'], before)

    def test_legacy_rows_do_not_invent_upload_date_from_publication(self):
        rows = [['РСО', 'Наименование', 'Дата'], ['Старое РСО', 'Труба', '01.09.2026']]
        core.PUBLISHED_FILE.write_text(json.dumps({'rows': rows, 'published_at': '2026-09-17T14:00:00'}))
        core.publish_clean()
        self.assertNotIn('loaded_at', core.load_state()['clean'][core.norm('Старое РСО')])


if __name__ == '__main__':
    unittest.main()
