"""Agent Skills metadata, lazy I/O and account isolation; all inputs are synthetic."""
import contextlib
import json
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi import HTTPException
from services.aichat.skills import registry, preferences


class SkillRegistryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / 'skills'
        self.root.mkdir()
        self.db_file = Path(self.temp.name) / 'preferences.sqlite'
        root_patch = patch.object(registry, 'SKILLS_DIR', self.root)
        root_patch.start()
        self.addCleanup(root_patch.stop)

        @contextlib.contextmanager
        def isolated_db():
            conn = sqlite3.connect(self.db_file)
            conn.row_factory = sqlite3.Row
            try:
                yield conn
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()
        db_patch = patch.object(preferences, 'get_db_connection', isolated_db)
        db_patch.start()
        self.addCleanup(db_patch.stop)
        tools_patch = patch.object(preferences, '_supported_tool_names', return_value=frozenset({'calculate', 'table_profile', 'attachment_text', 'read_resource', 'platform_report'}))
        tools_patch.start()
        self.addCleanup(tools_patch.stop)
        self.local = {'id': 7, 'username': 'synthetic-user'}

    def pack(self, name='sample', body='Only selected skill instructions.', *, front=None):
        folder = self.root / name
        folder.mkdir(exist_ok=True)
        metadata = front if front is not None else (
            f'name: {name}\ndescription: "Анализирует тестовые сведения, когда нужен пример."\n'
            'metadata:\n  neurona-title: "Тестовый навык"\n  version: "1"\n'
            'allowed-tools: calculate read_resource\n'
        )
        (folder / 'SKILL.md').write_text('---\n' + metadata + '---\n\n' + body, encoding='utf-8')
        return folder

    def test_discovery_reads_only_metadata_and_picks_up_changes(self):
        folder = self.pack()
        header = (folder / 'SKILL.md').read_bytes().split(b'---', 2)[:2]
        # Invalid UTF-8 and an oversized body do not affect metadata discovery.
        (folder / 'SKILL.md').write_bytes(b'---' + header[1] + b'---\n' + b'\xff' * (registry.MAX_INSTRUCTION_BYTES + 1))
        with patch.object(registry, 'load_instructions', side_effect=AssertionError('eager body load')):
            items = registry.discover()
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]['id'], 'sample')
        self.assertEqual(items[0]['title'], 'Тестовый навык')
        self.assertEqual(items[0]['allowed_tools'], ['calculate', 'read_resource'])
        self.assertEqual(set(items[0]), {'id', 'name', 'title', 'description', 'allowed_tools', 'metadata'})
        with self.assertRaises(registry.SkillError):
            registry.load_instructions('sample')
        self.pack(body='Changed lazily.')
        self.assertEqual(registry.load_instructions('sample'), 'Changed lazily.')
        self.pack('another')
        self.assertEqual({x['id'] for x in registry.discover()}, {'sample', 'another'})
        (folder / 'SKILL.md').unlink()
        snapshot = registry.catalog_snapshot()
        self.assertEqual([x['id'] for x in snapshot['items']], ['another'])
        self.assertTrue(snapshot['issues'])

    def test_public_titles_stay_russian_without_exposing_technical_ids(self):
        for metadata in ('', 'metadata:\n  neurona-title: English title\n'):
            self.pack(front='name: sample\ndescription: Пример навыка\n' + metadata)
            self.assertEqual(registry.discover()[0]['title'], 'Дополнительный навык')

    def test_yaml_name_and_metadata_validation_isolated_from_good_packs(self):
        self.pack('good')
        invalid = [
            'name: other\ndescription: text\n',
            'name: invalid\ndescription: ""\n',
            'name: invalid\ndescription: ' + 'x' * 1025 + '\n',
            'name: invalid\ndescription: text\ncompatibility: ' + 'x' * 501 + '\n',
            'name: invalid\ndescription: text\nmetadata:\n  version: 1\n',
            'name: invalid\ndescription: text\nallowed-tools: [calculate]\n',
            'name: invalid\nname: invalid\ndescription: text\n',
            'name: invalid\ndescription: text\nmetadata:\n  value: &a secret\n  other: *a\n',
            'name: invalid\ndescription: !!python/object/apply:os.system [never-execute]\n',
            'name: invalid\ndescription: text\nmetadata:\n  title: [bad, type]\n',
        ]
        for front in invalid:
            with self.subTest(front=front[:70]):
                self.pack('invalid', front=front)
                snapshot = registry.catalog_snapshot()
                self.assertEqual([item['id'] for item in snapshot['items']], ['good'])
                self.assertTrue(snapshot['issues'])
                self.assertNotIn(self.temp.name, '\n'.join(snapshot['issues']))
                self.assertNotIn('never-execute', '\n'.join(snapshot['issues']))
        for name in ['-invalid', 'invalid-', 'invalid--name', 'Upper', 'a' * 65]:
            with self.subTest(name=name):
                self.pack(name)
                self.assertNotIn(name, [item['id'] for item in registry.discover()])
        for name in ['../good', '/etc', '', None, 'good/../../etc']:
            with self.subTest(name=name), self.assertRaises(registry.SkillError):
                registry.load_instructions(name)

    def test_bounded_header_scan_and_valid_multiline_description(self):
        self.pack(front='name: sample\ndescription: >\n  Первый абзац.\n  Продолжение описания.\n')
        self.assertIn('Продолжение', registry.discover()[0]['description'])
        self.assertEqual(registry.discover()[0]['allowed_tools'], [])
        (self.root / 'sample' / 'SKILL.md').write_text('---\nname: sample\ndescription: ' + 'x' * registry.MAX_FRONTMATTER_BYTES)
        self.assertEqual(registry.discover(), [])
        self.assertTrue(registry.catalog_snapshot()['issues'])
        (self.root / 'sample' / 'SKILL.md').unlink()
        (self.root / 'sample').rmdir()
        for i in range(registry.MAX_PACKS + 3):
            self.pack(f'skill-{i:03d}')
        snapshot = registry.catalog_snapshot()
        self.assertEqual(len(snapshot['items']), registry.MAX_PACKS)
        self.assertTrue(any('64' in issue for issue in snapshot['issues']))

    def test_resources_are_lazy_bounded_and_restricted_to_reference_text(self):
        folder = self.pack()
        refs = folder / 'references'; refs.mkdir()
        (refs / 'facts.json').write_text('{"value": 42}', encoding='utf-8')
        (refs / 'nested').mkdir(); (refs / 'nested' / 'guide.md').write_text('Useful text.', encoding='utf-8')
        (folder / 'secret.txt').write_text('OUTSIDE REFERENCES', encoding='utf-8')
        self.assertEqual(registry.read_resource('sample', 'references/facts.json'), '{"value": 42}')
        self.assertEqual(registry.read_resource('sample', 'references/nested/guide.md'), 'Useful text.')
        (refs / 'facts.json').write_text('changed', encoding='utf-8')
        self.assertEqual(registry.read_resource('sample', 'references/facts.json'), 'changed')
        for path in ['../secret.txt', '/etc/passwd', 'references/../../secret.txt', 'references/../secret.txt',
                     'references//guide.md', 'references/./guide.md', 'references\\facts.json',
                     'references/run.py', 'scripts/run.py', 'SKILL.md', 'C:/file.txt', 'references/facts.json:stream', None]:
            with self.subTest(path=path), self.assertRaises(registry.SkillError):
                registry.read_resource('sample', path)
        (refs / 'big.txt').write_bytes(b'x' * (registry.MAX_RESOURCE_BYTES + 1))
        (refs / 'binary.txt').write_bytes(b'\x00text')
        (refs / 'encoding.txt').write_bytes(b'\xff')
        for path in ['references/big.txt', 'references/binary.txt', 'references/encoding.txt']:
            with self.subTest(path=path), self.assertRaises(registry.SkillError):
                registry.read_resource('sample', path)
        (folder / 'SKILL.md').write_text('invalid now')
        with self.assertRaises(registry.SkillError):
            registry.read_resource('sample', 'references/facts.json')

    def test_symlinked_packs_skill_files_and_resources_cannot_escape(self):
        folder = self.pack()
        outside = Path(self.temp.name) / 'outside.txt'; outside.write_text('PRIVATE OUTSIDE')
        refs = folder / 'references'; refs.mkdir()
        try:
            (refs / 'escape.txt').symlink_to(outside)
        except (OSError, NotImplementedError):
            self.skipTest('Symlinks unavailable on this host')
        with self.assertRaises(registry.SkillError):
            registry.read_resource('sample', 'references/escape.txt')
        (refs / 'alias').symlink_to(Path(self.temp.name), target_is_directory=True)
        with self.assertRaises(registry.SkillError):
            registry.read_resource('sample', 'references/alias/outside.txt')
        (self.root / 'alias-pack').symlink_to(folder, target_is_directory=True)
        self.assertNotIn('alias-pack', [x['id'] for x in registry.discover()])
        self.assertTrue(registry.catalog_snapshot()['issues'])
        (folder / 'SKILL.md').unlink(); (folder / 'SKILL.md').symlink_to(outside)
        self.assertEqual(registry.discover(), [])
        with self.assertRaises(registry.SkillError):
            registry.load_instructions('sample')

    def test_preferences_are_account_scoped_and_default_enabled(self):
        self.pack('sample'); self.pack('another')
        other = {'id': 8, 'username': 'synthetic-other'}
        keycloak = {'kc_sub': '7', 'id': 7}
        catalog = preferences.catalog(self.local)
        self.assertEqual(catalog['scope'], 'account')
        self.assertTrue(all(item['enabled'] for item in catalog['items']))
        changed = preferences.set_enabled(self.local, 'sample', False)
        self.assertFalse(next(i for i in changed['items'] if i['id'] == 'sample')['enabled'])
        self.assertEqual([x['id'] for x in preferences.enabled_skills(self.local)], ['another'])
        self.assertEqual({x['id'] for x in preferences.enabled_skills(other)}, {'sample', 'another'})
        self.assertEqual({x['id'] for x in preferences.enabled_skills(keycloak)}, {'sample', 'another'})
        self.assertNotIn('enabled', preferences.enabled_skills(other)[0])
        preferences.set_enabled(keycloak, 'another', False)
        self.assertEqual([x['id'] for x in preferences.enabled_skills(keycloak)], ['sample'])
        # Reopening the database does not reset the account switch.
        self.assertEqual([x['id'] for x in preferences.enabled_skills(self.local)], ['another'])
        preferences.set_enabled(self.local, 'sample', True)
        self.assertEqual(len(preferences.enabled_skills(self.local)), 2)

    def test_invalid_or_removed_packs_cannot_be_enabled_and_no_anonymous_writes(self):
        self.pack()
        for user in [None, {}, {'username': 'without-stable-id'}]:
            with self.subTest(user=user), self.assertRaises(HTTPException) as raised:
                preferences.set_enabled(user, 'sample', False)
            self.assertEqual(raised.exception.status_code, 401)
        for enabled in ['false', 0, 1, None]:
            with self.subTest(enabled=enabled), self.assertRaises(HTTPException) as raised:
                preferences.set_enabled(self.local, 'sample', enabled)
            self.assertEqual(raised.exception.status_code, 400)
        for skill_id in ['unknown', '../sample', None]:
            with self.subTest(skill_id=skill_id), self.assertRaises(HTTPException) as raised:
                preferences.set_enabled(self.local, skill_id, False)
            self.assertEqual(raised.exception.status_code, 404)
        preferences.set_enabled(self.local, 'sample', False)
        (self.root / 'sample' / 'SKILL.md').unlink()
        self.assertEqual(preferences.enabled_skills(self.local), [])
        self.assertTrue(preferences.catalog(self.local)['issues'])
        with self.assertRaises(HTTPException):
            preferences.set_enabled(self.local, 'sample', True)
        self.pack()  # Reinstallation preserves the previous personal disabled choice.
        self.assertFalse(preferences.catalog(self.local)['items'][0]['enabled'])

    def test_catalog_never_advertises_unknown_tools_as_executable(self):
        self.pack(front='name: sample\ndescription: text\nallowed-tools: calculate Bash unknown_tool\n')
        item = preferences.catalog(self.local)['items'][0]
        self.assertEqual(item['allowed_tools'], ['calculate'])
        self.assertEqual(item['unsupported_tools'], ['Bash', 'unknown_tool'])
        self.assertTrue(any('не поддерживаются' in issue for issue in preferences.catalog(self.local)['issues']))
        # Registry preserves the declared semantics; execution-facing preferences narrow them.
        self.assertEqual(registry.discover()[0]['allowed_tools'], ['calculate', 'Bash', 'unknown_tool'])
        self.assertEqual(preferences.enabled_skills(self.local)[0]['allowed_tools'], ['calculate'])
        with patch.object(preferences, '_supported_tool_names', side_effect=ImportError('private host path')):
            catalog = preferences.catalog(self.local)
        self.assertEqual(catalog['items'][0]['allowed_tools'], [])
        self.assertTrue(catalog['issues'])
        self.assertNotIn('private host path', str(catalog['issues']))

    def test_builtin_packs_have_valid_metadata_and_readable_references(self):
        installed = Path(__file__).resolve().parents[1] / 'skills' / 'aichat'
        with patch.object(registry, 'SKILLS_DIR', installed):
            snapshot = registry.catalog_snapshot()
            self.assertEqual(snapshot['issues'], [])
            self.assertEqual({i['id'] for i in snapshot['items']}, {
                'calculations', 'table-analysis', 'document-analysis', 'business-writing',
                'platform-report', 'appeal-response', 'operational-brief',
            })
            for item in snapshot['items']:
                self.assertTrue(registry.load_instructions(item['id']))
                self.assertLessEqual(set(item['allowed_tools']), {'calculate', 'table_profile', 'attachment_text', 'read_resource', 'platform_report'})
                for resource in (installed / item['id'] / 'references').glob('*'):
                    self.assertTrue(registry.read_resource(item['id'], 'references/' + resource.name))


if __name__ == '__main__':
    unittest.main()
