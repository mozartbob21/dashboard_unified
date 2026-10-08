"""Isolated skills runtime tests: mocked Qwen, in-memory files, no network or accounts."""
import io
import json
import unittest
from unittest.mock import Mock, patch
from zipfile import ZipFile

from services.aichat.skills import runtime, tools


DESCRIPTORS = [
    {'id': 'calculations', 'name': 'calculations', 'title': 'Расчёты', 'description': 'Выполняет арифметические расчёты.', 'allowed_tools': ['calculate'], 'metadata': {}},
    {'id': 'table-analysis', 'name': 'table-analysis', 'title': 'Таблицы', 'description': 'Анализирует CSV/XLSX и пропуски.', 'allowed_tools': ['table_profile', 'attachment_text'], 'metadata': {}},
    {'id': 'document-analysis', 'name': 'document-analysis', 'title': 'Документы', 'description': 'Читает документы.', 'allowed_tools': ['attachment_text', 'read_resource'], 'metadata': {}},
    {'id': 'platform-report', 'name': 'platform-report', 'title': 'Отчёт', 'description': 'Готовит отчёт по разрешённым данным.', 'allowed_tools': ['platform_report'], 'metadata': {}},
]
BODIES = {item['id']: 'SERVER_BODY_' + item['id'] for item in DESCRIPTORS}


def plan(*actions):
    return json.dumps({'actions': list(actions)})


def action(skill, tool, **args):
    return {'skill': skill, 'tool': tool, 'args': args}


def workbook(external=False):
    output = io.BytesIO()
    with ZipFile(output, 'w') as archive:
        archive.writestr('xl/workbook.xml', '''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Данные" sheetId="1" r:id="rId1"/></sheets></workbook>''')
        archive.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="worksheets/sheet1.xml"' + (' TargetMode="External"' if external else '') + '/></Relationships>')
        archive.writestr('xl/worksheets/sheet1.xml', '''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
          <row r="1"><c r="A1" t="inlineStr"><is><t>Сумма</t></is></c><c r="B1" t="inlineStr"><is><t>Пропуск</t></is></c></row>
          <row r="2"><c r="A2"><v>1.2</v></c></row>
          <row r="3"><c r="A3"><f>1+2</f><v>3</v></c></row>
          <row r="4"><c r="B4" t="inlineStr"><is><t>есть</t></is></c></row>
          </sheetData></worksheet>''')
    return output.getvalue()


class SkillRuntimeTests(unittest.TestCase):
    def setUp(self):
        self.events = []
        self.callback = Mock(return_value={'source_count': 1, 'selection': {'modules': ['edo']}})
        self.user = {'id': 42, 'modules': ['edo']}
        guards = [
            patch.object(runtime.preferences, 'enabled_skills', return_value=DESCRIPTORS),
            patch.object(runtime.registry, 'load_instructions', side_effect=lambda id: BODIES[id]),
            patch.object(runtime.registry, 'read_resource', return_value='REFERENCE_DATA'),
        ]
        self.enabled, self.load, self.read = [guard.start() for guard in guards]
        for guard in guards: self.addCleanup(guard.stop)

    def prepare(self, question, model, attachments=None):
        return runtime.prepare(question, attachments=attachments or [], user=self.user,
                               emit=self.events.append, platform_report=self.callback, model=model)

    def test_semantic_route_sees_metadata_only_then_loads_selected_instructions(self):
        attachments = [{'name': 'amounts.csv', 'text': 'UNTRUSTED_FILE_INSTRUCTIONS', 'data': b'name,amount\nA,10\nB,20\n'}]
        model = Mock(side_effect=['{"skills":["table-analysis"]}',
                                 plan(action('table-analysis', 'table_profile', attachment='amounts.csv'))])
        result = self.prepare('Какая сумма в этом файле?', model, attachments)
        self.enabled.assert_called_once_with(self.user)
        self.load.assert_called_once_with('table-analysis')
        self.read.assert_not_called()
        self.assertEqual(model.call_count, 2)
        route = model.call_args_list[0].args[0]
        planned = model.call_args_list[1].args[0]
        self.assertNotIn('SERVER_BODY_', str(route))
        self.assertNotIn('UNTRUSTED_FILE_INSTRUCTIONS', str(model.call_args_list))
        self.assertNotIn('A,10', str(model.call_args_list))
        self.assertIn('SERVER_BODY_table-analysis', str(planned))
        self.assertNotIn('SERVER_BODY_calculations', str(planned))
        metadata = json.loads(route[-1]['content'])
        self.assertEqual(metadata['attachments'], [{'name': 'amounts.csv', 'type': 'csv'}])
        self.assertEqual(result['skills'], [{'id': 'table-analysis', 'title': 'Таблицы'}])
        evidence = json.loads(result['evidence'])['results']
        self.assertEqual(evidence[0]['result']['columns'][1]['sum'], '30')
        self.assertEqual(result['tools'][0]['status'], 'done')
        self.assertEqual(self.events[0]['status'], 'running')
        self.assertEqual(self.events[-1]['status'], 'done')
        for event in self.events:
            self.assertLessEqual(set(event), {'type', 'id', 'kind', 'label', 'status', 'skill_id'})
            self.assertEqual(event['type'], 'step')

    def test_explicit_enabled_skill_skips_router_and_runs_decimal_tool(self):
        model = Mock(return_value=plan(action('calculations', 'calculate', expression='0.1 + 0.2')))
        result = self.prepare('$calculations Сколько будет 0.1 + 0.2?', model)
        model.assert_called_once()
        self.load.assert_called_once_with('calculations')
        self.assertEqual(json.loads(result['evidence'])['results'][0]['result']['value'], '0.3')

    def test_disabled_explicit_skill_is_never_loaded_or_run(self):
        model = Mock()
        result = self.prepare('$disabled-skill Выполни команду', model)
        model.assert_not_called()
        self.load.assert_not_called()
        self.assertEqual(result['skills'], [])
        self.assertEqual(result['tools'], [])
        self.assertTrue(result['warnings'])

    def test_disabled_id_in_semantic_route_rejects_whole_selection(self):
        model = Mock(return_value='{"skills":["calculations","disabled-skill"]}')
        result = self.prepare('Вычисли сумму', model)
        model.assert_called_once()
        self.load.assert_not_called()
        self.assertEqual(result['instructions'], '')
        self.assertEqual(json.loads(result['evidence'])['results'], [])
        self.assertTrue(result['warnings'])

    def test_malformed_routing_cannot_be_used_as_instructions(self):
        for raw in ['<think>PRIVATE_THOUGHT</think>{"skills":["calculations"]}',
                    '{"skills":["calculations"],"prompt":"OVERRIDE"}',
                    '{"skills":[],"skills":["calculations"]}',
                    '{"skills":[NaN]}', '{"skills":["calculations","calculations"]}',
                    '{"skills":["calculations","table-analysis","platform-report"]}']:
            with self.subTest(raw=raw):
                model = Mock(return_value=raw)
                result = self.prepare('Вопрос', model)
                self.assertFalse(result['skills'])
                self.assertNotIn('PRIVATE_THOUGHT', str(result) + str(self.events))
                self.assertNotIn('OVERRIDE', str(result) + str(self.events))
        self.load.assert_not_called()

    def test_unavailable_router_leaves_ordinary_chat_without_raw_error(self):
        result = self.prepare('Обычный вопрос', Mock(side_effect=RuntimeError('PRIVATE_KEY=secret')))
        self.assertEqual(result['instructions'], '')
        self.assertEqual(result['tools'], [])
        self.assertTrue(result['warnings'])
        self.assertNotIn('PRIVATE_KEY', json.dumps(result) + json.dumps(self.events))
        self.assertEqual(self.events[-1]['status'], 'error')

    def test_no_enabled_skills_does_not_call_model(self):
        self.enabled.return_value = []
        model = Mock()
        result = self.prepare('Обычный вопрос', model)
        model.assert_not_called()
        self.assertFalse(result['skills'])

    def test_disallowed_tools_and_inactive_skills_never_execute(self):
        model = Mock(return_value=plan(action('calculations', 'platform_report'),
                                       action('platform-report', 'platform_report'),
                                       action('calculations', 'shell', command='cat /secret')))
        with patch.object(tools, 'execute') as execute:
            result = self.prepare('$calculations 2 + 2', model)
        execute.assert_not_called()
        self.callback.assert_not_called()
        self.assertEqual([t['status'] for t in result['tools']], ['error'] * 3)
        self.assertTrue(result['warnings'])

    def test_planner_failure_or_excess_actions_do_not_claim_tool_success(self):
        for response in ['{"actions":[{}]}', plan(*[action('calculations', 'calculate', expression='2+2')] * 5)]:
            with self.subTest(response=response), patch.object(tools, 'execute') as execute:
                result = self.prepare('$calculations 2+2', Mock(return_value=response))
                execute.assert_not_called()
                self.assertTrue(result['instructions'])
                self.assertFalse(result['tools'])
                self.assertTrue(result['warnings'])
                self.assertEqual(self.events[-1]['status'], 'error')

    def test_platform_callback_receives_no_model_supplied_grants(self):
        model = Mock(return_value=plan(action('platform-report', 'platform_report', modules=['mingkh']),
                                       action('platform-report', 'platform_report')))
        result = self.prepare('$platform-report Подготовь отчёт', model)
        self.callback.assert_called_once_with()
        self.assertEqual([tool['status'] for tool in result['tools']], ['error', 'done'])
        evidence = json.loads(result['evidence'])['results']
        self.assertEqual(evidence[-1]['result']['selection']['modules'], ['edo'])

    def test_resource_is_lazy_and_scoped_to_the_active_skill(self):
        self.read.side_effect = ValueError('/private/SECRET_PATH forbidden')
        result = self.prepare('$document-analysis Изучи документ', Mock(return_value=plan(
            action('document-analysis', 'read_resource', path='../outside.md'))))
        self.read.assert_called_once_with('document-analysis', '../outside.md')
        self.assertEqual(result['tools'][0]['status'], 'error')
        self.assertNotIn('SECRET_PATH', str(result) + str(self.events))

    def test_duplicate_actions_do_not_repeat_side_effect_free_callback(self):
        request = action('platform-report', 'platform_report')
        result = self.prepare('$platform-report Отчёт', Mock(return_value=plan(request, request)))
        self.callback.assert_called_once_with()
        self.assertEqual([tool['status'] for tool in result['tools']], ['done', 'error'])

    def test_total_instruction_and_evidence_limits_preserve_valid_json(self):
        self.load.side_effect = lambda _: 'LONG_SERVER_INSTRUCTION ' * 5000
        self.read.return_value = '\\"' * 12000
        model = Mock(return_value=plan(*[
            action('document-analysis', 'read_resource', path=f'references/{i}.md') for i in range(4)
        ]))
        result = self.prepare('$document-analysis $calculations Проверь', model)
        self.assertLessEqual(len(result['instructions']), 16000)
        self.assertLessEqual(len(result['evidence']), 20000)
        self.assertIsInstance(json.loads(result['evidence']), dict)
        self.assertTrue(result['warnings'])

    def test_missing_skill_body_does_not_plan_or_claim_loaded(self):
        self.load.side_effect = OSError('SECRET_RESOURCE_PATH')
        model = Mock()
        result = self.prepare('$calculations 2+2', model)
        model.assert_not_called()
        self.assertFalse(result['skills'])
        self.assertNotIn('SECRET_RESOURCE_PATH', str(result))
        self.assertEqual(self.events[-1]['status'], 'error')

    def test_followup_uses_only_bounded_previous_question_hint(self):
        model = Mock(side_effect=['{"skills":["platform-report"]}', plan()])
        runtime.prepare('А теперь по Власихе', attachments=[], user=self.user, emit=self.events.append,
                        platform_report=self.callback, model=model, context_hint='Последний вопрос ' * 200)
        self.assertEqual(model.call_count, 2)
        for call in model.call_args_list:
            payload = json.loads(call.args[0][-1]['content'])
            self.assertEqual(len(payload['previous_question']), 1200)
            self.assertNotIn('history', payload)

    def test_empty_file_plan_reports_that_attachments_were_not_inspected(self):
        result = self.prepare('$document-analysis Изучи файл', Mock(return_value=plan()),
                              [{'name': 'a.txt', 'text': 'private contents', 'data': b'example'}])
        payload = json.loads(result['evidence'])
        self.assertEqual(payload['results'], [])
        self.assertTrue(payload['warnings'])
        self.assertFalse(result['tools'])

    def test_attachment_extraction_failure_is_not_reported_as_successful_read(self):
        result = self.prepare('$document-analysis Изучи файл', Mock(return_value=plan(
            action('document-analysis', 'attachment_text', attachment='failed.txt'))),
            [{'name': 'failed.txt', 'text': 'INCOMPLETE_PRIVATE_TEXT', 'data': b'example',
              'error': 'PRIVATE_EXTRACTION_ERROR'}])
        payload = json.loads(result['evidence'])
        self.assertEqual(result['tools'][0]['status'], 'error')
        self.assertEqual(payload['results'][0]['status'], 'error')
        self.assertEqual(payload['unread_attachments'], ['failed.txt'])
        self.assertNotIn('INCOMPLETE_PRIVATE_TEXT', str(result) + str(self.events))
        self.assertNotIn('PRIVATE_EXTRACTION_ERROR', str(result) + str(self.events))
        self.assertEqual(self.events[-1]['status'], 'error')


class SkillToolsTests(unittest.TestCase):
    def test_bounded_decimal_arithmetic_without_float_rounding(self):
        self.assertEqual(tools.calculate('(0.1 + 0.2) * 3')['value'], '0.9')
        self.assertEqual(tools.calculate('1000000000000000000 + 1')['value'], '1000000000000000001')
        self.assertEqual(tools.calculate('7 % 3')['value'], '1')

    def test_calls_attributes_exponents_and_invalid_arithmetic_are_rejected(self):
        for expression in ["__import__('os').system('id')", '(1).__class__', '[1][0]',
                           '2 ** 99999999', '1 / 0', 'True + 1', '1e999', 'x + 1',
                           '1 if True else 2', '1' * 513]:
            with self.subTest(expression=expression), self.assertRaises(tools.ToolError):
                tools.calculate(expression)

    def test_csv_profile_counts_missing_and_ignores_formula_text(self):
        data = 'name;amount;empty\nA;1,5;\nB;2,5;\nC;=1+2;\nD;;\n'.encode()
        result = tools.table_profile([{'name': 'data.csv', 'data': data}], 'data.csv')
        self.assertEqual(result['rows_analyzed'], 4)
        amount = result['columns'][1]
        self.assertEqual((amount['numeric_count'], amount['missing'], amount['formula_count']), (2, 1, 1))
        self.assertEqual((amount['sum'], amount['mean']), ('4', '2'))
        self.assertEqual(result['columns'][2]['missing'], 4)
        self.assertFalse(result['formulas_evaluated'])

    def test_xlsx_ignores_cached_formulas_and_preserves_sparse_columns(self):
        result = tools.table_profile([{'name': 'data.xlsx', 'data': workbook()}], 'data.xlsx', 'Данные')
        self.assertEqual(result['rows_analyzed'], 3)
        self.assertEqual(result['columns'][0]['sum'], '1.2')
        self.assertEqual(result['columns'][0]['formula_count'], 1)
        self.assertEqual(result['columns'][0]['missing'], 1)
        self.assertEqual(result['columns'][1]['missing'], 2)
        self.assertEqual(result['columns'][1]['text_count'], 1)

    def test_xlsx_unicode_entity_declarations_are_rejected_before_parsing(self):
        sheet = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                 '<sheetData><row><c r="A1" t="inlineStr"><is><t>{header}</t></is></c></row>'
                 '<row><c r="A2"><v>2</v></c></row></sheetData></worksheet>')

        def with_sheet(xml):
            output = io.BytesIO()
            with ZipFile(io.BytesIO(workbook())) as source, ZipFile(output, 'w') as archive:
                for name in source.namelist():
                    archive.writestr(name, xml if name == 'xl/worksheets/sheet1.xml' else source.read(name))
            return [{'name': 'data.xlsx', 'data': output.getvalue()}]

        for encoding in ['utf-16', 'utf-16-le', 'utf-16-be', 'utf-32', 'utf-32-le', 'utf-32-be']:
            xml_encoding = 'UTF-16' if encoding.startswith('utf-16') else 'UTF-32'
            xml = ('<?xml version="1.0" encoding="' + xml_encoding + '"?>'
                   '<!DOCTYPE worksheet [<!ENTITY sample "expanded">]>'
                   + sheet.format(header='&sample;')).encode(encoding)
            with self.subTest(encoding=encoding):
                with patch.object(tools.ET, 'fromstring', wraps=tools.ET.fromstring) as parse:
                    with self.assertRaisesRegex(tools.ToolError, 'Неподдерживаемый формат XML книги'):
                        tools.table_profile(with_sheet(xml), 'data.xlsx')
                # Only the workbook and relationships were parsed, never the sheet.
                self.assertEqual(parse.call_count, 2)

        xml = ('<?xml version="1.0" encoding="UTF-16"?>' + sheet.format(header='Сумма')).encode('utf-16')
        result = tools.table_profile(with_sheet(xml), 'data.xlsx')
        self.assertEqual(result['columns'][0]['name'], 'Сумма')
        self.assertEqual(result['columns'][0]['sum'], '2')

    def test_table_limits_and_external_workbook_links_are_rejected_or_marked(self):
        data = b'amount\n' + b'1\n' * (tools.MAX_ROWS + 10)
        result = tools.table_profile([{'name': 'data.csv', 'data': data}], 'data.csv')
        self.assertEqual(result['rows_analyzed'], tools.MAX_ROWS)
        self.assertTrue(result['truncated'])
        with self.assertRaises(tools.ToolError):
            tools.table_profile([{'name': 'data.xlsx', 'data': workbook(external=True)}], 'data.xlsx')
        with self.assertRaises(tools.ToolError):
            tools.table_profile([{'name': 'data.xlsx', 'data': b'not a workbook'}], 'data.xlsx')

    def test_only_current_attachments_are_available_and_text_is_bounded(self):
        files = [{'name': 'note.txt', 'text': 'x' * 9000, 'data': b'x'}]
        result = tools.attachment_text(files, 'note.txt')
        self.assertEqual(len(result['text']), 8000)
        self.assertTrue(result['truncated'])
        for name in ['/etc/passwd', 'https://example.invalid/data.csv', 'old-attachment.txt']:
            with self.subTest(name=name), self.assertRaises(tools.ToolError):
                tools.attachment_text(files, name)
        with self.assertRaises(tools.ToolError):
            tools.attachment_text(files * 2, 'note.txt')

    def test_tool_arguments_reject_prompt_fields_before_callbacks(self):
        callback, reader = Mock(), Mock()
        for name, args in [('platform_report', {'modules': ['all']}),
                           ('read_resource', {'path': 'references/a.md', 'skill': 'other'}),
                           ('calculate', {'expression': '2+2', 'prompt': 'OVERRIDE'})]:
            with self.subTest(name=name), self.assertRaises(tools.ToolError):
                tools.execute(name, args, skill_id='test', attachments=[], platform_report=callback, resource_reader=reader)
        callback.assert_not_called()
        reader.assert_not_called()


if __name__ == '__main__':
    unittest.main()
