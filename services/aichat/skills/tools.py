"""Small, bounded skill tools. No shell, eval, network or filesystem access.

The sole resource reader and platform callback are supplied by the server. Table
formulas and external workbook relationships are never evaluated or followed.
"""
import ast
import csv
import io
import json
import posixpath
import re
import zipfile
from decimal import Decimal, DecimalException, Underflow, localcontext
from xml.etree import ElementTree as ET


MAX_ATTACHMENT_BYTES = 8 * 1024 * 1024
MAX_ROWS = 500
MAX_COLUMNS = 32
MAX_CELLS = (MAX_ROWS + 1) * MAX_COLUMNS
MAX_TEXT = 8000
_FORMULA = object()
_NUMERIC = re.compile(r'^[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)(?:[eE][+-]?\d{1,3})?$')
_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

SPECS = {
    'calculate': {'title': 'Вычисление', 'description': 'Точная арифметика чисел: + - * / % и скобки. Без функций и степеней.',
                  'args': {'expression': 'string, required, up to 512 characters'}},
    'table_profile': {'title': 'Анализ таблицы', 'description': 'Профиль текущего CSV/XLSX: пропуски и числовые итоги. Первая строка — заголовок. Формулы не вычисляются.',
                      'args': {'attachment': 'exact current attachment name, required', 'sheet': 'optional XLSX sheet name'}},
    'attachment_text': {'title': 'Чтение вложения', 'description': 'Текст только указанного вложения текущего сообщения, до 8000 символов.',
                        'args': {'attachment': 'exact current attachment name, required'}},
    'read_resource': {'title': 'Чтение справки навыка', 'description': 'Текст справки внутри активного навыка, до 8000 символов.',
                      'args': {'path': 'relative references/ path from this skill instructions, required'}},
    'platform_report': {'title': 'Данные платформы', 'description': 'Сохранённые данные по исходному вопросу и разрешённым пользователю блокам. Права и вопрос задаёт сервер.',
                        'args': {}},
}


class ToolError(ValueError):
    """Only fixed, safe user-facing messages should cross the tool boundary."""


def _arguments(args, required, optional=()):
    if not isinstance(args, dict) or set(args) - set(required) - set(optional) or not set(required) <= set(args):
        raise ToolError('Некорректные параметры инструмента.')


def _string(value, limit):
    if not isinstance(value, str) or not value.strip() or len(value) > limit:
        raise ToolError('Некорректные параметры инструмента.')
    return value


def _number(value):
    text = re.sub(r'[\s\u00a0]', '', str(value))
    if len(text) > 160 or not _NUMERIC.fullmatch(text):
        return None
    try:
        result = Decimal(text.replace(',', '.'))
        if not result.is_finite() or (result and abs(result.adjusted()) > 100):
            return None
        return result
    except DecimalException:
        return None


def _display(value):
    if not value:
        return '0'
    text = format(value, 'f')
    return text.rstrip('0').rstrip('.') if '.' in text else text


def calculate(expression):
    expression = _string(expression, 512).strip()
    try:
        tree = ast.parse(expression, mode='eval')
        if sum(1 for _ in ast.walk(tree)) > 100:
            raise ToolError('Вычисление слишком сложное. Упростите выражение.')

        def visit(node, depth=0):
            if depth > 20:
                raise ToolError('Вычисление слишком сложное. Упростите выражение.')
            if isinstance(node, ast.Expression):
                result = visit(node.body, depth + 1)
            elif isinstance(node, ast.Constant) and type(node.value) in (int, float):
                # Use the original token, not an already rounded Python float.
                token = ast.get_source_segment(expression, node)
                result = _number(token) if token else None
                if result is None:
                    raise ToolError('Допустимы только конечные десятичные числа ограниченного размера.')
            elif isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
                value = visit(node.operand, depth + 1)
                result = value if isinstance(node.op, ast.UAdd) else -value
            elif isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Mod)):
                left, right = visit(node.left, depth + 1), visit(node.right, depth + 1)
                if isinstance(node.op, ast.Add): result = left + right
                elif isinstance(node.op, ast.Sub): result = left - right
                elif isinstance(node.op, ast.Mult): result = left * right
                elif isinstance(node.op, ast.Div): result = left / right
                else: result = left % right
            else:
                raise ToolError('Допустимы только числа, скобки и операции + − × / %.')
            if not result.is_finite() or (result and abs(result.adjusted()) > 100):
                raise ToolError('Результат выходит за допустимые пределы.')
            return result

        with localcontext() as context:
            context.prec, context.Emax, context.Emin = 40, 100, -100
            context.traps[Underflow] = True
            result = visit(tree)
        return {'expression': expression, 'value': _display(result), 'precision_digits': 40}
    except ToolError:
        raise
    except (SyntaxError, ValueError, DecimalException, RecursionError, OverflowError):
        raise ToolError('Не удалось вычислить выражение. Проверьте числа и деление на ноль.') from None


def _attachment(attachments, name):
    name = _string(name, 240)
    found = [item for item in attachments if isinstance(item, dict) and item.get('name') == name]
    if len(found) != 1:
        raise ToolError('Вложение текущего сообщения не найдено или его имя неоднозначно.')
    return found[0]


def attachment_text(attachments, name):
    item = _attachment(attachments, name)
    if item.get('error'):
        raise ToolError('Из этого вложения не удалось извлечь текст.')
    text = item.get('text')
    if not isinstance(text, str) or not text.strip():
        raise ToolError('У текущего вложения нет доступного текстового содержимого.')
    return {'attachment': name, 'text': text[:MAX_TEXT], 'truncated': len(text) > MAX_TEXT}


def _csv_rows(data):
    if len(data) > 2 * 1024 * 1024:
        raise ToolError('CSV слишком большой. Приложите файл до 2 МБ.')
    try:
        text = data.decode('utf-8-sig')
    except UnicodeDecodeError:
        try: text = data.decode('cp1251')
        except UnicodeDecodeError:
            raise ToolError('Не удалось прочитать кодировку CSV.') from None
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=',;\t')
    except csv.Error:
        dialect = csv.excel
    rows, truncated = [], False
    try:
        for row in csv.reader(io.StringIO(text), dialect):
            if len(rows) >= MAX_ROWS + 1:
                truncated = True
                break
            truncated = truncated or len(row) > MAX_COLUMNS
            rows.append(row[:MAX_COLUMNS])
    except csv.Error:
        raise ToolError('Некорректный формат CSV или слишком длинное поле.') from None
    return rows, truncated, ''


def _xlsx_rows(data, requested_sheet):
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            parts = archive.infolist()
            if (len(parts) > 1000 or len({part.filename for part in parts}) != len(parts)
                    or sum(part.file_size for part in parts) > 24 * 1024 * 1024
                    or any(part.file_size > 8 * 1024 * 1024 for part in parts)):
                raise ToolError('Книга слишком большая для безопасного анализа.')

            def xml(name):
                raw = archive.read(name)
                # XML declarations can use UTF-16/32. Remove their NUL padding
                # before scanning, so entities are rejected before XML parsing.
                declarations = raw.replace(b'\x00', b'').upper()
                if b'<!DOCTYPE' in declarations or b'<!ENTITY' in declarations:
                    raise ToolError('Неподдерживаемый формат XML книги.')
                return ET.fromstring(raw)

            workbook = xml('xl/workbook.xml')
            sheets_node = workbook.find(_NS + 'sheets')
            sheets = list(sheets_node) if sheets_node is not None else []
            sheet = next((s for s in sheets if s.get('name') == requested_sheet), None) if requested_sheet else next(iter(sheets), None)
            if sheet is None:
                raise ToolError('Указанный лист не найден в текущей книге.')
            relations = xml('xl/_rels/workbook.xml.rels')
            relation = next((r for r in relations if r.get('Id') == sheet.get(_REL)), None)
            if relation is None or relation.get('TargetMode') == 'External':
                raise ToolError('Неподдерживаемая ссылка на лист книги.')
            target = relation.get('Target', '')
            if '\\' in target or not target:
                raise ToolError('Некорректная ссылка на лист книги.')
            path = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            if not path.startswith('xl/') or '..' in path.split('/'):
                raise ToolError('Некорректная ссылка на лист книги.')
            shared = []
            if 'xl/sharedStrings.xml' in archive.namelist():
                total = 0
                for entry in xml('xl/sharedStrings.xml'):
                    value = ''.join(node.text or '' for node in entry.iter(_NS + 't'))
                    total += len(value)
                    if len(shared) >= 50000 or total > 2000000:
                        raise ToolError('Справочник строк книги слишком большой.')
                    shared.append(value)
            rows, truncated, cell_count = [], False, 0
            for row in xml(path).iter(_NS + 'row'):
                if len(rows) >= MAX_ROWS + 1:
                    truncated = True
                    break
                values = []
                for cell in row.findall(_NS + 'c'):
                    cell_count += 1
                    if cell_count > MAX_CELLS:
                        truncated = True
                        break
                    ref = cell.get('r', '')
                    letters = re.match(r'^([A-Z]{1,3})\d+$', ref)
                    index = 0
                    if letters:
                        for letter in letters[1]: index = index * 26 + ord(letter) - 64
                        index -= 1
                    else:
                        index = len(values)
                    if index >= MAX_COLUMNS:
                        truncated = True
                        continue
                    while len(values) <= index: values.append('')
                    node = cell.find(_NS + 'v')
                    value = node.text or '' if node is not None else ''
                    if cell.find(_NS + 'f') is not None:
                        value = _FORMULA
                    elif cell.get('t') == 's':
                        try:
                            shared_index = int(value)
                            if shared_index < 0: raise ValueError()
                            value = shared[shared_index]
                        except (ValueError, IndexError):
                            raise ToolError('Некорректный справочник строк книги.') from None
                    elif cell.get('t') == 'inlineStr':
                        value = ''.join(n.text or '' for n in cell.iter(_NS + 't'))
                    values[index] = value
                rows.append(values)
                if cell_count > MAX_CELLS: break
            return rows, truncated, sheet.get('name', '')
    except ToolError:
        raise
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError, ValueError, RuntimeError):
        raise ToolError('Не удалось прочитать структуру XLSX.') from None


def table_profile(attachments, name, sheet=None):
    item = _attachment(attachments, name)
    data = item.get('data')
    if not isinstance(data, bytes) or len(data) > MAX_ATTACHMENT_BYTES:
        raise ToolError('Нужен файл таблицы размером до 8 МБ.')
    if name.lower().endswith('.csv'):
        if sheet is not None: raise ToolError('Параметр листа применим только к XLSX.')
        rows, truncated, sheet_name = _csv_rows(data)
    elif name.lower().endswith('.xlsx'):
        rows, truncated, sheet_name = _xlsx_rows(data, sheet)
    else:
        raise ToolError('Поддерживаются таблицы CSV и XLSX текущего сообщения.')
    if not rows:
        raise ToolError('В таблице нет строк для анализа.')
    width = min(MAX_COLUMNS, max(len(row) for row in rows))
    header, body, columns = rows[0], rows[1:], []
    for index in range(width):
        label = str(header[index]) if index < len(header) and header[index] is not _FORMULA else ''
        stats = {'name': label[:160] or f'Столбец {index + 1}', 'missing': 0, 'numeric_count': 0,
                 'text_count': 0, 'formula_count': 0}
        numeric = []
        for row in body:
            value = row[index] if index < len(row) else ''
            if value is _FORMULA or str(value).lstrip().startswith('='):
                stats['formula_count'] += 1
            elif not str(value).strip():
                stats['missing'] += 1
            else:
                number = _number(value)
                if number is None: stats['text_count'] += 1
                else: numeric.append(number)
        stats['numeric_count'] = len(numeric)
        if numeric:
            with localcontext() as context:
                context.prec = 40
                total = sum(numeric, Decimal(0))
                stats.update({'min': _display(min(numeric)), 'max': _display(max(numeric)),
                              'sum': _display(total), 'mean': _display(total / len(numeric))})
        columns.append(stats)
    return {'attachment': name, 'sheet': sheet_name, 'rows_analyzed': len(body),
            'columns_analyzed': width, 'truncated': truncated,
            'formulas_evaluated': False, 'columns': columns}


def execute(name, args, *, skill_id, attachments, platform_report, resource_reader):
    """Dispatch only the fixed allowlist; runtime additionally enforces skill grants."""
    if name == 'calculate':
        _arguments(args, ('expression',))
        return calculate(args['expression'])
    if name in ('attachment_text', 'table_profile'):
        _arguments(args, ('attachment',), ('sheet',) if name == 'table_profile' else ())
        if name == 'attachment_text': return attachment_text(attachments, args['attachment'])
        sheet = _string(args['sheet'], 160) if 'sheet' in args else None
        return table_profile(attachments, args['attachment'], sheet)
    if name == 'read_resource':
        _arguments(args, ('path',))
        path = _string(args['path'], 240)
        # Registry performs canonical containment, symlink and extension checks.
        try: text = resource_reader(skill_id, path)
        except Exception: raise ToolError('Справка недоступна для активного навыка.') from None
        if not isinstance(text, str): raise ToolError('Не удалось прочитать справку навыка.')
        return {'path': path, 'text': text[:MAX_TEXT], 'truncated': len(text) > MAX_TEXT}
    if name == 'platform_report':
        _arguments(args, ())
        if not callable(platform_report): raise ToolError('Данные платформы сейчас недоступны.')
        try:
            result = platform_report()
            if result is None: raise ValueError()
            json.dumps(result, ensure_ascii=False, allow_nan=False)
            return result
        except Exception: raise ToolError('Не удалось получить разрешённые данные платформы.') from None
    raise ToolError('Инструмент недоступен.')
