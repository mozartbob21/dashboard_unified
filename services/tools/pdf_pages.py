"""Local page operations using the PDF engine already used by PDF → Word."""
import io
import re
import threading
import zipfile
from contextlib import ExitStack, contextmanager

import fitz
from .workspace import ToolError

MAX_PAGES = 500
MAX_FILES = 20
ACTIONS = {'merge', 'extract', 'delete', 'reorder', 'rotate', 'split'}
PDF_LOCK = threading.RLock()


@contextmanager
def open_pdf(data):
    with PDF_LOCK:
        with _open_pdf(data) as doc:
            yield doc


@contextmanager
def _open_pdf(data):
    if not data.startswith(b'%PDF-'):
        raise ToolError('Файл не является PDF.')
    try:
        doc = fitz.open(stream=data, filetype='pdf')
    except (RuntimeError, ValueError):
        raise ToolError('Не удалось прочитать PDF. Возможно, файл повреждён.') from None
    try:
        if doc.needs_pass:
            raise ToolError('PDF защищён паролем. Загрузите копию без пароля.')
        if not 1 <= doc.page_count <= MAX_PAGES:
            raise ToolError('Поддерживаются документы от 1 до 500 страниц.')
        yield doc
    finally:
        doc.close()


def info(files):
    if not 1 <= len(files) <= MAX_FILES:
        raise ToolError('Выберите от 1 до 20 PDF-файлов.')
    result = []
    for name, data in files:
        with open_pdf(data) as doc:
            result.append({'name': name, 'pages': doc.page_count})
    if sum(item['pages'] for item in result) > MAX_PAGES:
        raise ToolError('Суммарно допускается до 500 страниц.')
    return result


def page_numbers(text, count, *, all_if_empty=False):
    text = str(text or '').strip().replace('–', '-').replace('—', '-')
    if not text:
        if all_if_empty:
            return list(range(count))
        raise ToolError('Укажите номера страниц, например: 1, 3-5.')
    if len(text) > 5000:
        raise ToolError('Слишком длинный список страниц.')
    pages, used = [], set()
    for part in text.split(','):
        match = re.fullmatch(r'\s*(\d{1,4})\s*(?:-\s*(\d{1,4})\s*)?', part)
        if not match:
            raise ToolError('Формат страниц: 1, 3-5. Для обратного порядка: 5-3, 2, 1.')
        start, end = int(match[1]), int(match[2] or match[1])
        if not (1 <= start <= count and 1 <= end <= count):
            raise ToolError(f'В документе {count} страниц. Допустимые номера: от 1 до {count}.')
        step = 1 if end >= start else -1
        for number in range(start, end + step, step):
            if number in used:
                raise ToolError(f'Страница {number} указана несколько раз.')
            used.add(number)
            pages.append(number - 1)
    return pages


def _copy_pages(source, indexes, output):
    # Copy contiguous runs together: this also preserves links within each run.
    start = end = indexes[0]
    for index in indexes[1:] + [None]:
        if index is not None and index == end + 1:
            end = index
            continue
        output.insert_pdf(source, from_page=start, to_page=end)
        start = end = index


def _clean_output(doc):
    # Keep page content, comments and ordinary links; remove executable PDF actions.
    doc.scrub(attached_files=True, embedded_files=True, javascript=True,
              metadata=False, xml_metadata=False, clean_pages=False, hidden_text=False,
              redactions=False, remove_links=False, reset_fields=False,
              reset_responses=False, thumbnails=False)


def process(files, action, pages, angle, job):
    if action not in ACTIONS:
        raise ToolError('Выберите действие со страницами PDF.')
    if not 1 <= len(files) <= MAX_FILES or (action == 'merge' and len(files) < 2):
        raise ToolError('Для объединения выберите от 2 до 20 PDF-файлов.')
    if action != 'merge' and len(files) != 1:
        raise ToolError('Для этого действия выберите один PDF-файл.')
    with ExitStack() as stack:
        docs = [stack.enter_context(open_pdf(data)) for _, data in files]
        if sum(doc.page_count for doc in docs) > MAX_PAGES:
            raise ToolError('Суммарно допускается до 500 страниц.')
        # Freeze filled form values into page content before rearranging documents.
        for doc in docs:
            if doc.is_form_pdf:
                doc.bake(annots=False, widgets=True)
        source = docs[0]
        selected = page_numbers(pages, source.page_count, all_if_empty=action in {'rotate', 'split'}) if action != 'merge' else []
        if action == 'delete':
            remove = set(selected)
            selected = [i for i in range(source.page_count) if i not in remove]
            if not selected:
                raise ToolError('Нельзя удалить все страницы. Оставьте хотя бы одну.')
        if action == 'reorder' and len(selected) != source.page_count:
            raise ToolError('Для изменения порядка укажите все страницы ровно по одному разу. Для части документа выберите «Извлечь страницы».')
        if action == 'split':
            path = job / 'pages.zip'
            with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
                for index in selected:
                    with fitz.open() as output:
                        _copy_pages(source, [index], output)
                        _clean_output(output)
                        archive.writestr(f'page-{index + 1:03d}.pdf', output.tobytes(garbage=4, deflate=True))
            return {'file': path.name, 'message': f'Готово: {len(selected)} отдельных PDF в ZIP-архиве.'}
        with fitz.open() as output:
            if action == 'merge':
                for doc in docs:
                    _copy_pages(doc, list(range(doc.page_count)), output)
            elif action == 'rotate':
                if str(angle) not in {'90', '180', '270'}:
                    raise ToolError('Выберите поворот на 90°, 180° или 270°.')
                _copy_pages(source, list(range(source.page_count)), output)
                for index in selected:
                    page = output[index]
                    page.set_rotation((page.rotation + int(angle)) % 360)
            else:
                _copy_pages(source, selected, output)
            _clean_output(output)
            path = job / 'result.pdf'
            output.save(path, garbage=4, deflate=True)
            return {'file': path.name, 'message': f'PDF готов. Страниц в результате: {output.page_count}.'}
