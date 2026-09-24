"""Offline document conversion. No cloud services or source execution."""
import importlib.util
import io
import zipfile
from xml.etree import ElementTree as ET
from .workspace import ToolError


def validate_office(data):
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if len(archive.infolist()) > 3000 or sum(i.file_size for i in archive.infolist()) > 120 * 1024**2:
                raise ToolError('Слишком большой распакованный документ.')
            for info in archive.infolist():
                name = info.filename.lower()
                if 'vbaproject' in name or '/embeddings/' in name:
                    raise ToolError('Документы с макросами и вложенными программами не поддерживаются.')
                if name.endswith('.rels'):
                    for rel in ET.fromstring(archive.read(info)):
                        if rel.get('TargetMode') == 'External' and not rel.get('Type', '').endswith('/hyperlink'):
                            raise ToolError('В документе есть внешние ресурсы. Вставьте изображения в файл и повторите.')
                if name.startswith('word/') and name.endswith('.xml'):
                    tree = ET.fromstring(archive.read(info))
                    fields = ' '.join((e.text or '') for e in tree.iter() if e.tag.endswith('}instrText'))
                    import re
                    if re.search(r'\b(INCLUDETEXT|INCLUDEPICTURE|DDEAUTO|DDE|LINK)\b', fields, re.I):
                        raise ToolError('Удалите из документа поля, которые загружают внешние файлы.')
    except (zipfile.BadZipFile, ET.ParseError):
        raise ToolError('Файл Office повреждён или имеет другой формат.')


def convert(kind, name, data, job):
    from .pdf_pages import PDF_LOCK
    with PDF_LOCK:
        return _convert(kind, name, data, job)


def _convert(kind, name, data, job):
    if kind != 'pdf_docx':
        raise ToolError('Доступна только конвертация PDF → Word.')
    source_dir = job / 'input'
    source_dir.mkdir()
    source = source_dir / 'source.pdf'
    source.write_bytes(data)
    if not data.startswith(b'%PDF-'):
        raise ToolError('Файл не является PDF.')
    if not importlib.util.find_spec('pdf2docx'):
        raise ToolError('На сервере не установлен pdf2docx. Обновите зависимости проекта.')
    import pymupdf
    with pymupdf.open(source) as pdf:
        if pdf.is_encrypted:
            raise ToolError('Сначала снимите пароль с PDF.')
        if len(pdf) > 100:
            raise ToolError('Поддерживаются PDF до 100 страниц.')
        if not any(page.get_text().strip() for page in pdf):
            raise ToolError('В PDF только сканы. Для редактируемого Word сначала нужно распознавание текста.')
    from pdf2docx import Converter
    target = job / 'document.docx'
    converter = Converter(str(source))
    try:
        converter.convert(str(target), multi_processing=False)
    finally:
        converter.close()
    return {'file': target.name, 'message': 'Word готов. Проверьте перенос сложных таблиц и разметки.'}
