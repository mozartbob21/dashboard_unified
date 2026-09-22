"""Offline document conversion. No cloud services or source execution."""
import importlib.util
import io
import os
import shutil
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from .workspace import ToolError


def office_binary():
    candidates = [os.getenv('LIBREOFFICE_PATH', ''), shutil.which('soffice'),
                  '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                  r'C:\Program Files\LibreOffice\program\soffice.com',
                  r'C:\Program Files (x86)\LibreOffice\program\soffice.com']
    return next((p for p in candidates if p and Path(p).is_file()), None)


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
    source_dir = job / 'input'
    source_dir.mkdir()
    source = source_dir / ('source.docx' if kind == 'docx_pdf' else 'source.pdf')
    source.write_bytes(data)
    if kind == 'docx_pdf':
        validate_office(data)
        binary = office_binary()
        if not binary:
            raise ToolError('На сервере нужен LibreOffice. Установите его или укажите LIBREOFFICE_PATH.')
        profile = source_dir / 'office-profile'
        profile.mkdir()
        (profile / 'user').mkdir()
        (profile / 'user/registrymodifications.xcu').write_text('''<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop><prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop></item><item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>2</value></prop></item></oor:items>''', encoding='utf-8')
        try:
            result = subprocess.run([binary, '-env:UserInstallation=' + profile.resolve().as_uri(),
                                     '--headless', '--nologo', '--nodefault', '--norestore',
                                     '--convert-to', 'pdf:writer_pdf_Export', '--outdir', str(job), str(source)],
                                    capture_output=True, timeout=120, shell=False)
        except subprocess.TimeoutExpired:
            raise ToolError('Документ обрабатывался дольше двух минут. Уменьшите его размер.')
        output = job / 'source.pdf'
        if result.returncode or not output.exists():
            raise ToolError('LibreOffice не смог преобразовать документ.')
        target = job / 'document.pdf'
        output.rename(target)
        return {'file': target.name, 'message': 'PDF готов.'}
    if not data.startswith(b'%PDF-'):
        raise ToolError('Файл не является PDF.')
    if not importlib.util.find_spec('pdf2docx'):
        raise ToolError('На сервере не установлен pdf2docx. Обновите зависимости проекта.')
    import fitz
    with fitz.open(source) as pdf:
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
