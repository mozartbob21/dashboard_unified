"""Bounded existing file extraction for chat; no downloaded resources or execution."""
import io
from pathlib import PurePath
from zipfile import ZipFile, is_zipfile


MAX_EXPANDED_BYTES = 64 * 1024 * 1024
XML_CHUNK_SIZE = 64 * 1024


class UploadTextError(ValueError):
    pass


def extract_bounded(name, data):
    if len(data) > 8 * 1024 * 1024:
        raise UploadTextError('Файл превышает допустимый размер.')
    if PurePath(name).suffix.lower() in {'.png', '.jpg', '.jpeg', '.gif', '.webp'}:
        raise UploadTextError('Для изображения нужно приложить распознанный текст или документ.')
    if PurePath(name).suffix.lower() in {'.webm', '.mp3', '.wav', '.m4a', '.ogg', '.opus'}:
        raise UploadTextError('Для аудио нужно приложить текстовую расшифровку.')
    from services.aichat.extract import extract_any
    try:
        if is_zipfile(io.BytesIO(data)):
            with ZipFile(io.BytesIO(data)) as archive:
                entries = archive.infolist()
                if len(entries) > 2000 or sum(item.file_size for item in entries) > MAX_EXPANDED_BYTES:
                    raise UploadTextError('Слишком большой объём распакованных данных. Приложите меньший файл.')
                expanded_bytes = 0
                for item in entries:
                    if item.is_dir():
                        continue
                    # Office relationships can name XML parts without an .xml suffix.
                    # Check every part before any extractor can parse it.
                    # UTF-16/32 declarations and chunk boundaries must not hide entities.
                    with archive.open(item) as source:
                        tail = b''
                        while chunk := source.read(XML_CHUNK_SIZE):
                            expanded_bytes += len(chunk)
                            if expanded_bytes > MAX_EXPANDED_BYTES:
                                raise UploadTextError('Слишком большой объём распакованных данных. Приложите меньший файл.')
                            declarations = tail + chunk.replace(b'\x00', b'').upper()
                            if b'<!DOCTYPE' in declarations or b'<!ENTITY' in declarations:
                                raise UploadTextError('Документ содержит неподдерживаемые объявления XML. Пересохраните его без них.')
                            tail = declarations[-16:]
        result = extract_any(name, data)
        if result.startswith('[не удалось извлечь текст из'):
            raise UploadTextError('Не удалось прочитать файл. Проверьте формат и приложите его заново.') from None
        return result[:60000] + ('\n[Показаны первые 60 000 символов.]' if len(result) > 60000 else '')
    except UploadTextError:
        raise
    except Exception:
        raise UploadTextError('Не удалось прочитать файл. Проверьте формат и приложите его заново.') from None
