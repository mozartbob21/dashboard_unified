"""Offline upload checks with synthetic bytes; no model, network or real files."""
import io
import unittest
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from services.aichat import extract, uploads


def archive_bytes(entries):
    output = io.BytesIO()
    with ZipFile(output, 'w', compression=ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)
    return output.getvalue()


class ChatUploadTests(unittest.TestCase):
    def test_zip_probe_errors_are_sanitized_before_extraction(self):
        with patch.object(uploads, 'is_zipfile', side_effect=ValueError('PRIVATE_ZIP_LOCATION')):
            with patch.object(extract, 'extract_any') as extractor:
                with self.assertRaises(uploads.UploadTextError) as caught:
                    uploads.extract_bounded('broken.docx', b'PK\x03\x04broken')
        self.assertNotIn('PRIVATE_ZIP_LOCATION', str(caught.exception))
        self.assertEqual(str(caught.exception), 'Не удалось прочитать файл. Проверьте формат и приложите его заново.')
        extractor.assert_not_called()

    def test_malformed_office_zip_does_not_return_raw_extractor_error(self):
        with self.assertRaises(uploads.UploadTextError) as caught:
            uploads.extract_bounded('private-document.xlsx', b'PK\x03\x04broken')
        self.assertNotIn('private-document', str(caught.exception))
        self.assertNotIn('BadZipFile', str(caught.exception))

    def test_expanded_zip_bound_rejects_compressed_content_before_extraction(self):
        data = archive_bytes({'word/document.xml': b'x' * 4097})
        self.assertLess(len(data), 4096)
        with patch.object(uploads, 'MAX_EXPANDED_BYTES', 4096):
            with patch.object(extract, 'extract_any') as extractor:
                with self.assertRaisesRegex(uploads.UploadTextError, 'объём распакованных данных'):
                    uploads.extract_bounded('large.docx', data)
        extractor.assert_not_called()

    def test_office_xml_entities_are_rejected_for_unicode_and_chunk_boundaries(self):
        body = '<!DOCTYPE root [<!ENTITY sample "expanded">]><root>&sample;</root>'
        for extension in ['docx', 'pptx', 'xlsx']:
            for encoding in ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'utf-32-le', 'utf-32-be']:
                for part in ['customXml/unread.XML', '_rels/.rels', 'word/alternate.bin']:
                    data = archive_bytes({part: body.encode(encoding)})
                    with self.subTest(extension=extension, encoding=encoding, part=part):
                        # Deliberately split both declarations and UTF code units.
                        with patch.object(uploads, 'XML_CHUNK_SIZE', 7):
                            with patch.object(extract, 'extract_any') as extractor:
                                with self.assertRaisesRegex(uploads.UploadTextError, 'объявления XML'):
                                    uploads.extract_bounded('document.' + extension, data)
                        extractor.assert_not_called()

    def test_normal_unicode_office_xml_is_passed_to_existing_extractor(self):
        data = archive_bytes({'word/document.xml': '<?xml version="1.0" encoding="UTF-16"?><root>Текст</root>'.encode('utf-16'),
                              '_rels/.rels': b'<Relationships/>'})
        with patch.object(extract, 'extract_any', return_value='Текст документа') as extractor:
            self.assertEqual(uploads.extract_bounded('document.docx', data), 'Текст документа')
        extractor.assert_called_once_with('document.docx', data)

    def test_images_and_audio_never_enter_extractor_or_transcription(self):
        for extension in ['png', 'jpg', 'jpeg', 'gif', 'webp', 'webm', 'mp3', 'wav', 'm4a', 'ogg', 'opus']:
            with self.subTest(extension=extension), patch.object(extract, 'extract_any') as extractor:
                with self.assertRaises(uploads.UploadTextError):
                    uploads.extract_bounded('media.' + extension.upper(), b'unsupported media')
                # extract_any is the only entry to audio transcription.
                extractor.assert_not_called()

    def test_valid_utf8_text_and_extraction_limit(self):
        self.assertEqual(uploads.extract_bounded('note.txt', 'Простой текст'.encode()), 'Простой текст')
        result = uploads.extract_bounded('long.txt', b'x' * 60001)
        self.assertEqual(result, 'x' * 60000 + '\n[Показаны первые 60 000 символов.]')


if __name__ == '__main__':
    unittest.main()
