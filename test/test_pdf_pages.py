import io
import zipfile

import fitz
import pytest
from services.tools import pdf_pages as pdf
from services.tools.workspace import ToolError


def document(labels=('One', 'Two', 'Three'), rotation=0):
    with fitz.open() as doc:
        for label in labels:
            page = doc.new_page(width=300, height=400)
            page.insert_text((30, 50), label)
        doc[0].set_rotation(rotation)
        return doc.tobytes()


def result(job, action, pages='', angle='90', files=None):
    response = pdf.process(files or [('source.pdf', document())], action, pages, angle, job)
    return response, (job / response['file']).read_bytes()


def texts(data):
    with fitz.open(stream=data, filetype='pdf') as doc:
        return [page.get_text().strip() for page in doc]


def test_merge_uses_selected_file_order_and_keeps_originals(tmp_path):
    first, second = document(('One', 'Two')), document(('Three',))
    _, data = result(tmp_path, 'merge', files=[('same.pdf', second), ('same.pdf', first)])
    assert texts(data) == ['Three', 'One', 'Two']
    assert texts(first) == ['One', 'Two']


@pytest.mark.parametrize('action,pages,expected', [
    ('extract','3,1',['Three','One']), ('delete','2',['One','Three']),
    ('reorder','3-1',['Three','Two','One']), ('reorder','3,1-2',['Three','One','Two']),
])
def test_page_operations_preserve_text(tmp_path, action, pages, expected):
    _, data = result(tmp_path, action, pages)
    assert texts(data) == expected


def test_rotation_is_relative_and_only_affects_selected_pages(tmp_path):
    _, data = result(tmp_path, 'rotate', '1,3', files=[('original.pdf', document(rotation=90))])
    with fitz.open(stream=data) as doc:
        assert [page.rotation for page in doc] == [180,0,90]
    assert texts(data) == ['One','Two','Three']


def test_split_zip_has_one_pdf_per_selected_page(tmp_path):
    response, data = result(tmp_path, 'split', '3,1')
    assert response['file'].endswith('.zip')
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        assert z.namelist() == ['page-003.pdf','page-001.pdf']
        assert [texts(z.read(name)) for name in z.namelist()] == [['Three'],['One']]


def test_scanned_pdf_page_looks_the_same_after_extraction(tmp_path):
    from PIL import Image
    image = io.BytesIO(); Image.new('RGB',(120,80),'navy').save(image,format='PNG')
    with fitz.open() as source:
        page = source.new_page(width=300,height=400)
        page.insert_image(fitz.Rect(20,20,280,380),stream=image.getvalue())
        before = page.get_pixmap().samples
        data = source.tobytes()
    _, data = result(tmp_path,'extract','1',files=[('scan.pdf',data)])
    with fitz.open(stream=data) as output:
        assert output[0].get_pixmap().samples == before


def test_filled_form_values_survive_copying(tmp_path):
    with fitz.open() as source:
        page = source.new_page()
        widget = fitz.Widget(); widget.field_name='name'; widget.field_type=fitz.PDF_WIDGET_TYPE_TEXT
        widget.field_value='Filled value'; widget.rect=fitz.Rect(30,30,220,70)
        page.add_widget(widget)
        data = source.tobytes()
    _, data = result(tmp_path,'extract','1',files=[('form.pdf',data)])
    assert 'Filled value' in texts(data)[0]
    with fitz.open(stream=data) as output:
        assert not output.is_form_pdf


@pytest.mark.parametrize('action,pages,angle', [
    ('delete','1-3','90'), ('reorder','1,2','90'), ('extract','4','90'), ('extract','1,1','90'),
    ('extract','0','90'), ('extract','1-9999','90'), ('extract','bad','90'), ('extract','','90'),
    ('rotate','','45'), ('bad','','90')
])
def test_invalid_operations_never_publish_a_pdf(tmp_path, action, pages, angle):
    with pytest.raises(ToolError): result(tmp_path, action, pages, angle)
    assert not list(tmp_path.iterdir())


def test_password_invalid_file_and_page_limits():
    with fitz.open(stream=document()) as doc:
        locked = doc.tobytes(encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw='test-owner', user_pw='test-reader')
    for data in [locked, b'not a PDF', b'%PDF-invalid']:
        with pytest.raises(ToolError): pdf.info([('file.pdf',data)])
    with pytest.raises(ToolError): pdf.info([('file.pdf',document())]*21)
    with pytest.raises(ToolError): pdf.info([('file.pdf',document(['Page']*501))])


def test_info_returns_counts_without_retaining_documents(tmp_path):
    assert pdf.info([('one.pdf',document()),('two.pdf',document(['Last']))]) == [
        {'name':'one.pdf','pages':3},{'name':'two.pdf','pages':1}]
