"""Editable PPTX export, rebuilt with python-pptx from the supplied visual template.

No bytecode from the executable is imported or executed. The checked-in template
contains four layout slides; all original example complaints have been removed.
"""
import copy
import html
import io
import math
import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from services.collective.client import HEADERS

TEMPLATE = Path(__file__).parent / 'assets' / 'template.pptx'
THEMES = {'ii': 'Инженерная инфраструктура', 'kr': 'Капитальный ремонт МКД', 'pr': 'Прочее'}
MAX_ROWS = 2000
MAX_CELL = 100000


def clean(value):
    text = str(value if value is not None else '')
    for _ in range(2):
        text = html.unescape(text)
    return re.sub(r'\s+', ' ', re.sub(r'<[^>]*>', ' ', text)).strip()


def cut(text, limit):
    text = clean(text)
    return text if len(text) <= limit else text[:limit].rsplit(' ', 1)[0] + '…'


def signature(value):
    match = re.match(r'^\+?\d+', re.sub(r'\s', '', str(value or '')))
    return min(int(match[0]), 2**53 - 1) if match else 0


def validate_rows(rows):
    if not isinstance(rows, list) or not rows:
        raise ValueError('Нет обращений для презентации.')
    if len(rows) > MAX_ROWS:
        raise ValueError(f'В презентацию можно включить до {MAX_ROWS} обращений. Уточните фильтры.')
    result = []
    for row in rows:
        if not isinstance(row, dict) or row.get('_theme') not in THEMES:
            raise ValueError('Некорректная тема обращения.')
        item = {}
        for key in HEADERS:
            value = row.get(key, '')
            if value is None:
                value = ''
            if isinstance(value, (list, dict, bool)) or len(str(value)) > MAX_CELL:
                raise ValueError('Некорректные поля обращения.')
            item[key] = str(value)
        item['_theme'] = row['_theme']
        item['_sig'] = signature(item[HEADERS[10]])
        result.append(item)
    return result


def text_cell(cell, text, *, bold=False, size=13, centered=False):
    cell.text = '\n'.join(clean(line) for line in str(text).splitlines())
    cell.margin_left = cell.margin_right = Inches(.09)
    cell.margin_top = cell.margin_bottom = Inches(.08)
    cell.vertical_anchor = MSO_ANCHOR.TOP
    cell.text_frame.word_wrap = True
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER if centered else PP_ALIGN.LEFT
        paragraph.space_after = Pt(2)
        paragraph.font.name = 'Arial'
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = RGBColor.from_string('172033')


def shape_text(shape, value, size):
    shape.text = value
    for p in shape.text_frame.paragraphs:
        p.font.name = 'Arial'
        p.font.size = Pt(size)
        p.font.color.rgb = RGBColor.from_string('172033')


def named(slide, name):
    return next(shape for shape in slide.shapes if shape.name == name)


def clone_slide(prs, source):
    slide = prs.slides.add_slide(source.slide_layout)
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)
    rels = {}
    for rid, rel in source.part.rels.items():
        if rel.reltype.endswith(('/slideLayout', '/notesSlide')):
            continue
        rels[rid] = slide.part.relate_to(rel.target_ref if rel.is_external else rel.target_part,
                                       rel.reltype, is_external=rel.is_external)
    for shape in source.shapes:
        element = copy.deepcopy(shape.element)
        for child in element.iter():
            for attr in ('r:embed', 'r:id', 'r:link'):
                old = child.get(qn(attr))
                if old in rels:
                    child.set(qn(attr), rels[old])
        slide.shapes._spTree.insert_element_before(element, 'p:extLst')
    return slide


def details(row):
    annotation = cut(row.get('Аннотация/краткое содержание') or row.get('Факт'), 480)
    source = clean(row.get('Источник обращения ЕЦУР') or row.get('Источник'))
    executor = clean(row.get('Ведомство исполнителя'))
    if 'министерство жилищно-коммунального' in executor.lower():
        executor = 'Министерство ЖКХ'
    elif 'администрац' in executor.lower():
        executor = 'ОМСУ'
    source = cut(' / '.join(x for x in (source, executor) if x), 100)
    status = clean(row.get('Статус')) or 'Без статуса'
    answer = cut(row.get('Ответ'), 780)
    return '\n'.join(x for x in (source, annotation) if x), '\n'.join(x for x in (status, answer) if x)


def row_height(row):
    content, status = details(row)
    # Conservative wrap estimate for 12 pt Arial and the fixed table columns.
    lines = max(sum(max(1, math.ceil(len(s) / 43)) for s in content.splitlines()),
                sum(max(1, math.ceil(len(s) / 66)) for s in status.splitlines()),
                math.ceil(len(clean(row.get('ОМСУ'))) / 15))
    return min(5.4, max(.65, .19 * lines + .22))


def fill_summary(slide, rows):
    totals = named(slide, 'collective_totals').table
    labels = ['Поступило', 'Более 100 подписей', 'До 100 подписей включительно', 'Повторных']
    values = [len(rows), sum(r['_sig'] > 100 for r in rows), sum(r['_sig'] <= 100 for r in rows),
              sum(clean(r['Повтор']).lower() == 'да' for r in rows)]
    text_cell(totals.cell(0, 0), 'Всего обращений', bold=True, size=18)
    for i, (label, value) in enumerate(zip(labels, values)):
        totals.columns[i].width = Inches(12.34 / 4)
        text_cell(totals.cell(1, i), label, bold=True, centered=True)
        text_cell(totals.cell(2, i), value, size=24, centered=True)
    for row, height in zip(totals.rows, (.48, .6, .68)):
        row.height = Inches(height)
    theme_shape = named(slide, 'collective_themes')
    theme_shape.top = Inches(3.05)
    table = theme_shape.table
    while len(table.rows) < 5:
        table._tbl.append(copy.deepcopy(table._tbl.findall(qn('a:tr'))[-1]))
    text_cell(table.cell(0, 0), 'Учитываемые обращения по темам', bold=True, size=18)
    for i, label in enumerate(['Категория', 'Общее количество', 'Более 100 подписей']):
        text_cell(table.cell(1, i), label, bold=True, centered=i > 0)
    active = [r for r in rows if 'не учитывается' not in r['Статус'].lower()]
    for i, (key, title) in enumerate(THEMES.items(), 2):
        group = [r for r in active if r['_theme'] == key]
        for j, value in enumerate([title, len(group), sum(r['_sig'] > 100 for r in group)]):
            text_cell(table.cell(i, j), value, bold=j == 0, size=15, centered=j > 0)
    for row in table.rows:
        row.height = Inches(.7)


def build(rows, start, end):
    prs = Presentation(TEMPLATE)
    cover, summary, sample, thanks = list(prs.slides)
    period = f'{start:%d.%m.%Y} — {end:%d.%m.%Y}'
    shape_text(named(cover, 'collective_period'), 'Свод ' + period, 28)
    shape_text(named(summary, 'collective_title'), 'Коллективные обращения · ' + period, 26)
    fill_summary(summary, rows)
    generated = []
    for theme, label in THEMES.items():
        group = sorted((r for r in rows if r['_theme'] == theme), key=lambda r: r['_sig'], reverse=True)
        chunks, chunk, used = [], [], 0
        for row in group:
            height = row_height(row)
            if chunk and (used + height > 5.65 or len(chunk) >= 6):
                chunks.append(chunk); chunk, used = [], 0
            chunk.append(row); used += height
        if chunk:
            chunks.append(chunk)
        for chunk in chunks:
            slide = clone_slide(prs, sample)
            generated.append(slide)
            title = slide.shapes.add_textbox(Inches(.16), Inches(.17), Inches(12.3), Inches(.55))
            shape_text(title, label + ' · ' + period, 23)
            table_shape = named(slide, 'collective_rows')
            table_shape.top = Inches(.85)
            table = table_shape.table
            for column, width in zip(table.columns, (1.45, .85, 4, 6.07)):
                column.width = Inches(width)
            body = copy.deepcopy(table._tbl.findall(qn('a:tr'))[1])
            for tr in list(table._tbl.findall(qn('a:tr')))[1:]:
                table._tbl.remove(tr)
            table.rows[0].height = Inches(.4)
            for cell in table.rows[0].cells:
                text_cell(cell, cell.text, bold=True, size=12)
            for row in chunk:
                table._tbl.append(copy.deepcopy(body))
                out = table.rows[len(table.rows) - 1]
                out.height = Inches(row_height(row))
                content, status = details(row)
                for i, text in enumerate([row['ОМСУ'] or '—', row['_sig'], content, status]):
                    text_cell(out.cells[i], text, size=12, centered=i == 1)
            table_shape.height = sum(row.height for row in table.rows)
    # Keep cover, summary, generated tables, and closing slide in that order.
    ids = {id(slide): element for slide, element in zip(prs.slides, list(prs.slides._sldIdLst))}
    sample_id = ids[id(sample)]
    prs.part.drop_rel(sample_id.rId)
    prs.slides._sldIdLst.remove(sample_id)
    order = [cover, summary, *generated, thanks]
    for element in list(prs.slides._sldIdLst):
        prs.slides._sldIdLst.remove(element)
    for slide in order:
        prs.slides._sldIdLst.append(ids[id(slide)])
    for i, slide in enumerate(order, 1):
        for shape in slide.shapes:
            if shape.name == 'collective_page':
                shape_text(shape, str(i), 12)
    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()
