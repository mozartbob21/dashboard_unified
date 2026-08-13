"""HTML -> PPTX конвертер с поддержкой корпоративных шаблонов.

Шаблоны, у которых дизайн нарисован на самих слайдах, поддерживаются
через клонирование слайдов-доноров: титул берётся со слайда 1 шаблона,
контентные слайды — со слайда 2, последний слайд шаблона остаётся в конце.
"""
import base64
import copy
import io
import re
from pathlib import Path

from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

TOOLS_DIR = Path(__file__).resolve().parents[2] / "data" / "tools"
TEMPLATES_DIR = TOOLS_DIR / "templates"
OUTPUT_DIR = TOOLS_DIR / "output"


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def _extract_slide(container) -> dict:
    tags = container if isinstance(container, list) else container.find_all(True, recursive=False)
    slide = {"title": "", "subtitle": "", "bullets": [], "texts": [],
             "images": [], "tables": [], "notes": ""}
    for el in tags:
        name = getattr(el, "name", None)
        if not name:
            continue
        if name in ("h1", "h2", "h3") and not slide["title"]:
            slide["title"] = _clean(el.get_text())
        elif name in ("h2", "h3", "h4") and not slide["subtitle"]:
            slide["subtitle"] = _clean(el.get_text())
        elif name in ("ul", "ol"):
            slide["bullets"].extend(_clean(li.get_text()) for li in el.find_all("li"))
        elif name == "p":
            t = _clean(el.get_text())
            if t:
                slide["texts"].append(t)
        elif name == "blockquote":
            t = _clean(el.get_text())
            if t:
                slide["texts"].append("\u00ab" + t + "\u00bb")
        elif name == "table":
            rows = [[_clean(td.get_text()) for td in tr.find_all(["td", "th"])]
                    for tr in el.find_all("tr")]
            if rows:
                slide["tables"].append(rows)
        elif name in ("aside", "div") and "notes" in " ".join(el.get("class", [])):
            slide["notes"] = _clean(el.get_text())
    root = container if not isinstance(container, list) else None
    if root is not None:
        slide["images"] = [img.get("src", "") for img in root.find_all("img") if img.get("src")]
    return slide


def parse_html_to_slides(html: str) -> list:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style"]):
        tag.decompose()

    sections = soup.find_all("section")
    if sections:
        slides = [_extract_slide(s) for s in sections]
    else:
        body = soup.body or soup
        groups, cur = [], []
        for el in body.find_all(True, recursive=False):
            if getattr(el, "name", "") == "h2" and cur:
                groups.append(cur)
                cur = []
            cur.append(el)
        if cur:
            groups.append(cur)
        slides = [_extract_slide(g) for g in groups]

    slides = [s for s in slides
              if any([s["title"], s["bullets"], s["texts"], s["images"], s["tables"]])]
    if not slides:
        slides = [{"title": "Презентация", "subtitle": "", "bullets": [],
                   "texts": [_clean(soup.get_text())[:400]], "images": [],
                   "tables": [], "notes": ""}]
    return slides


def list_templates() -> list:
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    return sorted(p.name for p in TEMPLATES_DIR.glob("*.pptx"))


def _add_textbox(slide, lines):
    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(8.8), Inches(4.5))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = lines[0] if lines else ""
    for ln in lines[1:]:
        tf.add_paragraph().text = ln


def _add_image(slide, src):
    try:
        if src.startswith("data:"):
            data = base64.b64decode(src.split(",", 1)[1])
            slide.shapes.add_picture(io.BytesIO(data), Inches(1.2), Inches(2.2), width=Inches(7.5))
        elif src.startswith("http"):
            import httpx
            r = httpx.get(src, timeout=10)
            if r.status_code == 200:
                slide.shapes.add_picture(io.BytesIO(r.content), Inches(1.2), Inches(2.2), width=Inches(7.5))
        else:
            p = Path(src)
            if p.exists():
                slide.shapes.add_picture(str(p), Inches(1.2), Inches(2.2), width=Inches(7.5))
    except Exception as e:
        print(f"[tools] image error: {e}")


def _add_table(slide, rows):
    try:
        n_rows, n_cols = len(rows), max(len(r) for r in rows)
        g = slide.shapes.add_table(n_rows, n_cols, Inches(0.5), Inches(2.0),
                                   Inches(9), Inches(0.4 * n_rows))
        tbl = g.table
        for r_i, row in enumerate(rows):
            for c_i in range(n_cols):
                tbl.cell(r_i, c_i).text = row[c_i] if c_i < len(row) else ""
    except Exception as e:
        print(f"[tools] table error: {e}")


def duplicate_slide(prs, index):
    """Клонирует слайд шаблона вместе с фигурами и картинками (фон, логотип)."""
    source = prs.slides[index]
    new = prs.slides.add_slide(prs.slide_layouts[len(prs.slide_layouts) - 1])
    spTree = new.shapes._spTree
    for shp in list(new.shapes):
        spTree.remove(shp._element)
    for shp in source.shapes:
        el = copy.deepcopy(shp._element)
        if shp.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE):
            for b in el.iter(qn("a:blip")):
                rId = b.get(qn("r:embed"))
                if rId:
                    try:
                        img_part = source.part.related_part(rId)
                        new_rId = new.part.relate_to(img_part, RT.IMAGE)
                        b.set(qn("r:embed"), new_rId)
                    except Exception:
                        pass
        spTree.append(el)
    return new


def _style(tf, size, bold, color, align):
    for p in tf.paragraphs:
        p.alignment = align
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color


def _fill_slide(slide, s, is_first):
    """Клон остаётся как фон: удаляем демо-таблицы и рамки с текстом,
    свой контент добавляем аккуратными блоками."""
    spTree = slide.shapes._spTree
    for shp in list(slide.shapes):
        if shp.has_table:
            spTree.remove(shp._element)
            continue
        if shp.has_text_frame and shp.text_frame.text.strip():
            spTree.remove(shp._element)

    # Дочистка: узкие вертикальные рамки и мелкие значки у правого края
    for shp in list(slide.shapes):
        try:
            l, w, h = shp.left, shp.width, shp.height
        except Exception:
            continue
        if shp.shape_type == MSO_SHAPE_TYPE.TEXT_BOX and w < Inches(2):
            spTree.remove(shp._element)
            continue
        if w < Inches(1.2) and h < Inches(2) and l > Inches(7.5):
            spTree.remove(shp._element)
            continue

    if is_first:
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), Inches(8.4), Inches(1.6))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.text = s["title"] or "Презентация"
        _style(tf, 32, True, RGBColor(0xFF, 0xFF, 0xFF), PP_ALIGN.CENTER)
        sub = s["subtitle"] or " ".join(s["texts"])
        if sub:
            tb2 = slide.shapes.add_textbox(Inches(1.2), Inches(4.3), Inches(7.6), Inches(1.0))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            tf2.text = sub
            _style(tf2, 18, False, RGBColor(0xFF, 0xFF, 0xFF), PP_ALIGN.CENTER)
    else:
        # Фирменная зелёная плашка сверху (как в шаблоне)
        bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.3), Inches(0.25), Inches(9.4), Inches(0.75))
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(0x2E, 0x7D, 0x32)
        bar.line.fill.background()
        try:
            bar.shadow.inherit = False
        except Exception:
            pass

        tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.15), Inches(8.8), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.text = s["title"] or ""
        _style(tf, 26, True, RGBColor(0x1E, 0x29, 0x3B), PP_ALIGN.LEFT)
        body = s["bullets"] or s["texts"]
        if body:
            tb2 = slide.shapes.add_textbox(Inches(0.7), Inches(2.2), Inches(8.6), Inches(4.3))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            tf2.text = body[0]
            for ln in body[1:]:
                tf2.add_paragraph().text = ln
            _style(tf2, 18, False, RGBColor(0x33, 0x33, 0x33), PP_ALIGN.LEFT)

    for src in s["images"][:2]:
        _add_image(slide, src)
    for rows in s["tables"][:1]:
        _add_table(slide, rows)
    if s["notes"]:
        try:
            slide.notes_slide.notes_text_frame.text = s["notes"]
        except Exception:
            pass



def _build_standard(prs, slides):
    """Стандартное оформление (без шаблона)."""
    for i, s in enumerate(slides):
        layout_idx = 0 if i == 0 else (1 if s["bullets"] else 5)
        try:
            layout = prs.slide_layouts[layout_idx]
        except IndexError:
            layout = prs.slide_layouts[min(layout_idx, len(prs.slide_layouts) - 1)]
        slide = prs.slides.add_slide(layout)
        if slide.shapes.title and s["title"]:
            slide.shapes.title.text = s["title"]
        if i == 0 and (s["subtitle"] or s["texts"]):
            for ph in slide.placeholders:
                if ph.placeholder_format.idx == 1:
                    ph.text = s["subtitle"] or " ".join(s["texts"])
                    break
        if i != 0 and (s["bullets"] or s["texts"]):
            body_ph = None
            for ph in slide.placeholders:
                if ph.placeholder_format.idx == 1:
                    body_ph = ph
            lines = s["bullets"] or s["texts"]
            if body_ph is not None:
                tf = body_ph.text_frame
                tf.text = lines[0]
                for ln in lines[1:]:
                    tf.add_paragraph().text = ln
            else:
                _add_textbox(slide, lines)
        for src in s["images"][:2]:
            _add_image(slide, src)
        for rows in s["tables"][:1]:
            _add_table(slide, rows)
        if s["notes"]:
            try:
                slide.notes_slide.notes_text_frame.text = s["notes"]
            except Exception:
                pass


def build_pptx(slides: list, template_name, output_name: str) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    tpl = TEMPLATES_DIR / template_name if template_name else None
    use_tpl = bool(tpl and tpl.exists())
    print(f"[tools] build_pptx: template={template_name!r}, found={use_tpl}", flush=True)

    if not use_tpl:
        prs = Presentation()
        _build_standard(prs, slides)
    else:
        prs = Presentation(str(tpl))
        n_tpl = len(prs.slides._sldIdLst)
        if n_tpl == 0:
            _build_standard(prs, slides)
        else:
            title_donor = 0
            content_donor = 1 if n_tpl > 2 else 0
            for i, s in enumerate(slides):
                donor = title_donor if i == 0 else content_donor
                new_slide = duplicate_slide(prs, donor)
                _fill_slide(new_slide, s, i == 0)
            # Удаляем демо-слайды шаблона, кроме последнего ("Спасибо"),
            # и переносим его в конец
            sldIdLst = prs.slides._sldIdLst
            orig = list(sldIdLst)[:n_tpl]
            for el in orig[:-1]:
                rId = el.get(qn("r:id"))
                try:
                    prs.part.drop_rel(rId)
                except Exception:
                    pass
                sldIdLst.remove(el)
            last = orig[-1]
            sldIdLst.remove(last)
            sldIdLst.append(last)

    out = OUTPUT_DIR / output_name
    prs.save(str(out))
    return out
