"""Извлечение текста из файлов: txt, md, csv, docx, pdf, xlsx, xls, pptx."""
import io
import re
import zipfile
from xml.etree import ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def extract_txt(data: bytes) -> str:
    return data.decode("utf-8", errors="ignore").strip()


def extract_xlsx(data: bytes) -> str:
    """xlsx через zip+xml (стандартная библиотека)."""
    z = zipfile.ZipFile(io.BytesIO(data))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(f"{NS}si"):
            shared.append("".join(t.text or "" for t in si.iter(f"{NS}t")))
    out = []
    sheets = sorted([n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", n)])
    for sn in sheets[:5]:
        root = ET.fromstring(z.read(sn))
        for row in root.iter(f"{NS}row"):
            cells = []
            for c in row.findall(f"{NS}c"):
                v = c.find(f"{NS}v")
                t = c.get("t")
                if v is None:
                    cells.append("")
                elif t == "s":
                    i = int(v.text or 0)
                    cells.append(shared[i] if i < len(shared) else "")
                else:
                    cells.append(v.text or "")
            line = " | ".join(x for x in cells if x)
            if line:
                out.append(line)
    return "\n".join(out)


def extract_xls(data: bytes) -> str:
    """Старый бинарный .xls — пробуем xlrd, иначе просим сохранить как .xlsx."""
    try:
        import xlrd
        wb = xlrd.open_workbook(file_contents=data)
        out = []
        for sh in wb.sheets()[:5]:
            for r in range(sh.nrows):
                out.append(" | ".join(str(c) for c in sh.row_values(r) if str(c).strip()))
        return "\n".join(out)
    except ImportError:
        return ("[файл .xls (старый формат). Сохраните его как .xlsx — "
                "и я прочитаю содержимое.]")


def extract_pptx(data: bytes) -> str:
    """pptx через zip+xml: текст всех слайдов по порядку."""
    z = zipfile.ZipFile(io.BytesIO(data))
    slides = sorted(
        [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml", n)],
        key=lambda s: int(re.search(r"(\d+)", s).group(1)))
    out = []
    for i, sn in enumerate(slides, 1):
        root = ET.fromstring(z.read(sn))
        texts = [n.text.strip() for n in root.iter()
                 if n.tag.endswith("}t") and n.text and n.text.strip()]
        if texts:
            out.append(f"── Слайд {i} ── " + " ".join(texts))
    return "\n".join(out)


def extract_docx(data: bytes) -> str:
    try:
        from services.appeals.files import extract_text_from_docx
        return extract_text_from_docx(data)
    except Exception:
        z = zipfile.ZipFile(io.BytesIO(data))
        root = ET.fromstring(z.read("word/document.xml"))
        return "\n".join(p.text or "" for p in root.iter() if p.tag.endswith("}t"))


def extract_pdf(data: bytes) -> str:
    from services.appeals.files import extract_text_from_pdf
    return extract_text_from_pdf(data)


def extract_any(name: str, data: bytes) -> str:
    low = (name or "").lower()
    try:
        if low.endswith((".webm", ".mp3", ".wav", ".m4a", ".ogg", ".opus")):
            from services.aichat.engine import transcribe
            ext = low.rsplit(".", 1)[-1]
            mime = {"webm": "audio/webm", "mp3": "audio/mpeg", "wav": "audio/wav",
                    "m4a": "audio/mp4", "ogg": "audio/ogg",
                    "opus": "audio/opus"}.get(ext, "audio/webm")
            t = transcribe(data, mime)
            return f"🎤 Распознанная диктовка: {t}" if t else "[не удалось распознать аудио]"
        if low.endswith((".xlsx",)):
            return extract_xlsx(data) or "[пустая книга]"
        if low.endswith((".xls",)):
            return extract_xls(data)
        if low.endswith((".pptx",)):
            return extract_pptx(data) or "[презентация без текста]"
        if low.endswith((".docx",)):
            return extract_docx(data)
        if low.endswith((".pdf",)):
            return extract_pdf(data)
        return extract_txt(data)
    except Exception as e:
        return f"[не удалось извлечь текст из {name}: {e}]"
