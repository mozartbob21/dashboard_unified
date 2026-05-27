from io import BytesIO


def extract_text_from_txt(file_bytes):
    return file_bytes.decode("utf-8", errors="ignore").strip()


def extract_text_from_docx(file_bytes):
    try:
        from docx import Document
    except Exception as exc:
        raise RuntimeError("Для чтения .docx нужен пакет python-docx") from exc

    document = Document(BytesIO(file_bytes))
    parts = []

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            parts.append(text)

    return "\n".join(parts).strip()


def extract_text_from_pdf(file_bytes):
    try:
        from pypdf import PdfReader
    except Exception:
        try:
            from PyPDF2 import PdfReader
        except Exception as exc:
            raise RuntimeError("Для чтения .pdf нужен пакет pypdf или PyPDF2") from exc

    reader = PdfReader(BytesIO(file_bytes))
    parts = []

    for page in reader.pages:
        text = page.extract_text() or ""
        text = text.strip()
        if text:
            parts.append(text)

    return "\n".join(parts).strip()


def extract_text_from_file(filename, file_bytes):
    filename = (filename or "").lower()

    if filename.endswith(".txt"):
        return extract_text_from_txt(file_bytes)

    if filename.endswith(".docx"):
        return extract_text_from_docx(file_bytes)

    if filename.endswith(".pdf"):
        return extract_text_from_pdf(file_bytes)

    raise ValueError("Поддерживаются только файлы .txt, .docx и .pdf")