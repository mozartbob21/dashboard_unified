from pathlib import Path
from datetime import datetime
from zipfile import ZipFile
from docx import Document

BASE_DIR = Path(__file__).resolve().parent.parent
GENERATED_DIR = BASE_DIR / "generated" / "prescriptions"
GENERATED_DIR.mkdir(parents=True, exist_ok=True)


def sanitize_filename(value: str) -> str:
    bad_chars = '\\/:*?"<>|'
    for ch in bad_chars:
        value = value.replace(ch, "_")
    return value.strip().replace(" ", "_")


def build_demo_document(doc: Document, owner: str, address: str, checked_at: str | None = None):
    date_str = datetime.now().strftime("%d.%m.%Y")

    safe_owner = owner or "Неизвестная организация"
    safe_address = address or "Адрес не указан"
    safe_checked_at = checked_at or "—"

    doc.add_heading("ПРЕДПИСАНИЕ", level=0)
    doc.add_paragraph(f"Дата формирования: {date_str}")
    doc.add_paragraph(f"Организация: {safe_owner}")
    doc.add_paragraph(f"Адрес объекта: {safe_address}")
    doc.add_paragraph(f"Время проверки: {safe_checked_at}")
    doc.add_paragraph("")

    doc.add_paragraph(
        "В ходе автоматической проверки системы видеонаблюдения выявлено, "
        "что камера по указанному адресу не подключена либо не работает."
    )

    doc.add_paragraph(
        "Необходимо устранить выявленное нарушение и обеспечить работоспособность "
        "камеры видеонаблюдения."
    )

    doc.add_paragraph("")
    doc.add_paragraph("Демонстрационный шаблон документа.")
    doc.add_paragraph("В дальнейшем текст будет заменен на шаблон из JSON.")


def generate_demo_prescription(address: str, owner: str, checked_at: str | None = None) -> str:
    filename = f"prescription_{sanitize_filename(owner)}_{sanitize_filename(address)}.docx"
    file_path = GENERATED_DIR / filename

    doc = Document()
    build_demo_document(doc, owner=owner, address=address, checked_at=checked_at)
    doc.save(file_path)

    return str(file_path)


def generate_individual_prescriptions(items: list[dict]) -> list[dict]:
    generated_items = []

    for item in items:
        owner = item.get("owner", "") or "Не указана организация"
        address = item.get("address", "") or "Адрес не указан"
        checked_at = item.get("checked_at", "") or "—"
        camera_status = item.get("camera_status", "")

        file_path = generate_demo_prescription(
            address=address,
            owner=owner,
            checked_at=checked_at,
        )

        generated_items.append(
            {
                "owner": owner,
                "address": address,
                "checked_at": checked_at,
                "camera_status": camera_status,
                "filename": Path(file_path).name,
            }
        )

    return generated_items


def generate_combined_prescription(items: list[dict]) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"combined_prescriptions_{timestamp}.docx"
    file_path = GENERATED_DIR / filename

    doc = Document()
    doc.add_heading("СВОДНОЕ ПРЕДПИСАНИЕ", level=0)
    doc.add_paragraph(f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    doc.add_paragraph("")

    if not items:
        doc.add_paragraph("Проблемных адресов не найдено.")
    else:
        for index, item in enumerate(items, start=1):
            owner = item.get("owner", "") or "Не указана организация"
            address = item.get("address", "") or "Адрес не указан"
            checked_at = item.get("checked_at", "") or "—"

            doc.add_heading(f"{index}. {owner}", level=1)
            doc.add_paragraph(f"Адрес объекта: {address}")
            doc.add_paragraph(f"Время проверки: {checked_at}")
            doc.add_paragraph(
                "В ходе автоматической проверки системы видеонаблюдения выявлено, "
                "что камера по указанному адресу не подключена либо не работает."
            )
            doc.add_paragraph(
                "Необходимо устранить выявленное нарушение и обеспечить работоспособность "
                "камеры видеонаблюдения."
            )
            doc.add_paragraph("")

    doc.save(file_path)
    return str(file_path)


def generate_zip_with_prescriptions(items: list[dict]) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_filename = f"prescriptions_{timestamp}.zip"
    zip_path = GENERATED_DIR / zip_filename

    generated_files = []

    for item in items:
        owner = item.get("owner", "") or "Не указана организация"
        address = item.get("address", "") or "Адрес не указан"
        checked_at = item.get("checked_at", "") or "—"

        file_path = generate_demo_prescription(
            address=address,
            owner=owner,
            checked_at=checked_at,
        )
        generated_files.append(Path(file_path))

    with ZipFile(zip_path, "w") as zip_file:
        for file_path in generated_files:
            zip_file.write(file_path, arcname=file_path.name)

    return str(zip_path)
