"""Source fidelity: keep query ordering and the full text used by K7 rules."""
import io
from datetime import datetime

import httpx
from openpyxl import Workbook
from openpyxl.utils.datetime import to_excel

from services.omsu_rating import calculate, source


def test_saved_redmine_query_order_preserves_equal_value_order():
    def handler(request):
        assert "sort" not in request.url.params
        assert request.url.params["query_id"] == "52"
        return httpx.Response(200, json={
            "total_count": 2,
            "issues": [
                {"id": ident, "assigned_to": {"name": name}, "custom_fields": [
                    {"id": 17, "value": "1"}, {"id": 158, "value": "5"}]}
                for ident, name in [(20, "ХИМКИ"), (10, "ДУБНА")]
            ],
        })

    with httpx.Client(transport=httpx.MockTransport(handler)) as client:
        issues = source.fetch_redmine(52, "synthetic-key", client)
    rows = calculate.rating(5, 0, issues)
    assert [row["omsu"] for row in rows] == ["ХИМКИ", "ДУБНА"]
    assert [row["place"] for row in rows] == [1, 1]


def test_full_sheet_cell_reaches_classification_and_explanation():
    book = Workbook()
    sheet = book.active
    sheet.title = "Исполнение дорожной карты"
    sheet.append(["header"])
    text = "Подробное пояснение. " * 80 + " выполнено"
    sheet.append([None, None, "МУП «БКС»", "БАЛАШИХА", text])
    content = io.BytesIO()
    book.save(content)
    with httpx.Client(transport=httpx.MockTransport(
            lambda _: httpx.Response(200, content=content.getvalue()))) as client:
        source_rows = source.fetch_sheet(client)
    assert source_rows[0][4] == text
    rows, _ = calculate.rating7(source_rows)
    rso = next(row for row in rows if row["omsu"] == "БАЛАШИХА")["details"][0]
    assert rso["items"][0]["ok"] is True
    assert rso["items"][0]["raw"] == text


def test_date_formatted_progress_cells_keep_original_online_raw_number():
    book = Workbook()
    sheet = book.active
    sheet.title = "Исполнение дорожной карты"
    sheet.append(["header"])
    value = datetime(2026, 9, 30)
    sheet.append([None, None, "МУП «БКС»", "БАЛАШИХА", value])
    output = io.BytesIO()
    book.save(output)
    with httpx.Client(transport=httpx.MockTransport(
            lambda _: httpx.Response(200, content=output.getvalue()))) as client:
        source_rows = source.fetch_sheet(client)
    assert source_rows[0][4] == to_excel(value, book.epoch)
    rows, _ = calculate.rating7(source_rows)
    rso = next(row for row in rows if row['omsu'] == 'БАЛАШИХА')['details'][0]
    assert rso['items'][0]['ok'] is True
