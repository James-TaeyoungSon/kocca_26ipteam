import base64
import csv
import io
import json
import os
import sys
from typing import Any, Dict

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

NOTION_VERSION = "2025-09-03"


def load_event_payload(event_path: str) -> Dict[str, Any]:
    with open(event_path, "r", encoding="utf-8") as f:
        event = json.load(f)

    payload = event.get("client_payload") or event.get("inputs") or {}
    if not payload:
        raise RuntimeError("No payload found in GitHub event. Expected client_payload or inputs.")
    return payload


def unescape_csv_text(text: str) -> str:
    """Convert escaped JSON-string style CSV text back to raw CSV text."""
    return (
        text.replace("\\r\\n", "\n")
        .replace("\\n", "\n")
        .replace("\\r", "\n")
        .replace('\\\"', '"')
        .replace("\\\\", "\\")
    )


def build_xlsx_from_csv_text(csv_text: str, delimiter: str = ",", report_title: str = "") -> bytes:
    reader = csv.reader(io.StringIO(csv_text), delimiter=delimiter)
    rows = list(reader)
    if not rows:
        raise RuntimeError("CSV text is empty.")

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    num_cols = len(rows[0])
    data_start_row = 1

    # ── 1. Title row ──────────────────────────────────────────────
    if report_title:
        title_cell = ws.cell(row=1, column=1, value=report_title)
        if num_cols > 1:
            ws.merge_cells(
                start_row=1, start_column=1,
                end_row=1,   end_column=num_cols,
            )
        title_cell.font = Font(bold=True, size=14)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32
        data_start_row = 2

    # ── 2. Write CSV rows (header + data) ────────────────────────
    for row_idx, row in enumerate(rows):
        for col_idx, value in enumerate(row):
            ws.cell(row=data_start_row + row_idx, column=col_idx + 1, value=value)

    # ── 3. Style header row (bold, center) ───────────────────────
    header_row_idx = data_start_row
    for cell in ws[header_row_idx]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[header_row_idx].height = 20

    # ── 4. Borders on all data cells ─────────────────────────────
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows(
        min_row=data_start_row,
        max_row=data_start_row + len(rows) - 1,
        min_col=1,
        max_col=num_cols,
    ):
        for cell in row:
            cell.border = border

    # ── 5. Auto-fit column widths ─────────────────────────────────
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value is not None:
                v = str(cell.value)
                if len(v) > max_len:
                    max_len = len(v)
        ws.column_dimensions[col[0].column_letter].width = min(max(10, max_len + 2), 80)

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def notion_headers(token: str, with_json: bool = False) -> Dict[str, str]:
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": NOTION_VERSION,
    }
    if with_json:
        headers["Content-Type"] = "application/json"
    return headers


def upload_xlsx_to_notion(token: str, file_bytes: bytes, filename: str) -> str:
    create_resp = requests.post(
        "https://api.notion.com/v1/file_uploads",
        headers=notion_headers(token, with_json=True),
        json={
            "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "filename": filename,
        },
        timeout=30,
    )
    create_resp.raise_for_status()
    create_obj = create_resp.json()

    file_upload_id = create_obj["id"]
    upload_url = create_obj["upload_url"]

    send_resp = requests.post(
        upload_url,
        headers=notion_headers(token),
        files={"file": (filename, file_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        timeout=60,
    )
    send_resp.raise_for_status()
    return file_upload_id


def attach_file_to_page(token: str, page_id: str, file_upload_id: str, filename: str, file_property_name: str) -> None:
    body = {
        "properties": {
            file_property_name: {
                "files": [
                    {
                        "type": "file_upload",
                        "file_upload": {"id": file_upload_id},
                        "name": filename,
                    }
                ]
            }
        }
    }

    patch_resp = requests.patch(
        f"https://api.notion.com/v1/pages/{page_id}",
        headers=notion_headers(token, with_json=True),
        json=body,
        timeout=30,
    )
    patch_resp.raise_for_status()


def main() -> int:
    notion_token = os.getenv("NOTION_TOKEN", "").strip()
    if not notion_token:
        raise RuntimeError("NOTION_TOKEN secret is required.")

    event_path = os.getenv("GITHUB_EVENT_PATH", "")
    if not event_path:
        raise RuntimeError("GITHUB_EVENT_PATH is missing.")

    payload = load_event_payload(event_path)
    csv_text = payload.get("csv_text", "")
    csv_b64 = payload.get("csv_b64", "")
    page_id = payload.get("notion_page_id", "")
    filename = payload.get("report_name", "weekly_report.xlsx")
    file_property_name = payload.get("file_property_name", "파일과 미디어")
    delimiter = payload.get("delimiter", ",")
    report_title = payload.get("report_title", "")

    if not csv_text and not csv_b64:
        raise RuntimeError("payload.csv_text or payload.csv_b64 is required.")
    if not page_id:
        raise RuntimeError("payload.notion_page_id is required.")

    if csv_b64:
        csv_text = base64.b64decode(csv_b64).decode("utf-8", errors="replace")
    else:
        csv_text = unescape_csv_text(csv_text)

    if delimiter == "tab":
        delimiter = "\t"

    xlsx_bytes = build_xlsx_from_csv_text(csv_text, delimiter=delimiter, report_title=report_title)
    file_upload_id = upload_xlsx_to_notion(notion_token, xlsx_bytes, filename)
    attach_file_to_page(notion_token, page_id, file_upload_id, filename, file_property_name)

    print("OK")
    print(json.dumps({"page_id": page_id, "filename": filename, "file_upload_id": file_upload_id}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        raise
