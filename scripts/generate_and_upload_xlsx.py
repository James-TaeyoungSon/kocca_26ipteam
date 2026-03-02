import csv
import io
import json
import os
import sys
from typing import Dict, Any

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

NOTION_VERSION = "2025-09-03"


def load_event_payload(event_path: str) -> Dict[str, Any]:
    with open(event_path, "r", encoding="utf-8") as f:
        event = json.load(f)

    payload = event.get("client_payload") or event.get("inputs") or {}
    if not payload:
        raise RuntimeError("No payload found in GitHub event. Expected client_payload or inputs.")
    return payload


def build_xlsx_from_csv_text(csv_text: str) -> bytes:
    reader = csv.reader(io.StringIO(csv_text))
    rows = list(reader)
    if not rows:
        raise RuntimeError("CSV text is empty.")

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    for r in rows:
        ws.append(r)

    # Header style
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")

    # Simple column width sizing
    for col in ws.columns:
        max_len = 0
        for cell in col:
            v = "" if cell.value is None else str(cell.value)
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
    # 1) create file upload slot
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

    # 2) send file bytes
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
    page_id = payload.get("notion_page_id", "")
    filename = payload.get("report_name", "weekly_report.xlsx")
    file_property_name = payload.get("file_property_name", "파일과 미디어")

    if not csv_text:
        raise RuntimeError("payload.csv_text is required.")
    if not page_id:
        raise RuntimeError("payload.notion_page_id is required.")

    xlsx_bytes = build_xlsx_from_csv_text(csv_text)
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
