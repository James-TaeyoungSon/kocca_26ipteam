# KOCCA Notion XLSX Uploader

This repository contains a GitHub Actions workflow that receives CSV text from Make.com,
converts it into an `.xlsx` file, and uploads the file to a Notion page's `파일과 미디어` property.

## Required GitHub Secret

- `NOTION_TOKEN`: Internal integration token with access to the target Notion page/database.

## Workflow Trigger

The workflow listens to `repository_dispatch` with event type:

- `make_notion_weekly_report`

Expected `client_payload` fields:

- `csv_text` (string)
- `notion_page_id` (string)
- `report_name` (string, optional)
- `file_property_name` (string, optional, default: `파일과 미디어`)

## Make.com Notes

In your Make HTTP module, call:

- `POST https://api.github.com/repos/James-TaeyoungSon/kocca_26ipteam/dispatches`

Headers:

- `Authorization: Bearer <GITHUB_PAT_WITH_repo_scope>`
- `Accept: application/vnd.github+json`
- `X-GitHub-Api-Version: 2022-11-28`
- `Content-Type: application/json`

Body format:

```json
{
  "event_type": "make_notion_weekly_report",
  "client_payload": {
    "csv_text": "...",
    "notion_page_id": "...",
    "report_name": "weekly_report.xlsx",
    "file_property_name": "파일과 미디어"
  }
}
```
