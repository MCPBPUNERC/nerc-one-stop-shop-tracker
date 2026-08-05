import json
import os
import smtplib
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from openpyxl import load_workbook

EXCEL_URL = "https://www.nerc.com/globalassets/align-reports/one-stop-shop.xlsx"
SNAPSHOT_PATH = Path("last_snapshot.xlsx")
DATA_PATH = Path("docs/data/nerc.json")
HISTORY_PATH = Path("docs/data/nerc-history.json")
MAX_WEB_CHANGES = 500
MAX_EMAIL_CHANGES = 40

GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASS = os.environ.get("GMAIL_PASS")
RECIPIENTS = [x.strip() for x in os.environ.get(
    "RECIPIENTS", "mikep@mcphersonpower.com,tommys@mcphersonpower.com"
).split(",") if x.strip()]


def download_excel():
    response = requests.get(EXCEL_URL, timeout=90, headers={"User-Agent": "BPU-NERC-Tracker/1.0"})
    response.raise_for_status()
    return response.content


def value_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat()
    return str(value).strip()


def read_workbook(path):
    workbook = load_workbook(path, read_only=True, data_only=False)
    sheets = {}
    metadata = []
    total_cells = 0
    for sheet in workbook.worksheets:
        values = {}
        populated = 0
        for row in sheet.iter_rows():
            for cell in row:
                text = value_text(cell.value)
                if text != "":
                    values[cell.coordinate] = text
                    populated += 1
        sheets[sheet.title] = values
        total_cells += populated
        metadata.append({
            "name": sheet.title,
            "rows": sheet.max_row,
            "columns": sheet.max_column,
            "cells": populated,
            "changes": 0,
        })
    workbook.close()
    return sheets, metadata, total_cells


def compare_workbooks(old_path, new_path):
    old_sheets, _, _ = read_workbook(old_path)
    new_sheets, sheet_meta, total_cells = read_workbook(new_path)
    changes = []
    counts = {"added": 0, "removed": 0, "modified": 0}
    per_sheet = {item["name"]: 0 for item in sheet_meta}

    for sheet_name in sorted(set(old_sheets) | set(new_sheets)):
        old_cells = old_sheets.get(sheet_name, {})
        new_cells = new_sheets.get(sheet_name, {})
        for coordinate in sorted(set(old_cells) | set(new_cells)):
            old = old_cells.get(coordinate, "")
            new = new_cells.get(coordinate, "")
            if old == new:
                continue
            if not old and new:
                change_type = "added"
            elif old and not new:
                change_type = "removed"
            else:
                change_type = "modified"
            counts[change_type] += 1
            per_sheet[sheet_name] = per_sheet.get(sheet_name, 0) + 1
            if len(changes) < MAX_WEB_CHANGES:
                changes.append({"sheet": sheet_name, "cell": coordinate, "old": old, "new": new, "type": change_type})

    for item in sheet_meta:
        item["changes"] = per_sheet.get(item["name"], 0)
    total_changes = sum(counts.values())
    return changes, counts, sheet_meta, total_cells, total_changes


def load_history():
    try:
        return json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def write_dashboard(payload, history):
    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    DATA_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    HISTORY_PATH.write_text(json.dumps(history[:90], indent=2, ensure_ascii=False), encoding="utf-8")


def send_email(subject, body):
    if not GMAIL_USER or not GMAIL_PASS:
        raise RuntimeError("GMAIL_USER or GMAIL_PASS is not set.")
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = GMAIL_USER
    message["To"] = ", ".join(RECIPIENTS)
    message.set_content(body)
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as server:
        server.starttls()
        server.login(GMAIL_USER, GMAIL_PASS)
        server.send_message(message)


def main():
    now_utc = datetime.now(timezone.utc)
    now_ct = now_utc.astimezone(ZoneInfo("America/Chicago"))
    today = now_ct.strftime("%Y-%m-%d")
    new_bytes = download_excel()
    temp_path = Path("current_snapshot.xlsx")
    temp_path.write_bytes(new_bytes)

    history = load_history()
    if SNAPSHOT_PATH.exists():
        changes, counts, sheets, cells, total = compare_workbooks(SNAPSHOT_PATH, temp_path)
        status_text = "Changes detected" if total else "No changes detected"
        summary_text = (
            f"The daily check reviewed {len(sheets)} worksheets and {cells:,} populated cells. "
            f"{total:,} cell-level changes were detected: {counts['added']:,} added, "
            f"{counts['removed']:,} removed and {counts['modified']:,} modified."
        )
    else:
        _, sheets, cells = read_workbook(temp_path)
        changes, counts, total = [], {"added": 0, "removed": 0, "modified": 0}, 0
        status_text = "Initial baseline stored"
        summary_text = f"The initial baseline reviewed {len(sheets)} worksheets and {cells:,} populated cells."

    history_entry = {"date": now_ct.strftime("%b %d, %Y %I:%M %p CT"), "total_changes": total, "status_text": status_text}
    history = [history_entry] + history
    payload = {
        "status": "ok",
        "message": status_text,
        "checked_at": now_utc.isoformat(),
        "checked_at_display": now_ct.strftime("%b %d, %Y %I:%M %p"),
        "source_url": EXCEL_URL,
        "summary": {"sheets_checked": len(sheets), "cells_reviewed": cells, "total_changes": total, **counts},
        "summary_text": summary_text,
        "changes": changes,
        "changes_truncated": total > MAX_WEB_CHANGES,
        "sheets": sheets,
        "history": history[:30],
    }
    write_dashboard(payload, history)
    SNAPSHOT_PATH.write_bytes(new_bytes)
    temp_path.unlink(missing_ok=True)

    details = []
    for item in changes[:MAX_EMAIL_CHANGES]:
        details.append(f"{item['sheet']}!{item['cell']} [{item['type']}]\n  Previous: {item['old'] or '(blank)'}\n  Current:  {item['new'] or '(blank)'}")
    if total > MAX_EMAIL_CHANGES:
        details.append(f"...and {total - MAX_EMAIL_CHANGES} additional changes. See the dashboard for more detail.")
    subject = f"[NERC Tracking] {status_text} - {today}"
    body = summary_text + ("\n\n" + "\n\n".join(details) if details else "")
    body += "\n\nDashboard: https://mcpbpunerc.github.io/nerc-one-stop-shop-tracker/"
    send_email(subject, body)


if __name__ == "__main__":
    main()
