import os
import io
from datetime import datetime
import requests
import pandas as pd
import smtplib
from email.message import EmailMessage
from pathlib import Path

# Direct NERC Excel URL (decoded from the Office viewer link)
EXCEL_URL = "https://www.nerc.com/globalassets/align-reports/one-stop-shop.xlsx"

# Snapshot file stored inside the repo
SNAPSHOT_PATH = Path("last_snapshot.xlsx")

# Environment variables (values will come from GitHub Secrets)
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASS = os.environ.get("GMAIL_PASS")
RECIPIENTS = os.environ.get(
    "RECIPIENTS",
    "mikep@mcphersonpower.com,tommys@mcphersonpower.com",
).split(",")


def download_excel():
    """Download the current NERC One Stop Shop Excel file as bytes."""
    resp = requests.get(EXCEL_URL, timeout=60)
    resp.raise_for_status()
    return resp.content


def bytes_to_dataframe(file_bytes):
    """Convert Excel bytes to a pandas DataFrame (first sheet)."""
    buffer = io.BytesIO(file_bytes)
    df = pd.read_excel(buffer, sheet_name=0)
    return df


def compute_row_set(df):
    """
    Represent each row as a tuple of strings so we can compare
    entire rows between days. NaNs become empty strings.
    """
    df_filled = df.fillna("")
    rows = []
    for _, row in df_filled.iterrows():
        rows.append(tuple(str(v).strip() for v in row.tolist()))
    return set(rows)


def build_diff_report(old_bytes, new_bytes, max_sample=20):
    """
    Compare yesterday vs today at the row level.
    We treat any row that disappeared as 'removed' and any new row as 'added'.
    """
    old_df = bytes_to_dataframe(old_bytes)
    new_df = bytes_to_dataframe(new_bytes)

    old_set = compute_row_set(old_df)
    new_set = compute_row_set(new_df)

    added = new_set - old_set
    removed = old_set - new_set

    lines = []
    lines.append(f"Total rows yesterday: {len(old_set)}")
    lines.append(f"Total rows today    : {len(new_set)}")
    lines.append(f"Rows added          : {len(added)}")
    lines.append(f"Rows removed        : {len(removed)}")
    lines.append("")

    if added:
        lines.append("=== Added rows (sample) ===")
        for i, row in enumerate(sorted(added)):
            if i >= max_sample:
                lines.append(f"... ({len(added) - max_sample} more added rows)")
                break
            lines.append(" | ".join(row))
        lines.append("")

    if removed:
        lines.append("=== Removed rows (sample) ===")
        for i, row in enumerate(sorted(removed)):
            if i >= max_sample:
                lines.append(f"... ({len(removed) - max_sample} more removed rows)")
                break
            lines.append(" | ".join(row))

    if not added and not removed:
        lines.append("No row-level changes detected (files are identical at row level).")

    return "\n".join(lines)


def send_email(subject, body):
    if not GMAIL_USER or not GMAIL_PASS:
        raise RuntimeError("GMAIL_USER or GMAIL_PASS environment variables are not set.")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = GMAIL_USER
    msg["To"] = ", ".join(RECIPIENTS)
    msg.set_content(body)

    with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as server:
        server.starttls()
        server.login(GMAIL_USER, GMAIL_PASS)
        server.send_message(msg)


def main():
    today_str = datetime.utcnow().strftime("%Y-%m-%d")

    # Download today's file
    excel_bytes = download_excel()

    if SNAPSHOT_PATH.exists():
        # Compare with yesterday's snapshot
        old_bytes = SNAPSHOT_PATH.read_bytes()
        report = build_diff_report(old_bytes, excel_bytes)
        changed = "No row-level changes detected" not in report

        if changed:
            subject = f"[NERC One Stop Shop] Changes detected for {today_str}"
        else:
            subject = f"[NERC One Stop Shop] No changes detected for {today_str}"

        body = (
            f"NERC One Stop Shop daily check for {today_str}.\n\n"
            f"{report}\n"
            "\n--\n"
            "This email was generated automatically by your NERC tracking bot."
        )
    else:
        # First run: just store baseline
        subject = f"[NERC One Stop Shop] Initial snapshot stored ({today_str})"
        body = (
            "This is the first run of the NERC One Stop Shop tracker.\n"
            "Today's file has been stored as the baseline. No diff to report yet.\n"
        )

    # Save today's snapshot so tomorrow has something to compare to
    SNAPSHOT_PATH.write_bytes(excel_bytes)

    # Email report
    send_email(subject, body)


if __name__ == "__main__":
    main()
