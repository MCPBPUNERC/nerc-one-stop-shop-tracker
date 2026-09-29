import json
import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

DATA_PATH = Path("docs/data/nerc.json")
MAX_EMAIL_CHANGES = 25

def send_email(subject, body):
    user = os.environ.get("GMAIL_USER")
    password = os.environ.get("GMAIL_PASS")
    recipients = [x.strip() for x in os.environ.get("RECIPIENTS", "").split(",") if x.strip()]
    if not user or not password or not recipients:
        raise RuntimeError("Gmail credentials or recipients are not configured.")
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = user
    message["To"] = ", ".join(recipients)
    message.set_content(body)
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as server:
        server.starttls()
        server.login(user, password)
        server.send_message(message)

def main():
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    checked = datetime.fromisoformat(data["checked_at"].replace("Z", "+00:00")).astimezone(ZoneInfo("America/Chicago"))
    today = datetime.now(ZoneInfo("America/Chicago")).date()
    if checked.date() != today:
        raise RuntimeError(f"No completed NERC scan is available for today; latest is {checked:%Y-%m-%d %I:%M %p CT}.")

    status = data.get("message", "NERC morning summary")
    brief = data.get("brief", "")
    changes = data.get("changes", [])
    details = []
    for item in changes[:MAX_EMAIL_CHANGES]:
        details.append(
            f"{item.get('standard','Unclassified')} | {item.get('field','')} | {item.get('severity','informational').upper()}\n"
            f"  Previous: {item.get('old') or '(blank)'}\n"
            f"  Current:  {item.get('new') or '(blank)'}"
        )
    if len(changes) > MAX_EMAIL_CHANGES:
        details.append(f"...and {len(changes)-MAX_EMAIL_CHANGES} additional tracked changes. See the Control Room for detail.")

    subject = f"[BPU NERC Control Room] {status} - {today:%Y-%m-%d}"
    body = brief + (("\n\n" + "\n\n".join(details)) if details else "")
    body += f"\n\nMorning scan completed: {checked:%I:%M %p CT}"
    body += "\nControl Room: https://mcpbpunerc.github.io/nerc-one-stop-shop-tracker/"
    send_email(subject, body)

if __name__ == "__main__":
    main()
