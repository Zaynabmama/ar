"""Email the finished Credit AR Report workbook to a fixed distribution list.

Recipients live in ar_report/data/recipients.csv (To, CC, Subject, Body
columns, one row) using the same Outlook-style "Name" <email>; "Name2"
<email2> format as BUM/data/emails.csv. Unlike the BUM tool there is only
one output file, so there is only one row of recipients here.
"""

from __future__ import annotations

import base64
import csv
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import getaddresses
from pathlib import Path

import requests

_DATA_DIR = Path(__file__).parent / "data"
_XLSX_MIME = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")

DEFAULT_SUBJECT = "AR Ageing Report"
DEFAULT_BODY = (
    "Dears,\n\nPlease find attached the AR Ageing report.\n\n"
    "If you need any adjustments, please reach out to the Credit team.\n\n"
    "Kind Regards,\nCredit Department"
)


def _parse_addresses(raw: str) -> list[tuple[str, str]]:
    """Parse a "Name" <email>; "Name2" <email2> string into (name, email) pairs."""
    if not raw or not raw.strip():
        return []
    return [(name, addr) for name, addr in getaddresses([raw.replace(";", ",")]) if addr]


def _header_value(pairs: list[tuple[str, str]]) -> str:
    return ", ".join(f'"{name}" <{addr}>' if name else addr for name, addr in pairs)


def _graph_addr_list(pairs: list[tuple[str, str]]) -> list[dict]:
    return [
        {"emailAddress": {"address": addr, **({"name": name} if name else {})}}
        for name, addr in pairs
    ]


def load_recipients() -> dict:
    """Returns {'to', 'cc', 'subject', 'body'} from ar_report/data/recipients.csv,
    or {} if the To column is blank."""
    path = _DATA_DIR / "recipients.csv"
    with path.open(newline="", encoding="utf-8-sig") as f:
        rows = [r for r in csv.reader(f) if r]
    if len(rows) < 2 or not rows[1] or not rows[1][0].strip():
        return {}
    row = rows[1]
    to = _parse_addresses(row[0])
    cc = _parse_addresses(row[1]) if len(row) > 1 else []
    subject = row[2].strip() if len(row) > 2 and row[2].strip() else DEFAULT_SUBJECT
    body = row[3].strip() if len(row) > 3 and row[3].strip() else DEFAULT_BODY
    return {"to": to, "cc": cc, "subject": subject, "body": body}


# ── Microsoft Graph (app registration) ─────────────────────────────────────
# Same tenant/app registration already used by the BUM tool (Mail.Send,
# sender Donotreply@mindware.net) - no new IT request needed.

_GRAPH_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
_GRAPH_SEND_URL = "https://graph.microsoft.com/v1.0/users/{sender}/sendMail"


def _graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    resp = requests.post(
        _GRAPH_TOKEN_URL.format(tenant=tenant_id.strip()),
        data={
            "client_id": client_id.strip(),
            "client_secret": client_secret.strip(),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Could not sign in to Microsoft (HTTP {resp.status_code}): "
            f"{resp.json().get('error_description', resp.text)[:300]}"
        )
    return resp.json()["access_token"]


def send_report_email_graph(
    workbook_bytes: bytes,
    file_name: str,
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sender: str,
    recipients: dict | None = None,
) -> dict:
    """Send the workbook as a single email via Microsoft Graph.

    Returns {'to': ..., 'status': 'sent' | 'FAILED (...)'}.
    """
    if recipients is None:
        recipients = load_recipients()
    if not recipients or not recipients["to"]:
        raise ValueError(
            "No recipients configured - fill in ar_report/data/recipients.csv first."
        )
    token = _graph_token(tenant_id, client_id, client_secret)
    headers = {"Authorization": f"Bearer {token}"}
    url = _GRAPH_SEND_URL.format(sender=sender.strip())
    payload = {
        "message": {
            "subject": recipients["subject"],
            "body": {"contentType": "Text", "content": recipients["body"]},
            "toRecipients": _graph_addr_list(recipients["to"]),
            "ccRecipients": _graph_addr_list(recipients["cc"]),
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": file_name,
                    "contentType": _XLSX_MIME[0] + "/" + _XLSX_MIME[1],
                    "contentBytes": base64.b64encode(workbook_bytes).decode(),
                }
            ],
        },
        "saveToSentItems": True,
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=120)
    if resp.status_code == 202:
        status = "sent"
    else:
        try:
            detail = resp.json().get("error", {}).get("message", "")[:200]
        except Exception:
            detail = resp.text[:200]
        status = f"FAILED (HTTP {resp.status_code}): {detail}"
    to_display = ", ".join(addr for _, addr in recipients["to"])
    return {"to": to_display, "status": status}


def send_report_email_smtp(
    workbook_bytes: bytes,
    file_name: str,
    sender: str,
    app_password: str,
    recipients: dict | None = None,
    host: str = "smtp.gmail.com",
    port: int = 465,
) -> dict:
    """SMTP fallback (Gmail App Password, or Office365 with host/port set).

    Returns {'to': ..., 'status': 'sent'}.
    """
    if recipients is None:
        recipients = load_recipients()
    if not recipients or not recipients["to"]:
        raise ValueError(
            "No recipients configured - fill in ar_report/data/recipients.csv first."
        )
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = _header_value(recipients["to"])
    if recipients["cc"]:
        msg["Cc"] = _header_value(recipients["cc"])
    msg["Subject"] = recipients["subject"]
    msg.set_content(recipients["body"])
    msg.add_attachment(
        workbook_bytes,
        maintype=_XLSX_MIME[0],
        subtype=_XLSX_MIME[1],
        filename=file_name,
    )
    context = ssl.create_default_context()
    if port == 465:
        server = smtplib.SMTP_SSL(host, port, context=context)
    else:
        server = smtplib.SMTP(host, port)
        server.starttls(context=context)
    try:
        server.login(sender, app_password)
        server.send_message(msg)
    finally:
        server.quit()
    return {"to": msg["To"], "status": "sent"}
