"""Email the per-BUM files via Gmail SMTP.

Recipients live in BUM/data/emails.csv (File label -> email address);
rows with an empty email are skipped. Sending requires a Gmail address
and an App Password (Google Account -> Security -> 2-Step Verification
-> App passwords) - a normal Gmail password will NOT work.
"""

from __future__ import annotations

import base64
import io
import smtplib
import ssl
import zipfile
from datetime import date
from email.message import EmailMessage

import requests

from BUM.logic import _read_csv

_XLSX_MIME = ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def load_email_map() -> dict[str, str]:
    """File label -> recipient email, skipping blank rows."""
    return {
        row[0].strip(): row[1].strip()
        for row in _read_csv("emails.csv")[1:]
        if len(row) > 1 and row[1].strip()
    }


def _zip_name(label: str) -> str:
    return (
        "Overall Region AR (Ehab).xlsx" if label == "Ehab" else f"AR - {label}.xlsx"
    )


def build_messages(
    zip_bytes: bytes, as_of: date, sender: str, email_map: dict[str, str]
) -> list[tuple[str, EmailMessage]]:
    """Compose one email per mapped file present in the ZIP."""
    zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    names = set(zf.namelist())
    out = []
    for label, to in email_map.items():
        fname = _zip_name(label)
        if fname not in names:
            continue
        msg = EmailMessage()
        msg["From"] = sender
        msg["To"] = to
        msg["Subject"] = f"AR Report - {label} - as of {as_of:%d.%m.%Y}"
        msg.set_content(
            f"Dear {label},\n\n"
            f"Please find attached the AR report as of {as_of:%d.%m.%Y}.\n\n"
            "This email was generated automatically by the AR Backlog tool.\n"
        )
        msg.add_attachment(
            zf.read(fname),
            maintype=_XLSX_MIME[0],
            subtype=_XLSX_MIME[1],
            filename=fname,
        )
        out.append((label, msg))
    return out


def send_bum_emails(
    zip_bytes: bytes,
    as_of: date,
    sender: str,
    app_password: str,
    email_map: dict[str, str] | None = None,
    host: str = "smtp.gmail.com",
    port: int = 465,
) -> list[dict]:
    """Send one email per BUM file. Returns per-file results.

    Defaults to Gmail (SSL on 465). For Microsoft 365 use
    host="smtp.office365.com", port=587 (STARTTLS) - the mailbox needs
    'Authenticated SMTP' enabled by IT.
    """
    if email_map is None:
        email_map = load_email_map()
    if not email_map:
        raise ValueError(
            "No recipients configured - fill in BUM/data/emails.csv first."
        )
    messages = build_messages(zip_bytes, as_of, sender, email_map)
    results = []
    context = ssl.create_default_context()
    if port == 465:
        server = smtplib.SMTP_SSL(host, port, context=context)
    else:
        server = smtplib.SMTP(host, port)
        server.starttls(context=context)
    try:
        server.login(sender, app_password)
        for label, msg in messages:
            server.send_message(msg)
            results.append({"file": label, "to": msg["To"], "status": "sent"})
    finally:
        server.quit()
    return results


# ── Microsoft Graph (app registration) ─────────────────────────────────────
# Sends from a company mailbox using the app credentials IT provided
# (Tenant ID + Client ID + Client Secret with Mail.Send permission).

_GRAPH_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
_GRAPH_SEND_URL = "https://graph.microsoft.com/v1.0/users/{sender}/sendMail"

DEFAULT_BODY = "Dears,\n\nPlease find attached AR as of today.\n\nBest Regards"


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


def send_bum_emails_graph(
    zip_bytes: bytes,
    as_of: date,
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sender: str,
    recipients: list[str],
    body: str = DEFAULT_BODY,
) -> list[dict]:
    """Send every file in the ZIP as its own email to the same recipients.

    Returns per-file results; a failed file does not stop the rest.
    """
    if not recipients:
        raise ValueError("No recipients given.")
    token = _graph_token(tenant_id, client_id, client_secret)
    headers = {"Authorization": f"Bearer {token}"}
    url = _GRAPH_SEND_URL.format(sender=sender.strip())
    to_list = [{"emailAddress": {"address": a}} for a in recipients]

    zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    results = []
    for fname in zf.namelist():
        payload = {
            "message": {
                "subject": f"AR Report - {fname[:-5]} - as of {as_of:%d.%m.%Y}",
                "body": {"contentType": "Text", "content": body},
                "toRecipients": to_list,
                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": fname,
                        "contentType": _XLSX_MIME[0] + "/" + _XLSX_MIME[1],
                        "contentBytes": base64.b64encode(zf.read(fname)).decode(),
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
        results.append(
            {"file": fname, "to": ", ".join(recipients), "status": status}
        )
    return results
