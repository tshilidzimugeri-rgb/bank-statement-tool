"""Gmail API access: OAuth login, searching for statement emails, and
downloading PDF attachments. Read-only scope only - this tool never sends,
labels, or deletes anything in the inbox.
"""
from __future__ import annotations

import base64
import re
from dataclasses import dataclass, field
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from .config import ClientRule, Settings

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


@dataclass
class Attachment:
    filename: str
    attachment_id: str
    mime_type: str


@dataclass
class MessageInfo:
    message_id: str
    sender: str
    subject: str
    pdf_attachments: list[Attachment] = field(default_factory=list)


def get_service(settings: Settings):
    """Returns an authenticated Gmail API client, running the one-time
    browser OAuth consent flow if no cached token exists yet, and silently
    refreshing an expired token otherwise - so after the first authorization
    this never prompts again.
    """
    creds = None
    if settings.gmail_token_file.exists():
        creds = Credentials.from_authorized_user_file(str(settings.gmail_token_file), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not settings.gmail_client_secret_file.exists():
                raise FileNotFoundError(
                    f"Gmail OAuth client secret not found at {settings.gmail_client_secret_file}. "
                    "Download it from Google Cloud Console (APIs & Services > Credentials) and set "
                    "GMAIL_CLIENT_SECRET_FILE in .env - see README.md for the exact steps."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(settings.gmail_client_secret_file), SCOPES)
            creds = flow.run_local_server(port=0)
        settings.gmail_token_file.parent.mkdir(parents=True, exist_ok=True)
        settings.gmail_token_file.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds)


def build_search_query(client_rules: list[ClientRule], lookback_days: int) -> str:
    """Any plain PDF-attachment email in the lookback window, narrowed to
    senders/subjects mentioned in config/clients.yaml when that file has
    entries - keeps the search targeted without hardcoding clients here.
    """
    parts = ["has:attachment", "filename:pdf", f"newer_than:{lookback_days}d"]

    terms = []
    for rule in client_rules:
        for m in rule.match:
            terms.append(f'"{m}"' if " " in m else m)
    if terms:
        parts.append("(" + " OR ".join(terms) + ")")

    return " ".join(parts)


def search_message_ids(service, query: str) -> list[str]:
    ids: list[str] = []
    request = service.users().messages().list(userId="me", q=query)
    while request is not None:
        response = request.execute()
        ids.extend(m["id"] for m in response.get("messages", []))
        request = service.users().messages().list_next(previous_request=request, previous_response=response)
    return ids


def _header(headers: list[dict], name: str) -> str:
    for h in headers:
        if h.get("name", "").lower() == name.lower():
            return h.get("value", "")
    return ""


def _walk_parts(part: dict) -> list[dict]:
    parts = [part]
    for child in part.get("parts", []) or []:
        parts.extend(_walk_parts(child))
    return parts


def get_message_info(service, message_id: str) -> MessageInfo:
    msg = service.users().messages().get(userId="me", id=message_id, format="full").execute()
    payload = msg.get("payload", {})
    headers = payload.get("headers", [])
    sender = _header(headers, "From")
    subject = _header(headers, "Subject")

    attachments = []
    for part in _walk_parts(payload):
        filename = part.get("filename") or ""
        body = part.get("body", {})
        attachment_id = body.get("attachmentId")
        if filename.lower().endswith(".pdf") and attachment_id:
            attachments.append(
                Attachment(filename=filename, attachment_id=attachment_id, mime_type=part.get("mimeType", ""))
            )

    return MessageInfo(message_id=message_id, sender=sender, subject=subject, pdf_attachments=attachments)


def download_attachment(service, message_id: str, attachment_id: str) -> bytes:
    att = service.users().messages().attachments().get(
        userId="me", messageId=message_id, id=attachment_id
    ).execute()
    return base64.urlsafe_b64decode(att["data"])


_UNSAFE_CHARS = re.compile(r'[<>:"/\\|?*]')


def safe_local_filename(message_id: str, attachment_filename: str) -> str:
    cleaned = _UNSAFE_CHARS.sub("_", attachment_filename)
    return f"{message_id}_{cleaned}"
