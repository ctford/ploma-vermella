"""Google Docs API logic — fetch content, comments, and post comments."""

import os
import re
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
]

CREDENTIALS_DIR = Path(__file__).parent / "credentials"
CLIENT_SECRET = CREDENTIALS_DIR / "client_secret.json"
TOKEN_FILE = CREDENTIALS_DIR / "token.json"

_DOC_URL_RE = re.compile(r"/document/d/([a-zA-Z0-9_-]+)")
_FOLDER_URL_RE = re.compile(r"/folders/([a-zA-Z0-9_-]+)")


def _extract_doc_id(doc_id_or_url: str) -> str:
    """Return bare doc ID whether given a full URL or already a bare ID."""
    m = _DOC_URL_RE.search(doc_id_or_url)
    return m.group(1) if m else doc_id_or_url.strip()


def _get_credentials() -> Credentials:
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET), SCOPES)
            creds = flow.run_local_server(port=0)
        CREDENTIALS_DIR.mkdir(parents=True, exist_ok=True)
        TOKEN_FILE.write_text(creds.to_json())
    return creds


def _docs_service():
    return build("docs", "v1", credentials=_get_credentials())


def _drive_service():
    return build("drive", "v3", credentials=_get_credentials())


def _extract_text(doc: dict) -> str:
    """Extract plain text from a Google Docs document body."""
    parts = []
    body = doc.get("body", {})
    for element in body.get("content", []):
        paragraph = element.get("paragraph")
        if not paragraph:
            continue
        for pe in paragraph.get("elements", []):
            text_run = pe.get("textRun")
            if text_run:
                parts.append(text_run.get("content", ""))
    return "".join(parts)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def _extract_folder_id(folder_id_or_url: str) -> str:
    """Return bare folder ID whether given a full Drive URL or already a bare ID."""
    m = _FOLDER_URL_RE.search(folder_id_or_url)
    return m.group(1) if m else folder_id_or_url.strip()


def list_folder(folder_id_or_url: str) -> list[dict]:
    """Return all Google Docs in a Drive folder as [{id, name, url}]."""
    folder_id = _extract_folder_id(folder_id_or_url)
    service = _drive_service()
    result = (
        service.files()
        .list(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false",
            fields="files(id, name)",
            orderBy="name",
        )
        .execute()
    )
    return [
        {
            "id": f["id"],
            "name": f["name"],
            "url": f"https://docs.google.com/document/d/{f['id']}",
        }
        for f in result.get("files", [])
    ]


def fetch_document(doc_id_or_url: str) -> dict:
    """Return {title, text} for the given Google Doc."""
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()
    doc = service.documents().get(documentId=doc_id).execute()
    return {
        "title": doc.get("title", ""),
        "text": _extract_text(doc),
    }


def fetch_comments(doc_id_or_url: str) -> list[dict]:
    """Return existing comments as [{id, author, content, quoted_text}]."""
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _drive_service()
    result = (
        service.comments()
        .list(fileId=doc_id, fields="comments(id,author,content,quotedFileContent)", includeDeleted=False)
        .execute()
    )
    comments = []
    for c in result.get("comments", []):
        comments.append(
            {
                "id": c.get("id"),
                "author": c.get("author", {}).get("displayName", ""),
                "content": c.get("content", ""),
                "quoted_text": c.get("quotedFileContent", {}).get("value", ""),
            }
        )
    return comments


def post_comment(doc_id_or_url: str, quoted_text: str, comment: str) -> dict:
    """
    Post a new comment on the document associated with quoted_text.

    Comments appear in the sidebar with the quoted text visible. Google Docs
    does not expose an API to create text-highlighted (anchored) comments
    programmatically — that is locked to the UI.

    Does a server-side duplicate check: if quoted_text already appears in
    an existing comment, returns a no-op result instead of posting.
    """
    doc_id = _extract_doc_id(doc_id_or_url)

    existing = fetch_comments(doc_id)
    for c in existing:
        if c["quoted_text"].strip() == quoted_text.strip():
            return {
                "status": "skipped",
                "reason": "quoted_text already has a comment",
                "existing_comment_id": c["id"],
            }

    body = {
        "content": comment,
        "quotedFileContent": {"mimeType": "text/plain", "value": quoted_text},
    }
    created = _drive_service().comments().create(fileId=doc_id, body=body, fields="id,content,author").execute()
    return {
        "status": "posted",
        "id": created.get("id"),
        "content": created.get("content"),
        "author": created.get("author", {}).get("displayName", ""),
    }
