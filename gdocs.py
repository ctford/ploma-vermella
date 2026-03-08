"""Google Docs API logic — fetch content, comments, and post comments."""

import re
from datetime import datetime
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


_HEADING_STYLES = {"HEADING_1", "HEADING_2", "HEADING_3", "HEADING_4", "TITLE"}
_TAB_TITLE = "🪶 Ploma Vermella"


def _utf16_len(s: str) -> int:
    """Length of s in UTF-16 code units (the unit the Docs API uses for indices)."""
    return sum(2 if ord(c) > 0xFFFF else 1 for c in s)


def _paragraph_location(doc: dict, quoted_text: str) -> str:
    """Return a location string like 'Section 1: p2' for the paragraph containing quoted_text."""
    current_heading = None
    para_count = 0
    needle = quoted_text.strip()

    for element in doc.get("body", {}).get("content", []):
        paragraph = element.get("paragraph")
        if not paragraph:
            continue

        style = paragraph.get("paragraphStyle", {}).get("namedStyleType", "")
        text = "".join(
            pe.get("textRun", {}).get("content", "")
            for pe in paragraph.get("elements", [])
        ).rstrip("\n")

        if not text:
            continue

        if style in _HEADING_STYLES:
            current_heading = text
            para_count = 0
        else:
            para_count += 1
            if needle in text:
                prefix = f"{current_heading}: " if current_heading else ""
                return f"{prefix}p{para_count}"

    return ""


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
# Tab-based review helpers
# ---------------------------------------------------------------------------


def _get_review_tab(doc: dict) -> dict | None:
    """Return the review tab dict or None."""
    for tab in doc.get("tabs", []):
        if tab.get("tabProperties", {}).get("title") == _TAB_TITLE:
            return tab
    return None


def _tab_content(tab: dict) -> list:
    """Return the body content list from a tab."""
    return tab.get("documentTab", {}).get("body", {}).get("content", [])


def _tab_end_index(content: list) -> int:
    """Return the insert index just before the required final newline."""
    return content[-1]["endIndex"] - 1 if content else 1


def _is_fresh_tab(content: list) -> bool:
    """True if the tab has no meaningful text content."""
    for el in content:
        para = el.get("paragraph")
        if not para:
            continue
        text = "".join(
            pe.get("textRun", {}).get("content", "")
            for pe in para.get("elements", [])
        ).strip()
        if text:
            return False
    return True


def _tab_has_heading(content: list, heading_text: str) -> bool:
    """Return True if content already contains a heading matching heading_text."""
    for el in content:
        para = el.get("paragraph", {})
        if para.get("paragraphStyle", {}).get("namedStyleType") in _HEADING_STYLES:
            text = "".join(
                pe.get("textRun", {}).get("content", "")
                for pe in para.get("elements", [])
            ).strip()
            if text == heading_text:
                return True
    return False


def _section_end_index(content: list, heading_text: str) -> int:
    """Return the index at which to append content at the end of the named section.

    Scans forward from the matching heading; stops at the next heading or end of content.
    """
    in_section = False
    for el in content:
        para = el.get("paragraph")
        if not para:
            continue
        style = para.get("paragraphStyle", {}).get("namedStyleType", "")
        text = "".join(
            pe.get("textRun", {}).get("content", "")
            for pe in para.get("elements", [])
        ).strip()

        if in_section and style in _HEADING_STYLES:
            return el["startIndex"]

        if not in_section and style in _HEADING_STYLES and text == heading_text:
            in_section = True

    return _tab_end_index(content)


def _find_section_for_text(content: list, quoted_text: str) -> str | None:
    """Return the heading text of the section containing quoted_text."""
    current_heading = None
    needle = quoted_text.strip()
    for el in content:
        paragraph = el.get("paragraph")
        if not paragraph:
            continue
        style = paragraph.get("paragraphStyle", {}).get("namedStyleType", "")
        text = "".join(
            pe.get("textRun", {}).get("content", "")
            for pe in paragraph.get("elements", [])
        ).rstrip("\n")
        if not text:
            continue
        if style in _HEADING_STYLES:
            current_heading = text
        elif needle in text:
            return current_heading
    return None


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
            q=(
                f"'{folder_id}' in parents"
                " and mimeType='application/vnd.google-apps.document'"
                " and trashed=false"
            ),
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
        .list(
            fileId=doc_id,
            fields="comments(id,author,content,quotedFileContent)",
            includeDeleted=False,
        )
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


def append_review_note(doc_id_or_url: str, quoted_text: str, comment: str) -> dict:
    """
    Append a review note to the '🪶 Ploma Vermella' review tab, creating the tab if
    it doesn't exist yet.

    On first use the tab gets a Title ('Ploma Vermella') and a Subtitle with today's date.
    Notes are grouped under the same section heading as in the main document.
    Each note is prefixed with its paragraph location (e.g. 'Section 1: p2').
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()

    # Fetch with tabs so we can read both main body and existing review tab.
    doc = service.documents().get(documentId=doc_id, includeTabsContent=True).execute()

    # Main body lives in the first tab when includeTabsContent=True.
    tabs = doc.get("tabs", [])
    main_content = (
        tabs[0].get("documentTab", {}).get("body", {}).get("content", [])
        if tabs
        else doc.get("body", {}).get("content", [])
    )
    main_doc = {"body": {"content": main_content}}

    section_heading = _find_section_for_text(main_content, quoted_text)
    location = _paragraph_location(main_doc, quoted_text)
    prefix = f"{location}: " if location else ""
    note_text = f"🪶 {prefix}{comment}\n"

    # ── Create review tab if needed ──────────────────────────────────────────
    tab = _get_review_tab(doc)
    if tab is None:
        service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": [{"createTab": {"tabProperties": {"title": _TAB_TITLE}}}]},
        ).execute()
        doc = service.documents().get(documentId=doc_id, includeTabsContent=True).execute()
        tab = _get_review_tab(doc)

    tab_id = tab["tabProperties"]["tabId"]
    content = _tab_content(tab)

    # ── Initialise fresh tab with title + subtitle ────────────────────────────
    if _is_fresh_tab(content):
        date_str = datetime.now().strftime("%Y-%m-%d")
        title_text = "Ploma Vermella\n"
        subtitle_text = f"Reviewed: {date_str}\n"
        insert_at = _tab_end_index(content)
        t_end = insert_at + _utf16_len(title_text)
        s_end = t_end + _utf16_len(subtitle_text)
        service.documents().batchUpdate(
            documentId=doc_id,
            body={
                "requests": [
                    {
                        "insertText": {
                            "location": {"index": insert_at, "tabId": tab_id},
                            "text": title_text + subtitle_text,
                        }
                    },
                    {
                        "updateParagraphStyle": {
                            "range": {"startIndex": insert_at, "endIndex": t_end, "tabId": tab_id},
                            "paragraphStyle": {"namedStyleType": "TITLE"},
                            "fields": "namedStyleType",
                        }
                    },
                    {
                        "updateParagraphStyle": {
                            "range": {"startIndex": t_end, "endIndex": s_end, "tabId": tab_id},
                            "paragraphStyle": {"namedStyleType": "SUBTITLE"},
                            "fields": "namedStyleType",
                        }
                    },
                ]
            },
        ).execute()
        doc = service.documents().get(documentId=doc_id, includeTabsContent=True).execute()
        tab = _get_review_tab(doc)
        content = _tab_content(tab)

    # ── Add section heading in review tab if missing ─────────────────────────
    if section_heading and not _tab_has_heading(content, section_heading):
        insert_at = _tab_end_index(content)
        heading_text = f"{section_heading}\n"
        h_end = insert_at + _utf16_len(heading_text)
        service.documents().batchUpdate(
            documentId=doc_id,
            body={
                "requests": [
                    {
                        "insertText": {
                            "location": {"index": insert_at, "tabId": tab_id},
                            "text": heading_text,
                        }
                    },
                    {
                        "updateParagraphStyle": {
                            "range": {"startIndex": insert_at, "endIndex": h_end, "tabId": tab_id},
                            "paragraphStyle": {"namedStyleType": "HEADING_1"},
                            "fields": "namedStyleType",
                        }
                    },
                ]
            },
        ).execute()
        doc = service.documents().get(documentId=doc_id, includeTabsContent=True).execute()
        tab = _get_review_tab(doc)
        content = _tab_content(tab)

    # ── Insert the note ───────────────────────────────────────────────────────
    if section_heading and _tab_has_heading(content, section_heading):
        insert_at = _section_end_index(content, section_heading)
    else:
        insert_at = _tab_end_index(content)

    note_end = insert_at + _utf16_len(note_text)
    service.documents().batchUpdate(
        documentId=doc_id,
        body={
            "requests": [
                {
                    "insertText": {
                        "location": {"index": insert_at, "tabId": tab_id},
                        "text": note_text,
                    }
                },
                {
                    "updateParagraphStyle": {
                        "range": {"startIndex": insert_at, "endIndex": note_end, "tabId": tab_id},
                        "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                        "fields": "namedStyleType",
                    }
                },
            ]
        },
    ).execute()

    return {
        "status": "added",
        "location": location or "(none)",
        "note": note_text.strip(),
        "tab": _TAB_TITLE,
    }


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

    doc = _docs_service().documents().get(documentId=doc_id).execute()
    location = _paragraph_location(doc, quoted_text)
    prefix = f"{location}: " if location else ""
    body = {
        "content": f"🪶 {prefix}{comment}",
        "quotedFileContent": {"mimeType": "text/plain", "value": quoted_text},
    }
    created = (
        _drive_service()
        .comments()
        .create(fileId=doc_id, body=body, fields="id,content,author")
        .execute()
    )
    return {
        "status": "posted",
        "id": created.get("id"),
        "content": created.get("content"),
        "author": created.get("author", {}).get("displayName", ""),
    }
