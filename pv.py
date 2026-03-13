"""Google Docs API logic — fetch document content and write review notes."""

import argparse
import json
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
_REVIEW_HEADING = "🪶 Ploma Vermella Review"


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

        if text == _REVIEW_HEADING:
            break

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
    """Extract plain text from a Google Docs document body, stopping before the review section."""
    parts = []
    body = doc.get("body", {})
    for element in body.get("content", []):
        paragraph = element.get("paragraph")
        if not paragraph:
            continue
        text = "".join(
            pe.get("textRun", {}).get("content", "")
            for pe in paragraph.get("elements", [])
        ).rstrip("\n")
        if text == _REVIEW_HEADING:
            break
        for pe in paragraph.get("elements", []):
            text_run = pe.get("textRun")
            if text_run:
                content = text_run.get("content", "")
                style = text_run.get("textStyle", {})
                url = style.get("link", {}).get("url")
                strikethrough = style.get("strikethrough", False)
                if url and content.strip():
                    parts.append(f"[{content}]({url})")
                elif strikethrough and content.strip():
                    parts.append(f"~~{content}~~")
                else:
                    parts.append(content)
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
    """Return {title, text, comments} for the given Google Doc."""
    doc_id = _extract_doc_id(doc_id_or_url)
    doc = _docs_service().documents().get(documentId=doc_id).execute()
    result = _drive_service().comments().list(
        fileId=doc_id,
        fields="comments(author,content,quotedFileContent)",
        includeDeleted=False,
    ).execute()
    comments = [
        {
            "author": c.get("author", {}).get("displayName", ""),
            "content": c.get("content", ""),
            "quoted_text": c.get("quotedFileContent", {}).get("value", ""),
        }
        for c in result.get("comments", [])
    ]
    return {
        "title": doc.get("title", ""),
        "text": _extract_text(doc),
        "comments": comments,
    }


def clear_review_section(doc_id_or_url: str) -> dict:
    """Delete the '🪶 Ploma Vermella Review' section if it exists."""
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()
    doc = service.documents().get(documentId=doc_id).execute()
    content = doc.get("body", {}).get("content", [])
    for el in content:
        para = el.get("paragraph", {})
        text = "".join(
            pe.get("textRun", {}).get("content", "")
            for pe in para.get("elements", [])
        ).strip()
        if text == _REVIEW_HEADING:
            end = content[-1]["endIndex"] - 1
            service.documents().batchUpdate(
                documentId=doc_id,
                body={"requests": [{"deleteContentRange": {
                    "range": {"startIndex": el["startIndex"], "endIndex": end}
                }}]},
            ).execute()
            return {"status": "cleared"}
    return {"status": "nothing_to_clear"}


def append_review_note(doc_id_or_url: str, quoted_text: str, comment: str) -> dict:
    """
    Append a review note to the '🪶 Ploma Vermella Review' section at the end of the
    document, creating the section (Title + Subtitle) if it doesn't exist yet.

    Each note is prefixed with its paragraph location (e.g. 'Section 1: p2').
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()
    doc = service.documents().get(documentId=doc_id).execute()
    content = doc.get("body", {}).get("content", [])

    location = _paragraph_location(doc, quoted_text)
    prefix = f"{location}: " if location else ""
    note_text = f"🪶 {prefix}{comment}\n"

    # ── Create heading + subtitle if section doesn't exist yet ───────────────
    has_section = any(
        "".join(
            pe.get("textRun", {}).get("content", "")
            for pe in el.get("paragraph", {}).get("elements", [])
        ).strip() == _REVIEW_HEADING
        for el in content
    )
    if not has_section:
        insert_at = content[-1]["endIndex"] - 1
        heading_text = f"\n{_REVIEW_HEADING}\n"
        subtitle_text = f"{datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        h_start = insert_at + 1  # skip leading \n
        h_end = h_start + _utf16_len(_REVIEW_HEADING) + 1  # +1 for trailing \n
        s_start = h_end
        s_end = s_start + _utf16_len(subtitle_text)
        service.documents().batchUpdate(
            documentId=doc_id,
            body={
                "requests": [
                    {
                        "insertText": {
                            "location": {"index": insert_at},
                            "text": heading_text + subtitle_text,
                        }
                    },
                    {
                        "updateParagraphStyle": {
                            "range": {"startIndex": h_start, "endIndex": h_end},
                            "paragraphStyle": {"namedStyleType": "TITLE"},
                            "fields": "namedStyleType",
                        }
                    },
                    {
                        "updateParagraphStyle": {
                            "range": {"startIndex": s_start, "endIndex": s_end},
                            "paragraphStyle": {"namedStyleType": "SUBTITLE"},
                            "fields": "namedStyleType",
                        }
                    },
                ]
            },
        ).execute()
        # Re-fetch to get updated indices before inserting the note
        doc = service.documents().get(documentId=doc_id).execute()
        content = doc.get("body", {}).get("content", [])

    # ── Append the note at end of document ───────────────────────────────────
    insert_at = content[-1]["endIndex"] - 1
    note_end = insert_at + _utf16_len(note_text)
    service.documents().batchUpdate(
        documentId=doc_id,
        body={
            "requests": [
                {"insertText": {"location": {"index": insert_at}, "text": note_text}},
                {
                    "updateParagraphStyle": {
                        "range": {"startIndex": insert_at, "endIndex": note_end},
                        "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                        "fields": "namedStyleType",
                    }
                },
                {
                    "createParagraphBullets": {
                        "range": {"startIndex": insert_at, "endIndex": note_end},
                        "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE",
                    }
                },
            ]
        },
    ).execute()

    return {"status": "added", "location": location or "(none)", "note": note_text.strip()}


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pv",
        description="Ploma Vermella — Google Docs review tool.",
    )
    sub = parser.add_subparsers(dest="command", metavar="COMMAND")
    sub.required = True

    sub.add_parser("list", help="List all Google Docs in a Drive folder.").add_argument(
        "folder", metavar="FOLDER_URL"
    )

    sub.add_parser("fetch", help="Fetch the title and text of a Google Doc.").add_argument(
        "doc", metavar="DOC_URL"
    )

    sub.add_parser("clear", help="Delete the Ploma Vermella Review section.").add_argument(
        "doc", metavar="DOC_URL"
    )

    p_note = sub.add_parser(
        "note", help="Append a review note to the Ploma Vermella Review section."
    )
    p_note.add_argument("doc", metavar="DOC_URL")
    p_note.add_argument("quoted_text", metavar="QUOTED_TEXT",
                        help="Exact substring used to determine the note's location.")
    p_note.add_argument("comment", metavar="COMMENT", help="Note text to append.")

    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    if args.command == "list":
        result = list_folder(args.folder)
    elif args.command == "fetch":
        result = fetch_document(args.doc)
    elif args.command == "clear":
        result = clear_review_section(args.doc)
    else:
        result = append_review_note(args.doc, args.quoted_text, args.comment)

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
