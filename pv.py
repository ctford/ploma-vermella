"""Google Docs API logic — fetch document content and write review notes."""

import argparse
import html
import json
import re
import uuid
import zipfile
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
_EPUB_NS = "http://www.idpf.org/2007/ops"


def _utf16_len(s: str) -> int:
    """Length of s in UTF-16 code units (the unit the Docs API uses for indices)."""
    return sum(2 if ord(c) > 0xFFFF else 1 for c in s)


def _parse_table_row(line: str) -> list[str]:
    """Parse a markdown-style pipe table row into stripped cell values."""
    text = line.strip()
    if not (text.startswith("|") and text.endswith("|")):
        raise ValueError(f"not a pipe table row: {line!r}")
    return [cell.strip() for cell in text[1:-1].split("|")]


def _is_table_separator(line: str, columns: int) -> bool:
    """Return True if line looks like a markdown table separator row."""
    try:
        cells = _parse_table_row(line)
    except ValueError:
        return False
    if len(cells) != columns:
        return False
    for cell in cells:
        if not cell:
            return False
        if any(ch not in "-: " for ch in cell):
            return False
        if "-" not in cell:
            return False
    return True


def _parse_append_blocks(text: str) -> list[dict]:
    """Parse append text into paragraph, bullet, and table blocks."""
    lines = text.splitlines()
    blocks = []
    i = 0
    need_space_above = False

    while i < len(lines):
        line = lines[i]
        if not line.strip():
            need_space_above = True
            i += 1
            continue

        # Parse markdown-style pipe tables.
        if line.strip().startswith("|") and i + 1 < len(lines):
            try:
                header = _parse_table_row(line)
            except ValueError:
                header = None
            if header and _is_table_separator(lines[i + 1], len(header)):
                rows = [header]
                i += 2
                while i < len(lines):
                    next_line = lines[i]
                    if not next_line.strip():
                        break
                    if not next_line.strip().startswith("|"):
                        break
                    cells = _parse_table_row(next_line)
                    if len(cells) != len(header):
                        break
                    rows.append(cells)
                    i += 1
                blocks.append({
                    "type": "table",
                    "rows": rows,
                    "space_above": need_space_above,
                })
                need_space_above = False
                continue

        block_type = "bullet" if line.startswith("- ") else "paragraph"
        text_value = line[2:] if block_type == "bullet" else line
        blocks.append({
            "type": block_type,
            "text": text_value,
            "space_above": need_space_above,
        })
        need_space_above = False
        i += 1

    return blocks


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


def _text_from_elements(elements: list[dict]) -> str:
    """Extract plain paragraph text from Docs API paragraph elements."""
    return "".join(
        pe.get("textRun", {}).get("content", "")
        for pe in elements
    ).rstrip("\n")


def _extract_blocks(doc: dict) -> list[dict]:
    """
    Extract document blocks suitable for EPUB rendering.

    Returns a list of dictionaries with type keys:
    - {"type": "heading", "level": 1..4, "text": "..."}
    - {"type": "paragraph", "text": "..."}
    - {"type": "list_item", "text": "..."}
    """
    blocks = []
    for element in doc.get("body", {}).get("content", []):
        paragraph = element.get("paragraph")
        if not paragraph:
            continue

        text = _text_from_elements(paragraph.get("elements", []))
        if not text:
            continue
        if text == _REVIEW_HEADING:
            break

        style = paragraph.get("paragraphStyle", {}).get("namedStyleType", "")
        if "bullet" in paragraph:
            blocks.append({"type": "list_item", "text": text})
            continue

        if style == "TITLE":
            blocks.append({"type": "heading", "level": 1, "text": text})
        elif style == "HEADING_1":
            blocks.append({"type": "heading", "level": 2, "text": text})
        elif style == "HEADING_2":
            blocks.append({"type": "heading", "level": 3, "text": text})
        elif style in {"HEADING_3", "HEADING_4"}:
            blocks.append({"type": "heading", "level": 4, "text": text})
        else:
            blocks.append({"type": "paragraph", "text": text})

    return blocks


def _slugify(value: str) -> str:
    """Return a filesystem-friendly ASCII-ish slug."""
    slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return slug or "book"


def _chapter_filename(index: int) -> str:
    """Return a stable EPUB filename for a chapter."""
    return f"chapter-{index:02d}.xhtml"


def _blocks_to_xhtml(title: str, blocks: list[dict]) -> str:
    """Render extracted blocks as a simple XHTML chapter."""
    parts = []
    in_list = False
    for block in blocks:
        block_type = block["type"]
        text = html.escape(block["text"])
        if block_type == "list_item":
            if not in_list:
                parts.append("<ul>")
                in_list = True
            parts.append(f"<li>{text}</li>")
            continue

        if in_list:
            parts.append("</ul>")
            in_list = False

        if block_type == "heading":
            level = max(1, min(6, int(block["level"])))
            parts.append(f"<h{level}>{text}</h{level}>")
        else:
            parts.append(f"<p>{text}</p>")

    if in_list:
        parts.append("</ul>")

    body = "\n    ".join(parts)
    doc_title = html.escape(title)
    return (
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
        "<html xmlns=\"http://www.w3.org/1999/xhtml\" "
        f"xmlns:epub=\"{_EPUB_NS}\">\n"
        "<head>\n"
        f"  <title>{doc_title}</title>\n"
        "  <link rel=\"stylesheet\" type=\"text/css\" href=\"styles.css\"/>\n"
        "</head>\n"
        "<body>\n"
        f"  <section epub:type=\"chapter\">\n    {body}\n  </section>\n"
        "</body>\n"
        "</html>\n"
    )


def _epub_nav(book_title: str, chapters: list[dict]) -> str:
    """Return the EPUB navigation document."""
    items = "\n      ".join(
        f"<li><a href=\"{html.escape(ch['filename'])}\">{html.escape(ch['title'])}</a></li>"
        for ch in chapters
    )
    safe_title = html.escape(book_title)
    return (
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
        "<html xmlns=\"http://www.w3.org/1999/xhtml\" "
        f"xmlns:epub=\"{_EPUB_NS}\">\n"
        "<head>\n"
        f"  <title>{safe_title}</title>\n"
        "</head>\n"
        "<body>\n"
        "  <nav epub:type=\"toc\" id=\"toc\">\n"
        f"    <h1>{safe_title}</h1>\n"
        "    <ol>\n"
        f"      {items}\n"
        "    </ol>\n"
        "  </nav>\n"
        "</body>\n"
        "</html>\n"
    )


def _epub_package(book_title: str, book_id: str, chapters: list[dict]) -> str:
    """Return the OPF package document."""
    manifest_items = [
        '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
        '<item id="css" href="styles.css" media-type="text/css"/>',
    ]
    spine_items = []
    for i, chapter in enumerate(chapters, start=1):
        manifest_items.append(
            f'<item id="chap{i}" href="{chapter["filename"]}" media-type="application/xhtml+xml"/>'
        )
        spine_items.append(f'<itemref idref="chap{i}"/>')

    manifest = "\n    ".join(manifest_items)
    spine = "\n    ".join(spine_items)
    modified = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    safe_title = html.escape(book_title)
    return (
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
        "<package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" "
        f'unique-identifier="bookid" xml:lang="en">\n'
        "  <metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\n"
        f"    <dc:identifier id=\"bookid\">urn:uuid:{book_id}</dc:identifier>\n"
        f"    <dc:title>{safe_title}</dc:title>\n"
        "    <dc:language>en</dc:language>\n"
        f"    <meta property=\"dcterms:modified\">{modified}</meta>\n"
        "  </metadata>\n"
        "  <manifest>\n"
        f"    {manifest}\n"
        "  </manifest>\n"
        "  <spine>\n"
        f"    {spine}\n"
        "  </spine>\n"
        "</package>\n"
    )


def _default_epub_title(chapter_titles: list[str]) -> str:
    """Return a default book title for a collection of chapters."""
    if len(chapter_titles) == 1:
        return chapter_titles[0]
    return "Ploma Vermella Export"


def _default_epub_output_path(book_title: str, stamp: datetime | None = None) -> Path:
    """Return the default gitignored EPUB output path."""
    stamp = stamp or datetime.now()
    return Path("dist") / f"{_slugify(book_title)}-{stamp.strftime('%Y%m%d')}.epub"


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
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
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


def move_document(doc_id_or_url: str, folder_id_or_url: str) -> dict:
    """Move a Drive file into a folder. Removes all existing parents."""
    doc_id = _extract_doc_id(doc_id_or_url)
    folder_id = _extract_folder_id(folder_id_or_url)
    service = _drive_service()
    current = (
        service.files()
        .get(fileId=doc_id, fields="id,name,parents", supportsAllDrives=True)
        .execute()
    )
    prev_parents = ",".join(current.get("parents", []))
    updated = (
        service.files()
        .update(
            fileId=doc_id,
            addParents=folder_id,
            removeParents=prev_parents,
            fields="id,name,parents,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )
    return {
        "status": "moved",
        "id": updated["id"],
        "name": updated["name"],
        "from": current.get("parents", []),
        "to": updated.get("parents", []),
        "url": updated.get("webViewLink", f"https://docs.google.com/document/d/{updated['id']}"),
    }


def copy_document(
    doc_id_or_url: str,
    folder_id_or_url: str,
    name: str | None = None,
) -> dict:
    """Copy a Drive file into a folder, optionally renaming it."""
    doc_id = _extract_doc_id(doc_id_or_url)
    folder_id = _extract_folder_id(folder_id_or_url)
    service = _drive_service()
    body: dict = {"parents": [folder_id]}
    if name is not None:
        body["name"] = name
    copied = (
        service.files()
        .copy(
            fileId=doc_id,
            body=body,
            fields="id,name,parents,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )
    return {
        "status": "copied",
        "source_id": doc_id,
        "id": copied["id"],
        "name": copied["name"],
        "parents": copied.get("parents", []),
        "url": copied.get("webViewLink", f"https://docs.google.com/document/d/{copied['id']}"),
    }


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


def build_epub(
    doc_ids_or_urls: list[str],
    output: str | None = None,
    title: str | None = None,
) -> dict:
    """Build an EPUB from multiple Google Docs, excluding PV review sections."""
    docs_service = _docs_service()
    chapters = []
    for index, doc_ref in enumerate(doc_ids_or_urls, start=1):
        doc_id = _extract_doc_id(doc_ref)
        doc = docs_service.documents().get(documentId=doc_id).execute()
        chapter_title = doc.get("title", f"Chapter {index}")
        chapters.append({
            "title": chapter_title,
            "filename": _chapter_filename(index),
            "xhtml": _blocks_to_xhtml(chapter_title, _extract_blocks(doc)),
        })

    book_title = title or _default_epub_title([chapter["title"] for chapter in chapters])
    output_path = Path(output) if output else _default_epub_output_path(book_title)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    book_id = str(uuid.uuid4())
    nav = _epub_nav(book_title, chapters)
    package = _epub_package(book_title, book_id, chapters)
    stylesheet = (
        "body { font-family: serif; line-height: 1.4; }\n"
        "h1, h2, h3, h4 { font-family: sans-serif; }\n"
        "section { max-width: 42em; margin: 0 auto; }\n"
        "p, li { margin: 0.6em 0; }\n"
    )

    with zipfile.ZipFile(output_path, "w") as epub:
        epub.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)
        epub.writestr(
            "META-INF/container.xml",
            "<?xml version=\"1.0\"?>\n"
            "<container version=\"1.0\" "
            "xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">\n"
            "  <rootfiles>\n"
            "    <rootfile full-path=\"OEBPS/content.opf\" "
            "media-type=\"application/oebps-package+xml\"/>\n"
            "  </rootfiles>\n"
            "</container>\n",
        )
        epub.writestr("OEBPS/content.opf", package)
        epub.writestr("OEBPS/nav.xhtml", nav)
        epub.writestr("OEBPS/styles.css", stylesheet)
        for chapter in chapters:
            epub.writestr(f"OEBPS/{chapter['filename']}", chapter["xhtml"])

    return {
        "status": "built",
        "title": book_title,
        "output": str(output_path),
        "chapters": [
            {"title": chapter["title"], "filename": chapter["filename"]}
            for chapter in chapters
        ],
    }


def append_content(doc_id_or_url: str, heading: str, text: str) -> dict:
    """
    Append a headed section into the Ploma Vermella Review section, creating
    it if needed. Heading is inserted as HEADING_2. In the body text, lines
    starting with '- ' become bullets; blank lines are skipped; all others
    are normal paragraphs.
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()
    doc = service.documents().get(documentId=doc_id).execute()
    content = doc.get("body", {}).get("content", [])

    # Ensure the review section exists
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
        h_start = insert_at + 1
        h_end = h_start + _utf16_len(_REVIEW_HEADING) + 1
        s_start = h_end
        s_end = s_start + _utf16_len(subtitle_text)
        service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": [
                {"insertText": {"location": {"index": insert_at},
                                "text": heading_text + subtitle_text}},
                {"updateParagraphStyle": {
                    "range": {"startIndex": h_start, "endIndex": h_end},
                    "paragraphStyle": {"namedStyleType": "TITLE"},
                    "fields": "namedStyleType",
                }},
                {"updateParagraphStyle": {
                    "range": {"startIndex": s_start, "endIndex": s_end},
                    "paragraphStyle": {"namedStyleType": "SUBTITLE"},
                    "fields": "namedStyleType",
                }},
            ]},
        ).execute()
        doc = service.documents().get(documentId=doc_id).execute()
        content = doc.get("body", {}).get("content", [])

    def refresh():
        current = service.documents().get(documentId=doc_id).execute()
        return current, current.get("body", {}).get("content", [])

    def append_paragraph(text_value: str, *, style: str, bullet: bool = False,
                         space_above: bool = False) -> None:
        _, current_content = refresh()
        insert_at = current_content[-1]["endIndex"] - 1
        inserted = "\n" + text_value + "\n"
        start = insert_at + 1
        end = start + _utf16_len(text_value) + 1
        requests = [{
            "insertText": {"location": {"index": insert_at}, "text": inserted}
        }, {
            "updateParagraphStyle": {
                "range": {"startIndex": start, "endIndex": end},
                "paragraphStyle": {"namedStyleType": style},
                "fields": "namedStyleType",
            }
        }]
        if bullet:
            requests.append({
                "createParagraphBullets": {
                    "range": {"startIndex": start, "endIndex": end},
                    "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE",
                }
            })
        if space_above:
            requests.append({
                "updateParagraphStyle": {
                    "range": {"startIndex": start, "endIndex": end},
                    "paragraphStyle": {"spaceAbove": {"magnitude": 10, "unit": "PT"}},
                    "fields": "spaceAbove",
                }
            })
        service.documents().batchUpdate(
            documentId=doc_id, body={"requests": requests}
        ).execute()

    def append_table(rows: list[list[str]], *, space_above: bool = False) -> None:
        current_doc, current_content = refresh()
        insert_at = current_content[-1]["endIndex"] - 1
        if space_above:
            service.documents().batchUpdate(
                documentId=doc_id,
                body={"requests": [{
                    "insertText": {"location": {"index": insert_at}, "text": "\n"}
                }]},
            ).execute()
            _, current_content = refresh()
            insert_at = current_content[-1]["endIndex"] - 1

        service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": [{
                "insertTable": {
                    "rows": len(rows),
                    "columns": len(rows[0]),
                    "location": {"index": insert_at},
                }
            }]},
        ).execute()

        _, current_content = refresh()
        table = next(el["table"] for el in reversed(current_content) if "table" in el)
        requests = []
        for row_index, row in enumerate(rows):
            for col_index, cell_text in enumerate(row):
                cell = table["tableRows"][row_index]["tableCells"][col_index]
                if not cell_text:
                    continue
                cell_insert_at = cell["content"][0]["endIndex"] - 1
                requests.append({
                    "insertText": {
                        "location": {"index": cell_insert_at},
                        "text": cell_text,
                    }
                })
                if row_index == 0:
                    requests.append({
                        "updateTextStyle": {
                            "range": {
                                "startIndex": cell_insert_at,
                                "endIndex": cell_insert_at + _utf16_len(cell_text),
                            },
                            "textStyle": {"bold": True},
                            "fields": "bold",
                        }
                    })
        current_end = current_content[-1]["endIndex"] - 1
        requests.append({
            "insertText": {"location": {"index": current_end}, "text": "\n"}
        })
        service.documents().batchUpdate(
            documentId=doc_id, body={"requests": requests}
        ).execute()

    append_paragraph(heading, style="HEADING_2")
    blocks = _parse_append_blocks(text)
    for block in blocks:
        if block["type"] == "table":
            append_table(block["rows"], space_above=block["space_above"])
        else:
            append_paragraph(
                block["text"],
                style="NORMAL_TEXT",
                bullet=block["type"] == "bullet",
                space_above=block["space_above"],
            )

    return {"status": "appended", "heading": heading, "lines": 1 + len(blocks)}


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

    p_append = sub.add_parser(
        "append", help="Append a headed content section to the document body."
    )
    p_append.add_argument("doc", metavar="DOC_URL")
    p_append.add_argument("heading", metavar="HEADING",
                          help="Section heading (rendered as Heading 2).")
    p_append.add_argument("text", metavar="TEXT",
                          help="Body text. Lines starting with '- ' become bullets.")

    p_note = sub.add_parser(
        "note", help="Append a review note to the Ploma Vermella Review section."
    )
    p_note.add_argument("doc", metavar="DOC_URL")
    p_note.add_argument("quoted_text", metavar="QUOTED_TEXT",
                        help="Exact substring used to determine the note's location.")
    p_note.add_argument("comment", metavar="COMMENT", help="Note text to append.")

    p_epub = sub.add_parser(
        "build-epub",
        help="Build an EPUB from one or more Google Docs.",
    )
    p_epub.add_argument(
        "docs",
        metavar="DOC_URL",
        nargs="+",
        help="One or more Google Doc URLs/IDs to include as chapters.",
    )
    p_epub.add_argument(
        "-o", "--output",
        help="Output EPUB path. Defaults to dist/<slugified-title>-YYYYMMDD.epub.",
    )
    p_epub.add_argument(
        "--title",
        help="Book title for the EPUB metadata and default filename.",
    )

    p_mv = sub.add_parser("mv", help="Move a Google Doc into a Drive folder.")
    p_mv.add_argument("doc", metavar="DOC_URL")
    p_mv.add_argument("folder", metavar="FOLDER_URL")

    p_cp = sub.add_parser("cp", help="Copy a Google Doc into a Drive folder.")
    p_cp.add_argument("doc", metavar="DOC_URL")
    p_cp.add_argument("folder", metavar="FOLDER_URL")
    p_cp.add_argument(
        "--name",
        help="Name for the copy. Defaults to the source document's name.",
    )

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
    elif args.command == "append":
        result = append_content(args.doc, args.heading, args.text)
    elif args.command == "build-epub":
        result = build_epub(args.docs, output=args.output, title=args.title)
    elif args.command == "mv":
        result = move_document(args.doc, args.folder)
    elif args.command == "cp":
        result = copy_document(args.doc, args.folder, name=args.name)
    else:
        result = append_review_note(args.doc, args.quoted_text, args.comment)

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
