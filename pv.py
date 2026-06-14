"""Google Docs API logic — fetch document content and write review notes."""

import argparse
import difflib
import html
import json
import mimetypes
import re
import urllib.request
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from google.auth.transport.requests import AuthorizedSession, Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

CREDENTIALS_DIR = Path(__file__).parent / "credentials"
CLIENT_SECRET = CREDENTIALS_DIR / "client_secret.json"
TOKEN_FILE = CREDENTIALS_DIR / "token.json"

_DOC_URL_RE = re.compile(r"/document/d/([a-zA-Z0-9_-]+)")
_FOLDER_URL_RE = re.compile(r"/folders/([a-zA-Z0-9_-]+)")
_PRESENTATION_URL_RE = re.compile(r"/presentation/d/([a-zA-Z0-9_-]+)")
_SPREADSHEET_URL_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")


def _extract_doc_id(doc_id_or_url: str) -> str:
    """Return bare doc ID whether given a full URL or already a bare ID."""
    m = _DOC_URL_RE.search(doc_id_or_url)
    return m.group(1) if m else doc_id_or_url.strip()


def _extract_presentation_id(presentation_id_or_url: str) -> str:
    """Return bare presentation ID whether given a full URL or already bare."""
    m = _PRESENTATION_URL_RE.search(presentation_id_or_url)
    return m.group(1) if m else presentation_id_or_url.strip()


def _extract_spreadsheet_id(spreadsheet_id_or_url: str) -> str:
    """Return bare spreadsheet ID whether given a full URL or already bare."""
    m = _SPREADSHEET_URL_RE.search(spreadsheet_id_or_url)
    return m.group(1) if m else spreadsheet_id_or_url.strip()


def _review_copy_title(
    title: str,
    stamp: datetime | None = None,
    suffix_template: str | None = None,
) -> str:
    """Return a dated review-copy title."""
    stamp = stamp or datetime.now()
    suffix_template = suffix_template or " - DRAFT {date}"
    return f"{title}{suffix_template.format(date=stamp.strftime('%Y-%m-%d'))}"


def _get_credentials() -> Credentials:
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET), SCOPES)
            creds = flow.run_local_server(port=0, prompt="select_account")
        CREDENTIALS_DIR.mkdir(parents=True, exist_ok=True)
        TOKEN_FILE.write_text(creds.to_json())
    return creds


def _docs_service():
    return build("docs", "v1", credentials=_get_credentials())


def _drive_service():
    return build("drive", "v3", credentials=_get_credentials())


def _slides_service():
    return build("slides", "v1", credentials=_get_credentials())


def _sheets_service():
    return build("sheets", "v4", credentials=_get_credentials())


_HEADING_STYLES = {"HEADING_1", "HEADING_2", "HEADING_3", "HEADING_4", "TITLE"}
_REVIEW_HEADING = "🪶 Ploma Vermella Review"
_EPUB_NS = "http://www.idpf.org/2007/ops"


def _utf16_len(s: str) -> int:
    """Length of s in UTF-16 code units (the unit the Docs API uses for indices)."""
    return sum(2 if ord(c) > 0xFFFF else 1 for c in s)


# Curly quotes/apostrophes map to their straight equivalents. Each is a single
# BMP character, so normalising is position-preserving — safe to use when
# locating text without disturbing document indices.
_QUOTE_MAP = str.maketrans({
    "‘": "'", "’": "'", "‚": "'", "‛": "'", "′": "'",
    "“": '"', "”": '"', "„": '"', "‟": '"', "″": '"',
})


def _normalize_quotes(s: str) -> str:
    """Fold curly quotes/apostrophes to straight ones so matching is quote-agnostic."""
    return s.translate(_QUOTE_MAP)


def _parse_table_row(line: str) -> list[str]:
    """Parse a markdown-style pipe table row into stripped cell values."""
    text = line.strip()
    if not (text.startswith("|") and text.endswith("|")):
        raise ValueError(f"not a pipe table row: {line!r}")
    return [cell.strip() for cell in text[1:-1].split("|")]


def _shape_text(element: dict) -> str:
    """Return concatenated visible text from a Google Slides shape element."""
    shape = element.get("shape", {})
    runs = []
    for text_element in shape.get("text", {}).get("textElements", []):
        text_run = text_element.get("textRun")
        if text_run:
            runs.append(text_run.get("content", ""))
    return "".join(runs).strip()


def _paragraph_text(element: dict) -> str:
    """Return concatenated paragraph text for a body element."""
    paragraph = element.get("paragraph", {})
    return "".join(
        pe.get("textRun", {}).get("content", "")
        for pe in paragraph.get("elements", [])
    )


def _is_image_paragraph(element: dict) -> bool:
    """Return True when a body element paragraph contains an inline image."""
    paragraph = element.get("paragraph", {})
    return any(pe.get("inlineObjectElement") for pe in paragraph.get("elements", []))


def _figure_map_from_doc(doc: dict) -> list[dict]:
    """Return inline-image neighborhoods keyed by body element index."""
    content = doc.get("body", {}).get("content", [])
    figures = []
    for index, element in enumerate(content):
        if not _is_image_paragraph(element):
            continue
        prev_text = ""
        next_text = ""
        caption_text = ""

        if index > 0:
            prev_text = _paragraph_text(content[index - 1]).strip()

        scan = index + 1
        while scan < len(content):
            candidate = _paragraph_text(content[scan]).strip()
            if candidate:
                next_text = candidate
                break
            scan += 1

        if next_text.startswith("Figure ") or next_text.startswith("Listing "):
            caption_text = next_text
            scan += 1
            next_text = ""
            while scan < len(content):
                candidate = _paragraph_text(content[scan]).strip()
                if candidate:
                    next_text = candidate
                    break
                scan += 1

        figures.append({
            "body_index": index,
            "start_index": element.get("startIndex"),
            "end_index": element.get("endIndex"),
            "prev_text": prev_text,
            "caption_text": caption_text,
            "next_text": next_text,
        })
    return figures


# Font families treated as code/monospace, so safe-edit tooling can skip them.
_MONOSPACE_FONTS = frozenset({
    "Courier New", "Consolas", "Roboto Mono", "Source Code Pro", "Inconsolata",
    "Fira Code", "JetBrains Mono", "Menlo", "Monaco", "Cousine", "PT Mono",
    "Ubuntu Mono", "IBM Plex Mono", "Space Mono", "Anonymous Pro",
})


def _is_code_paragraph(element: dict) -> bool:
    """True when a body element's visible text is entirely in a monospace font."""
    paragraph = element.get("paragraph")
    if not paragraph:
        return False
    families = []
    for pe in paragraph.get("elements", []):
        text_run = pe.get("textRun")
        if not text_run or not text_run.get("content", "").strip():
            continue
        families.append(
            text_run.get("textStyle", {}).get("weightedFontFamily", {}).get("fontFamily")
        )
    return bool(families) and all(f in _MONOSPACE_FONTS for f in families)


def _doc_text_runs(element_source: dict) -> list[tuple[int, str]]:
    """Return [(doc_start_index, text)] for every text run in the document body."""
    runs = []
    for el in element_source.get("body", {}).get("content", []):
        paragraph = el.get("paragraph")
        if not paragraph:
            continue
        for pe in paragraph.get("elements", []):
            text_run = pe.get("textRun")
            if text_run is not None:
                runs.append((pe["startIndex"], text_run.get("content", "")))
    return runs


def _doc_index_at(runs: list[tuple[int, str]], flat_pos: int) -> int:
    """Map a position in the flattened body text to a document index."""
    pos = 0
    for start, text in runs:
        if pos + len(text) > flat_pos:
            return start + (flat_pos - pos)
        pos += len(text)
    raise IndexError(f"flat position {flat_pos} out of range")


def _body_element_at(content: list[dict], doc_index: int) -> tuple[int, dict | None]:
    """Return (ordinal, element) of the body element containing a document index."""
    for i, el in enumerate(content):
        start, end = el.get("startIndex"), el.get("endIndex")
        if start is None or end is None:
            continue
        if start <= doc_index < end:
            return i, el
    return -1, None


def _find_matches(doc: dict, text: str) -> list[dict]:
    """Locate every occurrence of `text` in the body. Returns [] when none."""
    if not text:
        raise ValueError("search text must not be empty")
    content = doc.get("body", {}).get("content", [])
    runs = _doc_text_runs(doc)
    flat = "".join(t for _, t in runs)
    # Match on quote-normalized text; normalization is position-preserving.
    nflat = _normalize_quotes(flat)
    ntext = _normalize_quotes(text)
    matches = []
    pos = nflat.find(ntext)
    while pos != -1:
        start_index = _doc_index_at(runs, pos)
        body_index, el = _body_element_at(content, start_index)
        matches.append({
            "flat_pos": pos,
            "start_index": start_index,
            "end_index": start_index + _utf16_len(text),
            "body_index": body_index,
            "paragraph_style": (
                el.get("paragraph", {}).get("paragraphStyle", {}).get("namedStyleType", "")
                if el else ""
            ),
            "is_code": _is_code_paragraph(el) if el else False,
            "context": _paragraph_text(el).strip() if el else "",
        })
        pos = nflat.find(ntext, pos + 1)
    return matches


def _insert_after_plan(
    doc: dict, anchor: str, text: str, require_unique: bool = True
) -> tuple[dict, int]:
    """Build the insertText request that places `text` after the anchor's paragraph."""
    if not anchor:
        raise ValueError("anchor must not be empty")
    content = doc.get("body", {}).get("content", [])
    nanchor = _normalize_quotes(anchor)
    hits = [
        (i, el) for i, el in enumerate(content)
        if el.get("paragraph") and nanchor in _normalize_quotes(_paragraph_text(el))
    ]
    if not hits:
        raise ValueError(f"anchor not found: {anchor!r}")
    if len(hits) > 1 and require_unique:
        raise ValueError(
            f"anchor matches {len(hits)} paragraphs; make it unique or pass allow_multiple"
        )
    body_index, el = hits[0]
    # Insert just before the paragraph's terminating newline so the new
    # paragraph(s) inherit the anchor paragraph's style.
    insert_index = el["endIndex"] - 1
    request = {"insertText": {"location": {"index": insert_index}, "text": "\n" + text}}
    return request, body_index


def _parse_hex_color(hex_color: str) -> dict:
    """Parse '#rrggbb' (or 'rrggbb') into a Docs API rgbColor (floats 0..1)."""
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise ValueError(f"expected a 6-digit hex color, got {hex_color!r}")
    return {
        "red": int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue": int(h[4:6], 16) / 255,
    }


def _style_plan(
    doc: dict,
    text: str,
    italic: bool = False,
    bold: bool = False,
    underline: bool = False,
    color: str | None = None,
    all_occurrences: bool = False,
) -> tuple[list[dict], list[dict]]:
    """Build updateTextStyle requests applying character styles to `text`."""
    style: dict = {}
    fields: list[str] = []
    if italic:
        style["italic"] = True
        fields.append("italic")
    if bold:
        style["bold"] = True
        fields.append("bold")
    if underline:
        style["underline"] = True
        fields.append("underline")
    if color:
        style["foregroundColor"] = {"color": {"rgbColor": _parse_hex_color(color)}}
        fields.append("foregroundColor")
    if not fields:
        raise ValueError("specify at least one of italic, bold, underline, or color")

    matches = _find_matches(doc, text)
    if not matches:
        raise ValueError(f"text not found: {text!r}")
    if len(matches) > 1 and not all_occurrences:
        raise ValueError(
            f"text matches {len(matches)} times; make it unique or pass all_occurrences"
        )
    targets = matches if all_occurrences else matches[:1]
    requests = [
        {"updateTextStyle": {
            "range": {"startIndex": m["start_index"], "endIndex": m["end_index"]},
            "textStyle": style,
            "fields": ",".join(fields),
        }}
        for m in targets
    ]
    spans = [{"start_index": m["start_index"], "end_index": m["end_index"]} for m in targets]
    return requests, spans


def _link_plan(
    doc: dict, text: str, url: str, all_occurrences: bool = False
) -> tuple[list[dict], list[dict]]:
    """Build updateTextStyle requests that hyperlink occurrences of `text`."""
    matches = _find_matches(doc, text)
    if not matches:
        raise ValueError(f"text not found: {text!r}")
    if len(matches) > 1 and not all_occurrences:
        raise ValueError(
            f"text matches {len(matches)} times; make it unique or pass all_occurrences"
        )
    targets = matches if all_occurrences else matches[:1]
    requests = [
        {"updateTextStyle": {
            "range": {"startIndex": m["start_index"], "endIndex": m["end_index"]},
            "textStyle": {"link": {"url": url}},
            "fields": "link",
        }}
        for m in targets
    ]
    spans = [
        {"start_index": m["start_index"], "end_index": m["end_index"]}
        for m in targets
    ]
    return requests, spans


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


def _render_para_elements(elements: list[dict]) -> str:
    """Render paragraph text runs to markdown-ish text (links, strikethrough preserved)."""
    parts = []
    for pe in elements:
        text_run = pe.get("textRun")
        if not text_run:
            continue
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


def _render_table(table: dict) -> str:
    """Render a Docs table as text: one row per line, cells joined with ' | '."""
    lines = []
    for row in table.get("tableRows", []):
        cells = [
            "".join(
                _render_para_elements(ce.get("paragraph", {}).get("elements", []))
                for ce in cell.get("content", [])
            ).strip()
            for cell in row.get("tableCells", [])
        ]
        lines.append(" | ".join(cells))
    return "".join(line + "\n" for line in lines)


def _extract_text(doc: dict) -> str:
    """Extract plain text from a Google Docs document body, stopping before the review section.

    Table cells are rendered row by row (cells joined with ' | ') so tabular content
    is included in the fetched text rather than silently dropped.
    """
    parts = []
    body = doc.get("body", {})
    for element in body.get("content", []):
        table = element.get("table")
        if table:
            parts.append(_render_table(table))
            continue
        paragraph = element.get("paragraph")
        if not paragraph:
            continue
        text = "".join(
            pe.get("textRun", {}).get("content", "")
            for pe in paragraph.get("elements", [])
        ).rstrip("\n")
        if text == _REVIEW_HEADING:
            break
        parts.append(_render_para_elements(paragraph.get("elements", [])))
    return "".join(parts)


def _text_from_elements(elements: list[dict]) -> str:
    """Extract plain paragraph text from Docs API paragraph elements."""
    return "".join(
        pe.get("textRun", {}).get("content", "")
        for pe in elements
    ).rstrip("\n")


def _inline_html(elements: list[dict]) -> str:
    """Render paragraph text runs as inline HTML, preserving italic/bold/links."""
    parts = []
    for pe in elements:
        text_run = pe.get("textRun")
        if text_run is None:
            continue
        content = text_run.get("content", "").replace("\n", "")
        if content == "":
            continue
        fragment = html.escape(content)
        style = text_run.get("textStyle", {})
        if style.get("italic"):
            fragment = f"<em>{fragment}</em>"
        if style.get("bold"):
            fragment = f"<strong>{fragment}</strong>"
        link_url = style.get("link", {}).get("url")
        if link_url:
            fragment = f'<a href="{html.escape(link_url, quote=True)}">{fragment}</a>'
        parts.append(fragment)
    return "".join(parts).strip()


def _inline_object_ids(element: dict) -> list[str]:
    """Return the inline object IDs (images) referenced by a body element."""
    ids = []
    for pe in element.get("paragraph", {}).get("elements", []):
        ioe = pe.get("inlineObjectElement")
        if ioe and ioe.get("inlineObjectId"):
            ids.append(ioe["inlineObjectId"])
    return ids


def _extract_blocks(doc: dict) -> list[dict]:
    """
    Extract document blocks suitable for EPUB rendering.

    Returns a list of dictionaries with type keys:
    - {"type": "heading", "level": 1..4, "text": "...", "html": "..."}
    - {"type": "paragraph", "text": "...", "html": "..."}
    - {"type": "list_item", "text": "...", "html": "..."}
    - {"type": "image", "object_id": "..."}
    """
    blocks = []
    for element in doc.get("body", {}).get("content", []):
        paragraph = element.get("paragraph")
        if not paragraph:
            continue

        text = _text_from_elements(paragraph.get("elements", []))
        if text == _REVIEW_HEADING:
            break

        # Image-only paragraphs carry no text; emit them before the empty check
        # so figures are preserved rather than dropped.
        for object_id in _inline_object_ids(element):
            blocks.append({"type": "image", "object_id": object_id})

        if not text:
            continue

        inline = _inline_html(paragraph.get("elements", []))
        style = paragraph.get("paragraphStyle", {}).get("namedStyleType", "")
        if "bullet" in paragraph:
            blocks.append({"type": "list_item", "text": text, "html": inline})
            continue

        if style == "TITLE":
            blocks.append({"type": "heading", "level": 1, "text": text, "html": inline})
        elif style == "HEADING_1":
            blocks.append({"type": "heading", "level": 2, "text": text, "html": inline})
        elif style == "HEADING_2":
            blocks.append({"type": "heading", "level": 3, "text": text, "html": inline})
        elif style in {"HEADING_3", "HEADING_4"}:
            blocks.append({"type": "heading", "level": 4, "text": text, "html": inline})
        else:
            blocks.append({"type": "paragraph", "text": text, "html": inline})

    return blocks


def _slugify(value: str) -> str:
    """Return a filesystem-friendly ASCII-ish slug."""
    slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return slug or "book"


def _chapter_filename(index: int) -> str:
    """Return a stable EPUB filename for a chapter."""
    return f"chapter-{index:02d}.xhtml"


_MEDIA_EXTENSIONS = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/gif": "gif",
    "image/webp": "webp",
    "image/svg+xml": "svg",
}


def _media_extension(media_type: str) -> str:
    """Map an image media type to a file extension."""
    key = (media_type or "").split(";")[0].strip().lower()
    return _MEDIA_EXTENSIONS.get(key, "img")


def _image_content_uri(doc: dict, object_id: str) -> str | None:
    """Return the (short-lived) content URI for an inline image object, if any."""
    obj = doc.get("inlineObjects", {}).get(object_id, {})
    embedded = obj.get("inlineObjectProperties", {}).get("embeddedObject", {})
    return embedded.get("imageProperties", {}).get("contentUri")


def _download_image(content_uri: str) -> tuple[bytes, str]:
    """Download an inline image, returning (bytes, media_type)."""
    session = AuthorizedSession(_get_credentials())
    resp = session.get(content_uri)
    resp.raise_for_status()
    media_type = (resp.headers.get("Content-Type") or "image/png").split(";")[0].strip()
    return resp.content, media_type


def _downscale_image(data: bytes, media_type: str, max_width: int) -> tuple[bytes, str]:
    """Shrink an image for EPUB embedding; return (bytes, media_type).

    Downscales to max_width, then encodes as whichever of PNG or JPEG is smaller —
    PNG wins for flat diagrams (and stays crisp), JPEG wins for photographs and
    screenshots. Transparency forces PNG. Undecodable data, and images that do not
    get smaller, are returned unchanged.
    """
    import io

    from PIL import Image
    try:
        img = Image.open(io.BytesIO(data))
        width, height = img.size
        has_alpha = img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info)

        resized = width > max_width
        if resized:
            base = img.convert("RGBA") if has_alpha else img.convert("RGB")
            new_height = max(1, round(height * max_width / width))
            img = base.resize((max_width, new_height), Image.LANCZOS)

        candidates = []
        png_src = img.convert("RGBA" if has_alpha else "RGB") if img.mode == "P" else img
        png_buf = io.BytesIO()
        png_src.save(png_buf, format="PNG", optimize=True)
        candidates.append((png_buf.getvalue(), "image/png"))
        if not has_alpha:
            jpg_buf = io.BytesIO()
            img.convert("RGB").save(jpg_buf, format="JPEG", quality=82, optimize=True)
            candidates.append((jpg_buf.getvalue(), "image/jpeg"))

        best_bytes, best_mt = min(candidates, key=lambda c: len(c[0]))
        if resized or len(best_bytes) < len(data):
            return best_bytes, best_mt
        return data, media_type
    except Exception:
        return data, media_type


def _block_html(block: dict) -> str:
    """Inner HTML for a text block: prefer rich inline html, else escape text."""
    inline = block.get("html")
    return inline if inline is not None else html.escape(block.get("text", ""))


def _blocks_to_xhtml(title: str, blocks: list[dict], image_paths: dict | None = None) -> str:
    """Render extracted blocks as a simple XHTML chapter.

    image_paths maps an inline object_id to its EPUB-relative href; image blocks
    whose object_id is absent from the map are skipped.
    """
    image_paths = image_paths or {}
    parts = []
    in_list = False
    for block in blocks:
        block_type = block["type"]

        if block_type == "list_item":
            if not in_list:
                parts.append("<ul>")
                in_list = True
            parts.append(f"<li>{_block_html(block)}</li>")
            continue

        if in_list:
            parts.append("</ul>")
            in_list = False

        if block_type == "image":
            href = image_paths.get(block.get("object_id"))
            if href:
                src = html.escape(href, quote=True)
                parts.append(f'<figure><img src="{src}" alt=""/></figure>')
        elif block_type == "heading":
            level = max(1, min(6, int(block["level"])))
            parts.append(f"<h{level}>{_block_html(block)}</h{level}>")
        else:
            parts.append(f"<p>{_block_html(block)}</p>")

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


def _title_page_xhtml(title: str, subtitle: str | None = None, author: str | None = None) -> str:
    """Render a simple title page as an XHTML document."""
    parts = [f'<h1 class="title">{html.escape(title)}</h1>']
    if subtitle:
        parts.append(f'<p class="subtitle">{html.escape(subtitle)}</p>')
    if author:
        parts.append(f'<p class="author">{html.escape(author)}</p>')
    body = "\n    ".join(parts)
    return (
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
        "<html xmlns=\"http://www.w3.org/1999/xhtml\" "
        f"xmlns:epub=\"{_EPUB_NS}\">\n"
        "<head>\n"
        f"  <title>{html.escape(title)}</title>\n"
        "  <link rel=\"stylesheet\" type=\"text/css\" href=\"styles.css\"/>\n"
        "</head>\n"
        "<body>\n"
        f"  <section epub:type=\"titlepage\" class=\"titlepage\">\n    {body}\n  </section>\n"
        "</body>\n"
        "</html>\n"
    )


def _cover_page_xhtml(image_href: str) -> str:
    """Render a full-page cover image as an XHTML document."""
    src = html.escape(image_href, quote=True)
    return (
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
        "<html xmlns=\"http://www.w3.org/1999/xhtml\" "
        f"xmlns:epub=\"{_EPUB_NS}\">\n"
        "<head>\n"
        "  <title>Cover</title>\n"
        "  <link rel=\"stylesheet\" type=\"text/css\" href=\"styles.css\"/>\n"
        "</head>\n"
        "<body>\n"
        f"  <section epub:type=\"cover\">\n    <img class=\"cover\" src=\"{src}\" alt=\"Cover\"/>\n"
        "  </section>\n"
        "</body>\n"
        "</html>\n"
    )


def _read_cover_image(cover: str) -> tuple[bytes, str]:
    """Return (bytes, media_type) for a cover image given a local path or URL."""
    if cover.startswith(("http://", "https://")):
        with urllib.request.urlopen(cover) as resp:
            return resp.read(), (resp.headers.get_content_type() or "image/jpeg")
    path = Path(cover)
    return path.read_bytes(), (mimetypes.guess_type(path.name)[0] or "image/jpeg")


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


def _epub_package(
    book_title: str,
    book_id: str,
    chapters: list[dict],
    media_items: list[dict] | None = None,
    author: str | None = None,
    front_matter: list[dict] | None = None,
    cover_image_id: str | None = None,
) -> str:
    """Return the OPF package document.

    front_matter is a list of {id, href} XHTML docs (cover/title pages) placed
    ahead of the chapters in the spine. cover_image_id names the media item to
    flag as the EPUB cover image.
    """
    manifest_items = [
        '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
        '<item id="css" href="styles.css" media-type="text/css"/>',
    ]
    for media in media_items or []:
        props = ' properties="cover-image"' if media["id"] == cover_image_id else ""
        manifest_items.append(
            f'<item id="{media["id"]}" href="{media["href"]}" '
            f'media-type="{media["media_type"]}"{props}/>'
        )

    spine_items = []
    for page in front_matter or []:
        manifest_items.append(
            f'<item id="{page["id"]}" href="{page["href"]}" media-type="application/xhtml+xml"/>'
        )
        spine_items.append(f'<itemref idref="{page["id"]}"/>')
    for i, chapter in enumerate(chapters, start=1):
        manifest_items.append(
            f'<item id="chap{i}" href="{chapter["filename"]}" media-type="application/xhtml+xml"/>'
        )
        spine_items.append(f'<itemref idref="chap{i}"/>')

    manifest = "\n    ".join(manifest_items)
    spine = "\n    ".join(spine_items)
    modified = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    meta = [
        f'<dc:identifier id="bookid">urn:uuid:{book_id}</dc:identifier>',
        f"<dc:title>{html.escape(book_title)}</dc:title>",
    ]
    if author:
        meta.append(f"<dc:creator>{html.escape(author)}</dc:creator>")
    meta.append("<dc:language>en</dc:language>")
    meta.append(f'<meta property="dcterms:modified">{modified}</meta>')
    if cover_image_id:
        meta.append(f'<meta name="cover" content="{cover_image_id}"/>')
    metadata = "\n    ".join(meta)

    return (
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
        "<package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" "
        f'unique-identifier="bookid" xml:lang="en">\n'
        "  <metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\n"
        f"    {metadata}\n"
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


def fetch_presentation(presentation_id_or_url: str) -> dict:
    """Return slide text content from a Google Slides presentation."""
    presentation_id = _extract_presentation_id(presentation_id_or_url)
    presentation = _slides_service().presentations().get(
        presentationId=presentation_id
    ).execute()
    slides = []
    for index, slide in enumerate(presentation.get("slides", []), start=1):
        texts = []
        for element in slide.get("pageElements", []):
            text = _shape_text(element)
            if text:
                texts.append(text)
        slides.append({
            "index": index,
            "object_id": slide.get("objectId", ""),
            "speaker_notes_object_id": slide.get("slideProperties", {})
            .get("notesPage", {})
            .get("objectId", ""),
            "texts": texts,
        })
    return {
        "title": presentation.get("title", ""),
        "slides": slides,
    }


def presentation_thumbnail(
    presentation_id_or_url: str,
    page_object_id: str,
    size: str = "LARGE",
) -> dict:
    """Return a Google Slides page thumbnail URL."""
    presentation_id = _extract_presentation_id(presentation_id_or_url)
    result = _slides_service().presentations().pages().getThumbnail(
        presentationId=presentation_id,
        pageObjectId=page_object_id,
        thumbnailProperties_thumbnailSize=size,
    ).execute()
    return {
        "presentation_id": presentation_id,
        "page_object_id": page_object_id,
        "size": size,
        "content_url": result.get("contentUrl", ""),
        "width": result.get("width"),
        "height": result.get("height"),
    }


def fetch_sheet(spreadsheet_id_or_url: str, range_name: str | None = None) -> dict:
    """Return sheet metadata and values for an optional range."""
    spreadsheet_id = _extract_spreadsheet_id(spreadsheet_id_or_url)
    service = _sheets_service()
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    result = {
        "title": spreadsheet.get("properties", {}).get("title", ""),
        "spreadsheet_id": spreadsheet_id,
        "sheets": [
            {
                "title": sheet.get("properties", {}).get("title", ""),
                "sheet_id": sheet.get("properties", {}).get("sheetId"),
                "index": sheet.get("properties", {}).get("index"),
            }
            for sheet in spreadsheet.get("sheets", [])
        ],
    }
    if range_name:
        values = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name,
        ).execute()
        result["range"] = values.get("range", range_name)
        result["values"] = values.get("values", [])
    return result


def update_sheet(
    spreadsheet_id_or_url: str,
    range_name: str,
    values: list[list[str]],
) -> dict:
    """Write a rectangular value matrix to a Google Sheet range."""
    spreadsheet_id = _extract_spreadsheet_id(spreadsheet_id_or_url)
    updated = _sheets_service().spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()
    return {
        "status": "updated",
        "spreadsheet_id": spreadsheet_id,
        "updated_range": updated.get("updatedRange", range_name),
        "updated_rows": updated.get("updatedRows", 0),
        "updated_columns": updated.get("updatedColumns", 0),
        "updated_cells": updated.get("updatedCells", 0),
    }


def figure_map(doc_id_or_url: str) -> dict:
    """Return inline-image neighborhoods for a Google Doc."""
    doc_id = _extract_doc_id(doc_id_or_url)
    doc = _docs_service().documents().get(documentId=doc_id).execute()
    return {
        "title": doc.get("title", ""),
        "figures": _figure_map_from_doc(doc),
    }


def replace_body_range(
    doc_id_or_url: str,
    start_body_index: int,
    end_body_index: int,
    text: str,
) -> dict:
    """
    Replace a contiguous body-element range with text.

    Body indexes are looked up fresh at execution time, which is safer than
    carrying raw document indices across multiple edits.
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    doc = _docs_service().documents().get(documentId=doc_id).execute()
    content = doc.get("body", {}).get("content", [])
    if start_body_index < 0 or end_body_index < start_body_index:
        raise ValueError("invalid body index range")
    if end_body_index >= len(content):
        raise ValueError(f"body index {end_body_index} out of range")

    start_index = content[start_body_index].get("startIndex")
    end_index = content[end_body_index].get("endIndex")
    requests = [{
        "deleteContentRange": {
            "range": {"startIndex": start_index, "endIndex": end_index}
        }
    }]
    if text:
        requests.append({
            "insertText": {"location": {"index": start_index}, "text": text}
        })
    _docs_service().documents().batchUpdate(
        documentId=doc_id,
        body={"requests": requests},
    ).execute()
    return {
        "status": "replaced",
        "start_body_index": start_body_index,
        "end_body_index": end_body_index,
        "text": text,
    }


def insert_image_at_body_index(
    doc_id_or_url: str,
    body_index: int,
    image_url: str,
    width_pt: float = 468.0,
    height_pt: float = 263.25,
) -> dict:
    """Insert an inline image at the start index of the given body element."""
    doc_id = _extract_doc_id(doc_id_or_url)
    doc = _docs_service().documents().get(documentId=doc_id).execute()
    content = doc.get("body", {}).get("content", [])
    if body_index < 0 or body_index >= len(content):
        raise ValueError(f"body index {body_index} out of range")
    location = content[body_index].get("startIndex")
    _docs_service().documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{
            "insertInlineImage": {
                "location": {"index": location},
                "uri": image_url,
                "objectSize": {
                    "width": {"magnitude": width_pt, "unit": "PT"},
                    "height": {"magnitude": height_pt, "unit": "PT"},
                },
            }
        }]},
    ).execute()
    return {
        "status": "inserted_image",
        "body_index": body_index,
        "image_url": image_url,
        "width_pt": width_pt,
        "height_pt": height_pt,
    }


def _map_comments(raw_comments: list[dict], include_resolved: bool) -> list[dict]:
    """Flatten raw Drive comments (with replies) to structured dicts."""
    out = []
    for c in raw_comments:
        if not include_resolved and c.get("resolved", False):
            continue
        replies = [
            {
                "id": r.get("id", ""),
                "author": r.get("author", {}).get("displayName", ""),
                "content": r.get("content", ""),
            }
            for r in c.get("replies", [])
            if not r.get("deleted", False)
        ]
        out.append({
            "id": c.get("id", ""),
            "author": c.get("author", {}).get("displayName", ""),
            "content": c.get("content", ""),
            "quoted_text": c.get("quotedFileContent", {}).get("value", ""),
            "resolved": c.get("resolved", False),
            "replies": replies,
        })
    return out


def _fetch_comments(doc_id: str, include_resolved: bool = False) -> list[dict]:
    """Page through and return a doc's comments as flat dicts."""
    service = _drive_service()
    raw_comments: list[dict] = []
    page_token = None
    while True:
        kwargs: dict = {
            "fileId": doc_id,
            "fields": (
                "comments(id,author,content,quotedFileContent,resolved,"
                "replies(id,author,content,deleted)),"
                "nextPageToken"
            ),
            "includeDeleted": False,
            "pageSize": 100,
        }
        if page_token:
            kwargs["pageToken"] = page_token
        result = service.comments().list(**kwargs).execute()
        raw_comments.extend(result.get("comments", []))
        page_token = result.get("nextPageToken")
        if not page_token:
            break
    return _map_comments(raw_comments, include_resolved)


def list_comments(doc_id_or_url: str, include_resolved: bool = False) -> dict:
    """Return a doc's comments (id, author, content, quoted_text, resolved)."""
    doc_id = _extract_doc_id(doc_id_or_url)
    comments = _fetch_comments(doc_id, include_resolved)
    return {"count": len(comments), "comments": comments}


def fetch_document(doc_id_or_url: str, include_resolved: bool = False) -> dict:
    """
    Return {title, text, comments} for the given Google Doc.

    By default, resolved comments are filtered out. Pass include_resolved=True
    to include them; each comment carries a `resolved` boolean either way.
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    doc = _docs_service().documents().get(documentId=doc_id).execute()
    return {
        "title": doc.get("title", ""),
        "text": _extract_text(doc),
        "comments": _fetch_comments(doc_id, include_resolved),
    }


def resolve_comment(doc_id_or_url: str, comment_id: str) -> dict:
    """Resolve a comment by posting a reply with action=resolve."""
    doc_id = _extract_doc_id(doc_id_or_url)
    _drive_service().replies().create(
        fileId=doc_id,
        commentId=comment_id,
        body={"action": "resolve"},
        fields="id,action",
    ).execute()
    return {"status": "resolved", "comment_id": comment_id}


def resolve_all_comments(doc_id_or_url: str) -> dict:
    """Resolve every unresolved comment on the document."""
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _drive_service()
    listing = service.comments().list(
        fileId=doc_id,
        fields="comments(id,resolved)",
        includeDeleted=False,
    ).execute()
    unresolved = [c["id"] for c in listing.get("comments", []) if not c.get("resolved")]
    for comment_id in unresolved:
        service.replies().create(
            fileId=doc_id,
            commentId=comment_id,
            body={"action": "resolve"},
            fields="id",
        ).execute()
    return {"status": "resolved_all", "count": len(unresolved), "comment_ids": unresolved}


def comment_document(doc_id_or_url: str, quoted_text: str, comment: str) -> dict:
    """Create an anchored sidebar comment with quoted-text context."""
    doc_id = _extract_doc_id(doc_id_or_url)
    created = _drive_service().comments().create(
        fileId=doc_id,
        body={
            "content": comment,
            "quotedFileContent": {"mimeType": "text/plain", "value": quoted_text},
        },
        fields="id,content,quotedFileContent",
    ).execute()
    return {
        "status": "created",
        "id": created.get("id"),
        "content": created.get("content", ""),
        "quoted_text": created.get("quotedFileContent", {}).get("value", ""),
    }


def _ambiguous(
    reason: str,
    message: str,
    *,
    options: list | None = None,
    resolution: dict | None = None,
    question: str | None = None,
) -> dict:
    """Build a 'needs disambiguation' result: PV could act but won't choose for you."""
    out: dict = {"status": "ambiguous", "reason": reason, "message": message}
    if question is not None:
        out["question"] = question
    if options is not None:
        out["options"] = options
    if resolution is not None:
        out["resolution"] = resolution
    return out


def _occurrence_options(flat: str, positions: list[int], length: int) -> list[dict]:
    """One option per match, with surrounding context, for the caller to choose from."""
    options = []
    for n, pos in enumerate(positions, 1):
        ctx = flat[max(0, pos - 40):pos + length + 40].replace("\n", " ").strip()
        options.append({"id": n, "flat_pos": pos, "context": f"...{ctx}..."})
    return options


def _fuzzy_candidates(flat: str, nflat: str, nold: str, min_len: int = 8) -> list[dict]:
    """Best partial match for a no-match case, so the caller can recover from drift."""
    if not nflat or not nold:
        return []
    matcher = difflib.SequenceMatcher(None, nflat, nold, autojunk=False)
    block = matcher.find_longest_match(0, len(nflat), 0, len(nold))
    if block.size < min_len:
        return []
    pos = block.a
    ctx = flat[max(0, pos - 40):pos + block.size + 40].replace("\n", " ").strip()
    return [{"matched_chars": block.size, "context": f"...{ctx}..."}]


def _plan_edit_matches(
    flat: str,
    nflat: str,
    old: str,
    nold: str,
    all_occurrences: bool,
    occurrence: int | None,
) -> dict:
    """Decide which match(es) to edit, or report an ambiguous result.

    Pure (no API calls). Returns {"kind": "ok", "positions": [...]} or
    {"kind": "ambiguous", "result": <ambiguous payload>}.
    """
    positions = []
    i = nflat.find(nold)
    while i != -1:
        positions.append(i)
        i = nflat.find(nold, i + len(nold))

    if not positions:
        return {"kind": "ambiguous", "result": _ambiguous(
            "no_match",
            f"No match for {old!r}; the text may have changed since the doc was read.",
            options=_fuzzy_candidates(flat, nflat, nold),
            resolution={
                "how": "re_call_with",
                "field": "old",
                "hint": "re-fetch the doc and anchor on its current text",
            },
        )}

    if occurrence is not None:
        if occurrence < 1 or occurrence > len(positions):
            return {"kind": "ambiguous", "result": _ambiguous(
                "occurrence_out_of_range",
                f"occurrence {occurrence} is out of range; there are {len(positions)} matches.",
                options=_occurrence_options(flat, positions, len(old)),
                resolution={"how": "re_call_with", "field": "occurrence"},
            )}
        return {"kind": "ok", "positions": [positions[occurrence - 1]]}

    if len(positions) > 1 and not all_occurrences:
        return {"kind": "ambiguous", "result": _ambiguous(
            "multiple_matches",
            f"{len(positions)} matches for {old!r}; PV will not choose which.",
            question="Which occurrence did you mean?",
            options=_occurrence_options(flat, positions, len(old)),
            resolution={
                "how": "re_call_with",
                "field": "occurrence",
                "example": "pv edit <doc> <old> <new> --occurrence 2",
            },
        )}

    return {"kind": "ok", "positions": positions}


def edit_document(
    doc_id_or_url: str,
    old: str,
    new: str,
    all_occurrences: bool = False,
    occurrence: int | None = None,
) -> dict:
    """
    Replace text in a Google Doc.

    Requires a single match for `old` by default. When `old` matches several
    places, returns an 'ambiguous' result listing the occurrences (pass
    all_occurrences=True to replace all, or occurrence=N to pick one). When it
    matches nothing, returns an 'ambiguous' result with the closest partial
    match. Edits apply in reverse-index order so earlier inserts don't shift
    later positions.
    """
    if not old:
        raise ValueError("old must not be empty")
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()
    doc = service.documents().get(documentId=doc_id).execute()

    runs = []  # (utf16_start_in_doc, text)
    for el in doc.get("body", {}).get("content", []):
        para = el.get("paragraph")
        if not para:
            continue
        for pe in para.get("elements", []):
            tr = pe.get("textRun")
            if tr:
                runs.append((pe["startIndex"], tr.get("content", "")))

    flat = "".join(t for _, t in runs)
    # Match quote-agnostically; normalization preserves positions and lengths.
    nflat = _normalize_quotes(flat)
    nold = _normalize_quotes(old)

    def doc_index_at(flat_pos: int) -> int:
        pos = 0
        for start, text in runs:
            if pos + len(text) > flat_pos:
                return start + (flat_pos - pos)
            pos += len(text)
        raise IndexError(f"flat position {flat_pos} out of range")

    plan = _plan_edit_matches(flat, nflat, old, nold, all_occurrences, occurrence)
    if plan["kind"] == "ambiguous":
        return plan["result"]
    positions = plan["positions"]

    requests = []
    for flat_pos in sorted(positions, reverse=True):
        start = doc_index_at(flat_pos)
        end = start + _utf16_len(old)
        requests.append({"deleteContentRange": {"range": {"startIndex": start, "endIndex": end}}})
        if new:
            requests.append({"insertText": {"location": {"index": start}, "text": new}})

    service.documents().batchUpdate(documentId=doc_id, body={"requests": requests}).execute()
    return {"status": "edited", "occurrences_replaced": len(positions), "old": old, "new": new}


def find_text(doc_id_or_url: str, text: str) -> dict:
    """Locate `text` in a doc; report each match's indices, style, and code flag."""
    doc_id = _extract_doc_id(doc_id_or_url)
    doc = _docs_service().documents().get(documentId=doc_id).execute()
    matches = _find_matches(doc, text)
    return {"query": text, "match_count": len(matches), "matches": matches}


def insert_after(
    doc_id_or_url: str, anchor: str, text: str, allow_multiple: bool = False
) -> dict:
    """
    Insert `text` as new paragraph(s) after the paragraph containing `anchor`.

    The new paragraphs inherit the anchor paragraph's style. Use blank lines in
    `text` to create multiple paragraphs. Raises if the anchor is missing, or
    ambiguous unless allow_multiple is set (then the first match is used).
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()
    doc = service.documents().get(documentId=doc_id).execute()
    request, body_index = _insert_after_plan(
        doc, anchor, text, require_unique=not allow_multiple
    )
    service.documents().batchUpdate(
        documentId=doc_id, body={"requests": [request]}
    ).execute()
    return {
        "status": "inserted",
        "after_body_index": body_index,
        "anchor": anchor,
        "text": text,
    }


def link_text(
    doc_id_or_url: str, text: str, url: str, all_occurrences: bool = False
) -> dict:
    """
    Hyperlink occurrences of `text` to `url`, preserving other text styling.

    By default requires exactly one match; pass all_occurrences=True to link
    every occurrence. Raises if the text is not found.
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()
    doc = service.documents().get(documentId=doc_id).execute()
    requests, spans = _link_plan(doc, text, url, all_occurrences=all_occurrences)
    service.documents().batchUpdate(
        documentId=doc_id, body={"requests": requests}
    ).execute()
    return {
        "status": "linked",
        "text": text,
        "url": url,
        "occurrences": len(spans),
        "spans": spans,
    }


def style_text(
    doc_id_or_url: str,
    text: str,
    italic: bool = False,
    bold: bool = False,
    underline: bool = False,
    color: str | None = None,
    all_occurrences: bool = False,
) -> dict:
    """
    Apply character styling (italic/bold/underline/color) to occurrences of `text`,
    preserving other styling. Requires exactly one match unless all_occurrences is set.
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    service = _docs_service()
    doc = service.documents().get(documentId=doc_id).execute()
    requests, spans = _style_plan(
        doc, text, italic=italic, bold=bold, underline=underline,
        color=color, all_occurrences=all_occurrences,
    )
    service.documents().batchUpdate(
        documentId=doc_id, body={"requests": requests}
    ).execute()
    return {
        "status": "styled",
        "text": text,
        "occurrences": len(spans),
        "spans": spans,
    }


def build_epub(
    doc_ids_or_urls: list[str],
    output: str | None = None,
    title: str | None = None,
    subtitle: str | None = None,
    author: str | None = None,
    cover: str | None = None,
    max_image_width: int = 1600,
    optimize_images: bool = True,
) -> dict:
    """Build an EPUB from multiple Google Docs, excluding PV review sections.

    title/subtitle/author are book metadata (book-specific — supply from the
    work's context, not hardcoded). cover is a local path or URL to a cover image.
    When optimize_images is set, images wider than max_image_width are downscaled
    and re-encoded to keep the EPUB small.
    """
    docs_service = _docs_service()
    chapters = []
    media_items = []  # OPF manifest entries for images
    media_files = []  # (epub href, bytes) to write into the zip
    skipped_images = 0
    for index, doc_ref in enumerate(doc_ids_or_urls, start=1):
        doc_id = _extract_doc_id(doc_ref)
        doc = docs_service.documents().get(documentId=doc_id).execute()
        chapter_title = doc.get("title", f"Chapter {index}")
        blocks = _extract_blocks(doc)

        image_paths = {}
        for block in blocks:
            if block["type"] != "image":
                continue
            object_id = block["object_id"]
            if object_id in image_paths:
                continue
            uri = _image_content_uri(doc, object_id)
            if not uri:
                skipped_images += 1
                continue
            try:
                data, media_type = _download_image(uri)
            except Exception:
                # One bad image shouldn't abort a whole draft EPUB.
                skipped_images += 1
                continue
            if optimize_images:
                data, media_type = _downscale_image(data, media_type, max_image_width)
            seq = len(image_paths) + 1
            href = f"images/ch{index:02d}-img{seq:02d}.{_media_extension(media_type)}"
            image_paths[object_id] = href
            media_items.append({
                "id": f"img-{index:02d}-{seq:02d}",
                "href": href,
                "media_type": media_type,
            })
            media_files.append((href, data))

        chapters.append({
            "title": chapter_title,
            "filename": _chapter_filename(index),
            "xhtml": _blocks_to_xhtml(chapter_title, blocks, image_paths),
            "image_count": len(image_paths),
        })

    book_title = title or _default_epub_title([chapter["title"] for chapter in chapters])
    output_path = Path(output) if output else _default_epub_output_path(book_title)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    book_id = str(uuid.uuid4())

    # Front matter: an optional cover-image page, then a generated title page.
    front_matter = []  # [{id, href}] in spine order
    front_files = []   # [(href, xhtml)]
    cover_image_id = None
    if cover:
        cover_bytes, cover_type = _read_cover_image(cover)
        if optimize_images:
            cover_bytes, cover_type = _downscale_image(cover_bytes, cover_type, max_image_width)
        cover_href = f"images/cover.{_media_extension(cover_type)}"
        cover_image_id = "cover-image"
        media_items.insert(0, {"id": cover_image_id, "href": cover_href, "media_type": cover_type})
        media_files.append((cover_href, cover_bytes))
        front_matter.append({"id": "cover-page", "href": "cover.xhtml"})
        front_files.append(("cover.xhtml", _cover_page_xhtml(cover_href)))
    front_matter.append({"id": "titlepage", "href": "title.xhtml"})
    front_files.append(("title.xhtml", _title_page_xhtml(book_title, subtitle, author)))

    nav = _epub_nav(book_title, chapters)
    package = _epub_package(
        book_title, book_id, chapters, media_items,
        author=author, front_matter=front_matter, cover_image_id=cover_image_id,
    )
    stylesheet = (
        "body { font-family: serif; line-height: 1.4; }\n"
        "h1, h2, h3, h4 { font-family: sans-serif; }\n"
        "section { max-width: 42em; margin: 0 auto; }\n"
        "p, li { margin: 0.6em 0; }\n"
        "figure { margin: 1em 0; text-align: center; }\n"
        "img { max-width: 100%; height: auto; }\n"
        ".titlepage { text-align: center; margin-top: 20%; }\n"
        "h1.title { font-size: 2.2em; }\n"
        "p.subtitle { font-size: 1.2em; font-style: italic; }\n"
        "p.author { margin-top: 2em; font-size: 1.1em; }\n"
        "section[epub|type=\"cover\"] { max-width: none; }\n"
        "img.cover { width: 100%; height: 100%; object-fit: contain; }\n"
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
        for href, xhtml in front_files:
            epub.writestr(f"OEBPS/{href}", xhtml)
        for chapter in chapters:
            epub.writestr(f"OEBPS/{chapter['filename']}", chapter["xhtml"])
        for href, data in media_files:
            epub.writestr(f"OEBPS/{href}", data)

    return {
        "status": "built",
        "title": book_title,
        "subtitle": subtitle,
        "author": author,
        "has_cover": bool(cover),
        "output": str(output_path),
        "images_embedded": len(media_files) - (1 if cover else 0),
        "images_skipped": skipped_images,
        "chapters": [
            {
                "title": chapter["title"],
                "filename": chapter["filename"],
                "images": chapter["image_count"],
            }
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
        cell_writes = []
        for row_index, row in enumerate(rows):
            for col_index, cell_text in enumerate(row):
                cell = table["tableRows"][row_index]["tableCells"][col_index]
                if not cell_text:
                    continue
                cell_insert_at = cell["content"][0]["endIndex"] - 1
                cell_writes.append((cell_insert_at, cell_text, row_index == 0))

        requests = []
        for cell_insert_at, cell_text, is_header in sorted(
            cell_writes, key=lambda item: item[0], reverse=True
        ):
            requests.append({
                "insertText": {
                    "location": {"index": cell_insert_at},
                    "text": cell_text,
                }
            })
            if is_header:
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


def make_review_copy(
    doc_id_or_url: str,
    folder_id_or_url: str,
    suffix_template: str | None = None,
    clear_pv_review_section: bool = True,
) -> dict:
    """
    Copy a Google Doc into a folder, suffix the title with a dated template
    (default ' - DRAFT {date}'), and optionally clear the PV review section.

    suffix_template supports '{date}' for today's ISO date.
    """
    doc_id = _extract_doc_id(doc_id_or_url)
    folder_id = _extract_folder_id(folder_id_or_url)

    docs_service = _docs_service()
    drive_service = _drive_service()

    source_doc = docs_service.documents().get(documentId=doc_id).execute()
    source_title = source_doc.get("title", "Untitled")
    copy_title = _review_copy_title(source_title, suffix_template=suffix_template)

    copied = (
        drive_service.files()
        .copy(
            fileId=doc_id,
            body={"name": copy_title, "parents": [folder_id]},
            fields="id,name,parents,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )

    if clear_pv_review_section:
        clear_review_section(copied["id"])

    return {
        "status": "created",
        "source_id": doc_id,
        "source_title": source_title,
        "copy_id": copied["id"],
        "copy_title": copied["name"],
        "copy_url": copied.get("webViewLink", f"https://docs.google.com/document/d/{copied['id']}"),
        "folder_id": folder_id,
    }


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

    p_fetch = sub.add_parser("fetch", help="Fetch the title and text of a Google Doc.")
    p_fetch.add_argument("doc", metavar="DOC_URL")
    p_fetch.add_argument(
        "--include-resolved",
        action="store_true",
        help="Include resolved comments in the output (default: skip them).",
    )

    p_slides = sub.add_parser(
        "slides-fetch",
        help="Fetch slide text from a Google Slides presentation.",
    )
    p_slides.add_argument("presentation", metavar="PRESENTATION_URL")

    p_slide_thumb = sub.add_parser(
        "slides-thumb",
        help="Fetch a Google Slides page thumbnail URL.",
    )
    p_slide_thumb.add_argument("presentation", metavar="PRESENTATION_URL")
    p_slide_thumb.add_argument("page_object_id", metavar="PAGE_OBJECT_ID")
    p_slide_thumb.add_argument(
        "--size",
        default="LARGE",
        choices=["SMALL", "MEDIUM", "LARGE"],
        help="Thumbnail size. Default: LARGE.",
    )

    p_sheet_fetch = sub.add_parser(
        "sheet-fetch",
        help="Fetch Google Sheet metadata and optional range values.",
    )
    p_sheet_fetch.add_argument("sheet", metavar="SHEET_URL")
    p_sheet_fetch.add_argument(
        "--range",
        dest="range_name",
        default=None,
        help="A1 notation range to fetch, e.g. 'Sheet1!A1:G20'.",
    )

    p_sheet_update = sub.add_parser(
        "sheet-update",
        help="Update a Google Sheet range from a JSON matrix.",
    )
    p_sheet_update.add_argument("sheet", metavar="SHEET_URL")
    p_sheet_update.add_argument("range_name", metavar="RANGE")
    p_sheet_update.add_argument(
        "values_json",
        metavar="VALUES_JSON",
        help='JSON 2D array, e.g. \'[["A1", "B1"], ["A2", "B2"]]\'',
    )

    p_figure_map = sub.add_parser(
        "figure-map",
        help="List inline-image neighborhoods in a Google Doc by body index.",
    )
    p_figure_map.add_argument("doc", metavar="DOC_URL")

    p_replace_block = sub.add_parser(
        "replace-block",
        help="Replace a contiguous Google Doc body-element range with text.",
    )
    p_replace_block.add_argument("doc", metavar="DOC_URL")
    p_replace_block.add_argument("start_body_index", type=int)
    p_replace_block.add_argument("end_body_index", type=int)
    p_replace_block.add_argument("text", metavar="TEXT")

    p_insert_image = sub.add_parser(
        "insert-image",
        help="Insert an inline image at a Google Doc body-element index.",
    )
    p_insert_image.add_argument("doc", metavar="DOC_URL")
    p_insert_image.add_argument("body_index", type=int)
    p_insert_image.add_argument("image_url", metavar="IMAGE_URL")
    p_insert_image.add_argument("--width-pt", type=float, default=468.0)
    p_insert_image.add_argument("--height-pt", type=float, default=263.25)

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
    p_epub.add_argument(
        "--subtitle",
        help="Book subtitle, shown on the generated title page.",
    )
    p_epub.add_argument(
        "--author",
        help="Author, written to dc:creator and the title page.",
    )
    p_epub.add_argument(
        "--cover",
        help="Cover image (local path or URL) embedded as the EPUB cover.",
    )
    p_epub.add_argument(
        "--max-image-width",
        type=int,
        default=1600,
        help="Downscale images wider than this many pixels (default 1600).",
    )
    p_epub.add_argument(
        "--no-optimize",
        action="store_true",
        help="Don't downscale or re-encode images (embed at source resolution).",
    )

    p_review_copy = sub.add_parser(
        "review-copy",
        help="Copy a Google Doc into a folder with a dated title and clear its PV review section.",
    )
    p_review_copy.add_argument("doc", metavar="DOC_URL")
    p_review_copy.add_argument("folder", metavar="FOLDER_URL")
    p_review_copy.add_argument(
        "--suffix",
        default=None,
        help="Title suffix template, supports {date}. Default: ' - DRAFT {date}'.",
    )
    p_review_copy.add_argument(
        "--keep-review-section",
        action="store_true",
        help="Don't clear the PV review section in the new copy.",
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

    p_comments = sub.add_parser(
        "comments",
        help="List a doc's comments (id, author, content, quoted text, resolved).",
    )
    p_comments.add_argument("doc", metavar="DOC_URL")
    p_comments.add_argument(
        "--include-resolved",
        action="store_true",
        help="Include resolved comments. Default: unresolved only.",
    )

    p_resolve = sub.add_parser("resolve", help="Resolve a comment on a Google Doc.")
    p_resolve.add_argument("doc", metavar="DOC_URL")
    p_resolve.add_argument("comment_id", metavar="COMMENT_ID",
                           help="Comment ID (use `pv comments` to list ids).")

    p_resolve_all = sub.add_parser("resolve-all", help="Resolve every unresolved comment on a doc.")
    p_resolve_all.add_argument("doc", metavar="DOC_URL")

    p_comment = sub.add_parser(
        "comment",
        help="Create an anchored sidebar comment with quoted-text context.",
    )
    p_comment.add_argument("doc", metavar="DOC_URL")
    p_comment.add_argument("quoted_text", metavar="QUOTED_TEXT",
                           help="Text the comment is anchored to (shown in the sidebar).")
    p_comment.add_argument("comment", metavar="COMMENT", help="Comment body text.")

    p_edit = sub.add_parser("edit", help="Replace text in a Google Doc body.")
    p_edit.add_argument("doc", metavar="DOC_URL")
    p_edit.add_argument("old", metavar="OLD", help="Text to replace.")
    p_edit.add_argument("new", metavar="NEW", help="Replacement text.")
    p_edit.add_argument(
        "--all",
        action="store_true",
        dest="all_occurrences",
        help="Replace every occurrence. Default: require exactly one match.",
    )
    p_edit.add_argument(
        "--occurrence",
        type=int,
        default=None,
        help="Pick the Nth match (1-based) when OLD matches several places.",
    )

    p_find = sub.add_parser(
        "find",
        help="Find text in a doc; report each match's location, style, and code flag.",
    )
    p_find.add_argument("doc", metavar="DOC_URL")
    p_find.add_argument("text", metavar="TEXT", help="Exact substring to locate.")

    p_insert_after = sub.add_parser(
        "insert-after",
        help="Insert text as new paragraph(s) after the paragraph containing an anchor.",
    )
    p_insert_after.add_argument("doc", metavar="DOC_URL")
    p_insert_after.add_argument(
        "anchor", metavar="ANCHOR", help="Unique substring of the target paragraph."
    )
    p_insert_after.add_argument(
        "text", metavar="TEXT", help="Text to insert; use blank lines for multiple paragraphs."
    )
    p_insert_after.add_argument(
        "--allow-multiple",
        action="store_true",
        help="Insert after the first match even if the anchor is not unique.",
    )

    p_link = sub.add_parser("link", help="Hyperlink a span of text in a doc.")
    p_link.add_argument("doc", metavar="DOC_URL")
    p_link.add_argument("text", metavar="TEXT", help="Exact span to hyperlink.")
    p_link.add_argument("url", metavar="URL")
    p_link.add_argument(
        "--all",
        action="store_true",
        dest="all_occurrences",
        help="Link every occurrence. Default: require exactly one match.",
    )

    p_style = sub.add_parser(
        "style",
        help="Apply italic/bold/underline/color to a span of text.",
    )
    p_style.add_argument("doc", metavar="DOC_URL")
    p_style.add_argument("text", metavar="TEXT", help="Exact span to style.")
    p_style.add_argument("--italic", action="store_true")
    p_style.add_argument("--bold", action="store_true")
    p_style.add_argument("--underline", action="store_true")
    p_style.add_argument("--color", metavar="HEX", help="Foreground color, e.g. #d3002d.")
    p_style.add_argument(
        "--all",
        action="store_true",
        dest="all_occurrences",
        help="Style every occurrence. Default: require exactly one match.",
    )

    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    if args.command == "list":
        result = list_folder(args.folder)
    elif args.command == "fetch":
        result = fetch_document(args.doc, include_resolved=args.include_resolved)
    elif args.command == "slides-fetch":
        result = fetch_presentation(args.presentation)
    elif args.command == "slides-thumb":
        result = presentation_thumbnail(
            args.presentation,
            args.page_object_id,
            size=args.size,
        )
    elif args.command == "sheet-fetch":
        result = fetch_sheet(args.sheet, range_name=args.range_name)
    elif args.command == "sheet-update":
        result = update_sheet(args.sheet, args.range_name, json.loads(args.values_json))
    elif args.command == "figure-map":
        result = figure_map(args.doc)
    elif args.command == "replace-block":
        result = replace_body_range(
            args.doc,
            args.start_body_index,
            args.end_body_index,
            args.text,
        )
    elif args.command == "insert-image":
        result = insert_image_at_body_index(
            args.doc,
            args.body_index,
            args.image_url,
            width_pt=args.width_pt,
            height_pt=args.height_pt,
        )
    elif args.command == "clear":
        result = clear_review_section(args.doc)
    elif args.command == "append":
        result = append_content(args.doc, args.heading, args.text)
    elif args.command == "build-epub":
        result = build_epub(
            args.docs,
            output=args.output,
            title=args.title,
            subtitle=args.subtitle,
            author=args.author,
            cover=args.cover,
            max_image_width=args.max_image_width,
            optimize_images=not args.no_optimize,
        )
    elif args.command == "review-copy":
        result = make_review_copy(
            args.doc,
            args.folder,
            suffix_template=args.suffix,
            clear_pv_review_section=not args.keep_review_section,
        )
    elif args.command == "mv":
        result = move_document(args.doc, args.folder)
    elif args.command == "cp":
        result = copy_document(args.doc, args.folder, name=args.name)
    elif args.command == "comments":
        result = list_comments(args.doc, include_resolved=args.include_resolved)
    elif args.command == "resolve":
        result = resolve_comment(args.doc, args.comment_id)
    elif args.command == "resolve-all":
        result = resolve_all_comments(args.doc)
    elif args.command == "comment":
        result = comment_document(args.doc, args.quoted_text, args.comment)
    elif args.command == "edit":
        result = edit_document(
            args.doc, args.old, args.new,
            all_occurrences=args.all_occurrences, occurrence=args.occurrence,
        )
    elif args.command == "find":
        result = find_text(args.doc, args.text)
    elif args.command == "insert-after":
        result = insert_after(args.doc, args.anchor, args.text, allow_multiple=args.allow_multiple)
    elif args.command == "link":
        result = link_text(args.doc, args.text, args.url, all_occurrences=args.all_occurrences)
    elif args.command == "style":
        result = style_text(
            args.doc, args.text,
            italic=args.italic, bold=args.bold, underline=args.underline,
            color=args.color, all_occurrences=args.all_occurrences,
        )
    else:
        result = append_review_note(args.doc, args.quoted_text, args.comment)

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
