"""Tests for pv.py — pure functions."""

from datetime import datetime

from pv import (
    _blocks_to_xhtml,
    _default_epub_output_path,
    _extract_blocks,
    _extract_doc_id,
    _extract_folder_id,
    _extract_text,
    _paragraph_location,
    _parse_append_blocks,
    _review_copy_title,
    _slugify,
)

# ---------------------------------------------------------------------------
# _extract_doc_id
# ---------------------------------------------------------------------------

def test_extract_doc_id_from_url():
    url = "https://docs.google.com/document/d/abc123XYZ/edit"
    assert _extract_doc_id(url) == "abc123XYZ"

def test_extract_doc_id_bare():
    assert _extract_doc_id("abc123XYZ") == "abc123XYZ"

def test_extract_doc_id_strips_whitespace():
    assert _extract_doc_id("  abc123  ") == "abc123"


# ---------------------------------------------------------------------------
# _extract_folder_id
# ---------------------------------------------------------------------------

def test_extract_folder_id_from_url():
    url = "https://drive.google.com/drive/folders/folderXYZ"
    assert _extract_folder_id(url) == "folderXYZ"

def test_extract_folder_id_bare():
    assert _extract_folder_id("folderXYZ") == "folderXYZ"


# ---------------------------------------------------------------------------
# _extract_text
# ---------------------------------------------------------------------------

SAMPLE_DOC = {
    "body": {
        "content": [
            {"sectionBreak": {}},
            {
                "paragraph": {
                    "elements": [
                        {"textRun": {"content": "Hello, "}},
                        {"textRun": {"content": "world.\n"}},
                    ]
                }
            },
            {
                "paragraph": {
                    "elements": [
                        {"textRun": {"content": "Second paragraph.\n"}},
                    ]
                }
            },
        ]
    }
}

def test_extract_text_concatenates_runs():
    assert _extract_text(SAMPLE_DOC) == "Hello, world.\nSecond paragraph.\n"

def test_extract_text_skips_non_paragraph_elements():
    doc = {"body": {"content": [{"sectionBreak": {}}]}}
    assert _extract_text(doc) == ""

def test_extract_text_empty_doc():
    assert _extract_text({}) == ""

def test_extract_text_stops_at_review_heading():
    doc = {"body": {"content": [
        {"paragraph": {"elements": [{"textRun": {"content": "Chapter text.\n"}}]}},
        {"paragraph": {"elements": [{"textRun": {"content": "🪶 Ploma Vermella Review\n"}}]}},
        {"paragraph": {"elements": [{"textRun": {"content": "Review note.\n"}}]}},
    ]}}
    assert _extract_text(doc) == "Chapter text.\n"


# ---------------------------------------------------------------------------
# _parse_append_blocks
# ---------------------------------------------------------------------------

def test_parse_append_blocks_keeps_paragraphs_and_bullets():
    blocks = _parse_append_blocks("Intro line\n- Bullet one\n\nNext para")
    assert blocks == [
        {"type": "paragraph", "text": "Intro line", "space_above": False},
        {"type": "bullet", "text": "Bullet one", "space_above": False},
        {"type": "paragraph", "text": "Next para", "space_above": True},
    ]


def test_parse_append_blocks_parses_markdown_tables():
    blocks = _parse_append_blocks(
        "| A | B |\n"
        "| --- | --- |\n"
        "| 1 | 2 |\n"
        "| 3 | 4 |\n"
    )
    assert blocks == [{
        "type": "table",
        "rows": [["A", "B"], ["1", "2"], ["3", "4"]],
        "space_above": False,
    }]


def test_parse_append_blocks_applies_spacing_before_table():
    blocks = _parse_append_blocks(
        "Para\n\n"
        "| A | B |\n"
        "| --- | --- |\n"
        "| 1 | 2 |\n"
    )
    assert blocks[-1] == {
        "type": "table",
        "rows": [["A", "B"], ["1", "2"]],
        "space_above": True,
    }


# ---------------------------------------------------------------------------
# _extract_blocks / EPUB helpers
# ---------------------------------------------------------------------------

BLOCK_DOC = {
    "body": {
        "content": [
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "TITLE"},
                    "elements": [{"textRun": {"content": "Chapter Title\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "HEADING_1"},
                    "elements": [{"textRun": {"content": "Section\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                    "elements": [{"textRun": {"content": "Body paragraph.\n"}}],
                }
            },
            {
                "paragraph": {
                    "bullet": {},
                    "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                    "elements": [{"textRun": {"content": "Bullet item\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "TITLE"},
                    "elements": [{"textRun": {"content": "🪶 Ploma Vermella Review\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                    "elements": [{"textRun": {"content": "Should not appear\n"}}],
                }
            },
        ]
    }
}


def test_extract_blocks_preserves_structure_and_stops_at_review():
    assert _extract_blocks(BLOCK_DOC) == [
        {"type": "heading", "level": 1, "text": "Chapter Title"},
        {"type": "heading", "level": 2, "text": "Section"},
        {"type": "paragraph", "text": "Body paragraph."},
        {"type": "list_item", "text": "Bullet item"},
    ]


def test_slugify_builds_safe_filename():
    assert _slugify("Chapter 07: Example Chapter") == "chapter-07-example-chapter"


def test_default_epub_output_path_includes_date_suffix():
    path = _default_epub_output_path(
        "Example Book",
        stamp=datetime(2026, 3, 24, 10, 30),
    )
    assert str(path) == "dist/example-book-20260324.epub"


def test_review_copy_title_appends_iso_date_suffix():
    title = _review_copy_title(
        "Chapter 10: Buy vs Build vs Grow",
        stamp=datetime(2026, 4, 6, 9, 0),
    )
    assert title == "Chapter 10: Buy vs Build vs Grow - DRAFT 2026-04-06"


def test_review_copy_title_supports_custom_template():
    title = _review_copy_title(
        "Introduction",
        stamp=datetime(2026, 4, 6, 9, 0),
        suffix_template=" ({date} review)",
    )
    assert title == "Introduction (2026-04-06 review)"


def test_blocks_to_xhtml_renders_list_and_headings():
    xhtml = _blocks_to_xhtml(
        "Example",
        [
            {"type": "heading", "level": 2, "text": "Section"},
            {"type": "paragraph", "text": "Body & more"},
            {"type": "list_item", "text": "One"},
            {"type": "list_item", "text": "Two"},
        ],
    )
    assert "<h2>Section</h2>" in xhtml
    assert "<p>Body &amp; more</p>" in xhtml
    assert "<ul>" in xhtml
    assert "<li>One</li>" in xhtml
    assert "<li>Two</li>" in xhtml


# ---------------------------------------------------------------------------
# _paragraph_location
# ---------------------------------------------------------------------------

STRUCTURED_DOC = {
    "body": {
        "content": [
            {"sectionBreak": {}},
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "TITLE"},
                    "elements": [{"textRun": {"content": "My Book\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "HEADING_1"},
                    "elements": [{"textRun": {"content": "Introduction\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                    "elements": [{"textRun": {"content": "First body paragraph.\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                    "elements": [{"textRun": {"content": "Second body paragraph.\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "HEADING_1"},
                    "elements": [{"textRun": {"content": "Methods\n"}}],
                }
            },
            {
                "paragraph": {
                    "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                    "elements": [{"textRun": {"content": "First methods paragraph.\n"}}],
                }
            },
        ]
    }
}


def test_paragraph_location_first_para():
    assert _paragraph_location(STRUCTURED_DOC, "First body paragraph.") == "Introduction: p1"

def test_paragraph_location_second_para():
    assert _paragraph_location(STRUCTURED_DOC, "Second body paragraph.") == "Introduction: p2"

def test_paragraph_location_new_section():
    assert _paragraph_location(STRUCTURED_DOC, "First methods paragraph.") == "Methods: p1"

def test_paragraph_location_not_found():
    assert _paragraph_location(STRUCTURED_DOC, "Nonexistent text") == ""

def test_paragraph_location_no_heading():
    doc = {"body": {"content": [{"paragraph": {
        "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
        "elements": [{"textRun": {"content": "Orphan paragraph.\n"}}],
    }}]}}
    assert _paragraph_location(doc, "Orphan paragraph.") == "p1"

def test_paragraph_location_stops_at_review_heading():
    doc = {"body": {"content": [
        {"paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "elements": [{"textRun": {"content": "Body text.\n"}}],
        }},
        {"paragraph": {
            "paragraphStyle": {"namedStyleType": "TITLE"},
            "elements": [{"textRun": {"content": "🪶 Ploma Vermella Review\n"}}],
        }},
        {"paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "elements": [{"textRun": {"content": "Note inside review.\n"}}],
        }},
    ]}}
    assert _paragraph_location(doc, "Note inside review.") == ""
