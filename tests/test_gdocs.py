"""Tests for gdocs.py — pure functions."""

from gdocs import (
    _extract_doc_id,
    _extract_folder_id,
    _extract_text,
    _paragraph_location,
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
