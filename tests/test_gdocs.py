"""Tests for gdocs.py — pure functions and mocked API calls."""

from unittest.mock import MagicMock, patch

from gdocs import (
    _extract_doc_id,
    _extract_folder_id,
    _extract_text,
    _find_section_for_text,
    _is_fresh_tab,
    _paragraph_location,
    _section_end_index,
    _tab_has_heading,
    post_comment,
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


# ---------------------------------------------------------------------------
# post_comment — dedup
# ---------------------------------------------------------------------------

def test_post_comment_skips_duplicate():
    existing = [{"id": "c1", "quoted_text": "Some text", "author": "A", "content": "x"}]
    with patch("gdocs.fetch_comments", return_value=existing):
        result = post_comment("doc123", "Some text", "My comment")
    assert result["status"] == "skipped"
    assert result["existing_comment_id"] == "c1"

def test_post_comment_skips_duplicate_with_whitespace():
    existing = [{"id": "c1", "quoted_text": "  Some text  ", "author": "A", "content": "x"}]
    with patch("gdocs.fetch_comments", return_value=existing):
        result = post_comment("doc123", "Some text", "My comment")
    assert result["status"] == "skipped"


# ---------------------------------------------------------------------------
# post_comment — posts when no duplicate
# ---------------------------------------------------------------------------

def _mock_docs(doc=None):
    mock = MagicMock()
    mock.documents().get().execute.return_value = doc or {"body": {"content": []}}
    return mock


def test_post_comment_posts_new():
    mock_drive = MagicMock()
    mock_drive.comments().create().execute.return_value = {
        "id": "new1",
        "content": "🪶 My comment",
        "author": {"displayName": "Chris Ford"},
    }
    with patch("gdocs.fetch_comments", return_value=[]), \
         patch("gdocs._docs_service", return_value=_mock_docs()), \
         patch("gdocs._drive_service", return_value=mock_drive):
        result = post_comment("doc123", "Some text", "My comment")

    assert result["status"] == "posted"
    assert result["id"] == "new1"
    assert result["author"] == "Chris Ford"

def test_post_comment_prefixes_quill_emoji():
    mock_drive = MagicMock()
    mock_drive.comments().create().execute.return_value = {
        "id": "new1", "content": "🪶 My comment", "author": {"displayName": "Chris"},
    }
    with patch("gdocs.fetch_comments", return_value=[]), \
         patch("gdocs._docs_service", return_value=_mock_docs()), \
         patch("gdocs._drive_service", return_value=mock_drive):
        post_comment("doc123", "Some text", "My comment")

    call_body = mock_drive.comments().create.call_args[1]["body"]
    assert call_body["content"].startswith("🪶 ")

# ---------------------------------------------------------------------------
# Tab helpers
# ---------------------------------------------------------------------------

_TAB_CONTENT_FRESH = [
    {"sectionBreak": {}, "startIndex": 0, "endIndex": 1},
    {
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "elements": [{"textRun": {"content": "\n"}}],
        },
        "startIndex": 1,
        "endIndex": 2,
    },
]

_TAB_CONTENT_WITH_HEADING = [
    {"sectionBreak": {}, "startIndex": 0, "endIndex": 1},
    {
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "TITLE"},
            "elements": [{"textRun": {"content": "Ploma Vermella\n"}}],
        },
        "startIndex": 1,
        "endIndex": 16,
    },
    {
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "HEADING_1"},
            "elements": [{"textRun": {"content": "Introduction\n"}}],
        },
        "startIndex": 16,
        "endIndex": 29,
    },
    {
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "elements": [{"textRun": {"content": "A note.\n"}}],
        },
        "startIndex": 29,
        "endIndex": 37,
    },
    {
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "HEADING_1"},
            "elements": [{"textRun": {"content": "Methods\n"}}],
        },
        "startIndex": 37,
        "endIndex": 45,
    },
    {
        "paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "elements": [{"textRun": {"content": "\n"}}],
        },
        "startIndex": 45,
        "endIndex": 46,
    },
]


def test_is_fresh_tab_true():
    assert _is_fresh_tab(_TAB_CONTENT_FRESH) is True


def test_is_fresh_tab_false():
    assert _is_fresh_tab(_TAB_CONTENT_WITH_HEADING) is False


def test_tab_has_heading_found():
    assert _tab_has_heading(_TAB_CONTENT_WITH_HEADING, "Introduction") is True


def test_tab_has_heading_not_found():
    assert _tab_has_heading(_TAB_CONTENT_WITH_HEADING, "Conclusion") is False


def test_section_end_index_before_next_heading():
    # "Introduction" section ends at index 37 (start of "Methods" heading)
    assert _section_end_index(_TAB_CONTENT_WITH_HEADING, "Introduction") == 37


def test_section_end_index_last_section():
    # "Methods" is the last heading — falls back to end of content
    result = _section_end_index(_TAB_CONTENT_WITH_HEADING, "Methods")
    assert result == 45  # endIndex of last element minus 1


def test_find_section_for_text_found():
    content = STRUCTURED_DOC["body"]["content"]
    assert _find_section_for_text(content, "First body paragraph.") == "Introduction"


def test_find_section_for_text_different_section():
    content = STRUCTURED_DOC["body"]["content"]
    assert _find_section_for_text(content, "First methods paragraph.") == "Methods"


def test_find_section_for_text_not_found():
    content = STRUCTURED_DOC["body"]["content"]
    assert _find_section_for_text(content, "Nonexistent text") is None


# ---------------------------------------------------------------------------


def test_post_comment_includes_location():
    mock_drive = MagicMock()
    mock_drive.comments().create().execute.return_value = {
        "id": "new1", "content": "🪶 Section 1: p1: note", "author": {"displayName": "Chris"},
    }
    with patch("gdocs.fetch_comments", return_value=[]), \
         patch("gdocs._docs_service", return_value=_mock_docs(STRUCTURED_DOC)), \
         patch("gdocs._drive_service", return_value=mock_drive):
        post_comment("doc123", "First body paragraph.", "note")

    call_body = mock_drive.comments().create.call_args[1]["body"]
    assert "Introduction: p1" in call_body["content"]
