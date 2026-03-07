"""Tests for gdocs.py — pure functions and mocked API calls."""

from unittest.mock import MagicMock, patch

from gdocs import _extract_doc_id, _extract_folder_id, _extract_text, post_comment

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

def test_post_comment_posts_new():
    mock_drive = MagicMock()
    mock_drive.comments().create().execute.return_value = {
        "id": "new1",
        "content": "🪶 My comment",
        "author": {"displayName": "Chris Ford"},
    }
    with patch("gdocs.fetch_comments", return_value=[]), \
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
         patch("gdocs._drive_service", return_value=mock_drive):
        post_comment("doc123", "Some text", "My comment")

    call_body = mock_drive.comments().create.call_args[1]["body"]
    assert call_body["content"].startswith("🪶 ")
