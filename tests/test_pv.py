"""Tests for pv.py — pure functions."""

from datetime import datetime

import pytest

from pv import (
    _blocks_to_xhtml,
    _body_element_at,
    _default_epub_output_path,
    _epub_package,
    _extract_blocks,
    _extract_doc_id,
    _extract_folder_id,
    _extract_presentation_id,
    _extract_spreadsheet_id,
    _extract_text,
    _figure_map_from_doc,
    _find_matches,
    _image_content_uri,
    _inline_html,
    _insert_after_plan,
    _is_code_paragraph,
    _is_image_paragraph,
    _link_plan,
    _media_extension,
    _paragraph_location,
    _paragraph_text,
    _parse_append_blocks,
    _review_copy_title,
    _shape_text,
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
# _extract_presentation_id / _extract_spreadsheet_id
# ---------------------------------------------------------------------------

def test_extract_presentation_id_from_url():
    url = "https://docs.google.com/presentation/d/pres123/edit"
    assert _extract_presentation_id(url) == "pres123"

def test_extract_presentation_id_bare():
    assert _extract_presentation_id("pres123") == "pres123"

def test_extract_spreadsheet_id_from_url():
    url = "https://docs.google.com/spreadsheets/d/sheet123/edit#gid=0"
    assert _extract_spreadsheet_id(url) == "sheet123"

def test_extract_spreadsheet_id_bare():
    assert _extract_spreadsheet_id("sheet123") == "sheet123"


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


def test_shape_text_concatenates_slide_runs():
    element = {
        "shape": {
            "text": {
                "textElements": [
                    {"textRun": {"content": "Hello"}},
                    {"textRun": {"content": " world"}},
                    {"textRun": {"content": "\n"}},
                ]
            }
        }
    }
    assert _shape_text(element) == "Hello world"


def test_paragraph_text_concatenates_runs():
    element = {
        "paragraph": {
            "elements": [
                {"textRun": {"content": "Hello, "}},
                {"textRun": {"content": "world.\n"}},
            ]
        }
    }
    assert _paragraph_text(element) == "Hello, world.\n"


def test_is_image_paragraph_detects_inline_object():
    element = {
        "paragraph": {
            "elements": [
                {"inlineObjectElement": {"inlineObjectId": "kix.123"}},
            ]
        }
    }
    assert _is_image_paragraph(element) is True


def test_figure_map_from_doc_reports_neighbor_text():
    doc = {
        "body": {
            "content": [
                {"paragraph": {"elements": [{"textRun": {"content": "Before figure.\n"}}]}},
                {"paragraph": {"elements": [{"inlineObjectElement": {"inlineObjectId": "kix.1"}}]},
                 "startIndex": 14, "endIndex": 16},
                {"paragraph": {"elements": [{"textRun": {"content": "Figure 1-1. Caption.\n"}}]}},
                {"paragraph": {"elements": [{"textRun": {"content": "After figure.\n"}}]}},
            ]
        }
    }
    assert _figure_map_from_doc(doc) == [{
        "body_index": 1,
        "start_index": 14,
        "end_index": 16,
        "prev_text": "Before figure.",
        "caption_text": "Figure 1-1. Caption.",
        "next_text": "After figure.",
    }]


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
        {"type": "heading", "level": 1, "text": "Chapter Title", "html": "Chapter Title"},
        {"type": "heading", "level": 2, "text": "Section", "html": "Section"},
        {"type": "paragraph", "text": "Body paragraph.", "html": "Body paragraph."},
        {"type": "list_item", "text": "Bullet item", "html": "Bullet item"},
    ]


def test_extract_blocks_emits_image_block():
    doc = {"body": {"content": [
        {"paragraph": {"paragraphStyle": {"namedStyleType": "NORMAL_TEXT"}, "elements": [
            {"inlineObjectElement": {"inlineObjectId": "kix.img1"}},
        ]}},
        {"paragraph": {"paragraphStyle": {"namedStyleType": "NORMAL_TEXT"}, "elements": [
            {"textRun": {"content": "Figure 1-1. A caption.\n"}},
        ]}},
    ]}}
    assert _extract_blocks(doc) == [
        {"type": "image", "object_id": "kix.img1"},
        {"type": "paragraph", "text": "Figure 1-1. A caption.", "html": "Figure 1-1. A caption."},
    ]


def test_inline_html_renders_italic_bold_and_link():
    elements = [
        {"textRun": {"content": "see ", "textStyle": {}}},
        {"textRun": {"content": "Lean Startup", "textStyle": {
            "italic": True, "link": {"url": "http://example.com/?a=1&b=2"}}}},
        {"textRun": {"content": " now\n", "textStyle": {"bold": True}}},
    ]
    out = _inline_html(elements)
    assert out == (
        'see <a href="http://example.com/?a=1&amp;b=2"><em>Lean Startup</em></a>'
        "<strong> now</strong>"
    )


def test_image_content_uri_resolves_from_inline_objects():
    doc = {"inlineObjects": {"kix.img1": {"inlineObjectProperties": {"embeddedObject": {
        "imageProperties": {"contentUri": "https://example.com/image.png"}}}}}}
    assert _image_content_uri(doc, "kix.img1") == "https://example.com/image.png"
    assert _image_content_uri(doc, "missing") is None


def test_media_extension_maps_types():
    assert _media_extension("image/png") == "png"
    assert _media_extension("image/jpeg; charset=binary") == "jpg"
    assert _media_extension("image/unknown") == "img"


def test_blocks_to_xhtml_renders_image_and_inline_html():
    xhtml = _blocks_to_xhtml(
        "Example",
        [
            {"type": "paragraph", "html": "see <em>Lean Startup</em>"},
            {"type": "image", "object_id": "kix.img1"},
            {"type": "image", "object_id": "kix.missing"},
        ],
        image_paths={"kix.img1": "images/ch01-img01.png"},
    )
    assert "<p>see <em>Lean Startup</em></p>" in xhtml
    assert '<figure><img src="images/ch01-img01.png" alt=""/></figure>' in xhtml
    # An image with no downloaded path is skipped, not rendered broken.
    assert "kix.missing" not in xhtml


def test_epub_package_includes_image_manifest_items():
    package = _epub_package(
        "Book", "uuid-1",
        [{"filename": "chapter-01.xhtml", "title": "Ch1"}],
        media_items=[{"id": "img-01-01", "href": "images/ch01-img01.png",
                      "media_type": "image/png"}],
    )
    assert '<item id="img-01-01" href="images/ch01-img01.png" media-type="image/png"/>' in package


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


# ---------------------------------------------------------------------------
# find / insert-after / link helpers (pure planning logic)
# ---------------------------------------------------------------------------

def _para(start, text, *, style="NORMAL_TEXT", font=None):
    """Build a fake body paragraph element with consistent indices."""
    text_style = {"weightedFontFamily": {"fontFamily": font}} if font else {}
    return {
        "startIndex": start,
        "endIndex": start + len(text),
        "paragraph": {
            "paragraphStyle": {"namedStyleType": style},
            "elements": [{
                "startIndex": start,
                "endIndex": start + len(text),
                "textRun": {"content": text, "textStyle": text_style},
            }],
        },
    }


def _fake_doc(*paras):
    return {"body": {"content": list(paras)}}


def test_find_matches_locates_span():
    doc = _fake_doc(_para(1, "The quick brown fox\n"), _para(21, "jumps over\n"))
    matches = _find_matches(doc, "brown")
    assert len(matches) == 1
    m = matches[0]
    assert m["start_index"] == 11
    assert m["end_index"] == 16
    assert m["body_index"] == 0
    assert m["context"] == "The quick brown fox"
    assert m["is_code"] is False

def test_find_matches_no_match_returns_empty():
    assert _find_matches(_fake_doc(_para(1, "hello\n")), "zzz") == []

def test_find_matches_flags_code_paragraph():
    doc = _fake_doc(_para(1, '  "key": "value"\n', font="Consolas"))
    assert _find_matches(doc, "key")[0]["is_code"] is True

def test_body_element_at_returns_containing_element():
    doc = _fake_doc(_para(1, "first\n"), _para(7, "second\n"))
    idx, el = _body_element_at(doc["body"]["content"], 8)
    assert idx == 1
    assert el["startIndex"] == 7

def test_is_code_paragraph_mixed_fonts_is_false():
    mono = {"weightedFontFamily": {"fontFamily": "Consolas"}}
    para = {"paragraph": {"paragraphStyle": {"namedStyleType": "NORMAL_TEXT"}, "elements": [
        {"textRun": {"content": "code ", "textStyle": mono}},
        {"textRun": {"content": "prose\n", "textStyle": {}}},
    ]}}
    assert _is_code_paragraph(para) is False

def test_insert_after_plan_builds_request():
    doc = _fake_doc(_para(1, "Intro line.\n"), _para(13, "Anchor paragraph here.\n"))
    request, body_index = _insert_after_plan(doc, "Anchor paragraph", "NEW PARAGRAPH")
    assert body_index == 1
    assert request["insertText"]["location"]["index"] == 35
    assert request["insertText"]["text"] == "\nNEW PARAGRAPH"

def test_insert_after_plan_missing_anchor_raises():
    with pytest.raises(ValueError):
        _insert_after_plan(_fake_doc(_para(1, "x\n")), "nope", "y")

def test_insert_after_plan_ambiguous_raises_then_allows():
    doc = _fake_doc(_para(1, "shared token\n"), _para(14, "shared token again\n"))
    with pytest.raises(ValueError):
        _insert_after_plan(doc, "shared token", "y")
    _request, body_index = _insert_after_plan(doc, "shared token", "y", require_unique=False)
    assert body_index == 0

def test_link_plan_builds_update_request():
    doc = _fake_doc(_para(1, "See the Lean Startup here.\n"))
    requests, spans = _link_plan(doc, "Lean Startup", "http://example.com")
    assert len(requests) == 1
    style = requests[0]["updateTextStyle"]
    assert style["range"] == {"startIndex": 9, "endIndex": 21}
    assert style["textStyle"]["link"]["url"] == "http://example.com"
    assert style["fields"] == "link"
    assert spans == [{"start_index": 9, "end_index": 21}]

def test_link_plan_missing_text_raises():
    with pytest.raises(ValueError):
        _link_plan(_fake_doc(_para(1, "hello\n")), "zzz", "http://e")

def test_link_plan_ambiguous_requires_all_occurrences():
    doc = _fake_doc(_para(1, "go go\n"))
    with pytest.raises(ValueError):
        _link_plan(doc, "go", "http://e")
    requests, _spans = _link_plan(doc, "go", "http://e", all_occurrences=True)
    assert len(requests) == 2
