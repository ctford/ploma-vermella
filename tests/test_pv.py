"""Tests for pv.py — pure functions."""

import inspect
import io
import os
from datetime import datetime

import pytest
from PIL import Image

from pv import (
    _block_html,
    _blocks_to_xhtml,
    _body_element_at,
    _build_parser,
    _bullets_plan,
    _chapter_filename,
    _cover_page_xhtml,
    _default_epub_output_path,
    _default_epub_title,
    _doc_index_at,
    _doc_text_runs,
    _downscale_image,
    _epub_nav,
    _epub_package,
    _extract_blocks,
    _extract_doc_id,
    _extract_folder_id,
    _extract_presentation_id,
    _extract_spreadsheet_id,
    _extract_text,
    _figure_map_from_doc,
    _find_matches,
    _heading_plan,
    _image_content_uri,
    _inline_html,
    _inline_object_ids,
    _insert_after_plan,
    _is_code_paragraph,
    _is_image_paragraph,
    _is_table_separator,
    _link_plan,
    _map_comments,
    _media_extension,
    _named_style_for_level,
    _normalize_quotes,
    _outline_from_doc,
    _paragraph_location,
    _paragraph_text,
    _parse_append_blocks,
    _parse_hex_color,
    _parse_table_row,
    _plan_edit_matches,
    _preceding_image_id,
    _replace_image_plan,
    _review_copy_title,
    _shape_text,
    _slugify,
    _style_plan,
    _text_from_elements,
    _title_page_xhtml,
    _utf16_len,
    main,
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

def _cell(text):
    return {"content": [{"paragraph": {"elements": [{"textRun": {"content": text + "\n"}}]}}]}

def test_extract_text_renders_table_rows():
    doc = {"body": {"content": [
        {"table": {"tableRows": [
            {"tableCells": [_cell("Scenario"), _cell("Rate")]},
            {"tableCells": [_cell("Default"), _cell("52%")]},
        ]}},
    ]}}
    assert _extract_text(doc) == "Scenario | Rate\nDefault | 52%\n"

def test_extract_text_interleaves_table_with_paragraphs():
    doc = {"body": {"content": [
        {"paragraph": {"elements": [{"textRun": {"content": "Before.\n"}}]}},
        {"table": {"tableRows": [{"tableCells": [_cell("A"), _cell("B")]}]}},
        {"paragraph": {"elements": [{"textRun": {"content": "After.\n"}}]}},
    ]}}
    assert _extract_text(doc) == "Before.\nA | B\nAfter.\n"


# ---------------------------------------------------------------------------
# _plan_edit_matches (the ambiguous result contract)
# ---------------------------------------------------------------------------

def _plan(flat, old, all_occurrences=False, occurrence=None):
    return _plan_edit_matches(flat, flat, old, old, all_occurrences, occurrence)

def test_plan_edit_single_match_is_ok():
    plan = _plan("the quick brown fox", "quick")
    assert plan["kind"] == "ok"
    assert plan["positions"] == [4]

def test_plan_edit_multiple_matches_is_ambiguous():
    plan = _plan("a cat and a cat", "cat")
    assert plan["kind"] == "ambiguous"
    result = plan["result"]
    assert result["status"] == "ambiguous"
    assert result["reason"] == "multiple_matches"
    assert [o["id"] for o in result["options"]] == [1, 2]
    assert result["resolution"]["field"] == "occurrence"

def test_plan_edit_all_occurrences_replaces_every_match():
    plan = _plan("a cat and a cat", "cat", all_occurrences=True)
    assert plan["kind"] == "ok"
    assert plan["positions"] == [2, 12]

def test_plan_edit_occurrence_selects_one():
    plan = _plan("a cat and a cat", "cat", occurrence=2)
    assert plan["kind"] == "ok"
    assert plan["positions"] == [12]

def test_plan_edit_occurrence_out_of_range_is_ambiguous():
    plan = _plan("a cat and a cat", "cat", occurrence=5)
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "occurrence_out_of_range"

def test_plan_edit_no_match_offers_closest_partial():
    flat = "the regeneration process is the controller"
    plan = _plan(flat, "the regeneration process is the comptroller")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_match"
    assert "regeneration process" in plan["result"]["options"][0]["context"]


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
    plan = _insert_after_plan(doc, "Anchor paragraph", "NEW PARAGRAPH")
    assert plan["kind"] == "ok"
    assert plan["body_index"] == 1
    assert plan["request"]["insertText"]["location"]["index"] == 35
    assert plan["request"]["insertText"]["text"] == "\nNEW PARAGRAPH"

def test_insert_after_plan_missing_anchor_is_ambiguous():
    plan = _insert_after_plan(_fake_doc(_para(1, "x\n")), "nope", "y")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_match"

def test_insert_after_plan_ambiguous_then_allow_multiple():
    doc = _fake_doc(_para(1, "shared token\n"), _para(14, "shared token again\n"))
    plan = _insert_after_plan(doc, "shared token", "y")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "multiple_matches"
    assert [o["id"] for o in plan["result"]["options"]] == [1, 2]
    allowed = _insert_after_plan(doc, "shared token", "y", require_unique=False)
    assert allowed["kind"] == "ok"
    assert allowed["body_index"] == 0

def test_insert_after_plan_occurrence_selects():
    doc = _fake_doc(_para(1, "shared token\n"), _para(14, "shared token again\n"))
    plan = _insert_after_plan(doc, "shared token", "y", occurrence=2)
    assert plan["kind"] == "ok"
    assert plan["body_index"] == 1

def test_link_plan_builds_update_request():
    doc = _fake_doc(_para(1, "See the Lean Startup here.\n"))
    plan = _link_plan(doc, "Lean Startup", "http://example.com")
    assert plan["kind"] == "ok"
    requests, spans = plan["requests"], plan["spans"]
    assert len(requests) == 1
    style = requests[0]["updateTextStyle"]
    assert style["range"] == {"startIndex": 9, "endIndex": 21}
    assert style["textStyle"]["link"]["url"] == "http://example.com"
    assert style["fields"] == "link"
    assert spans == [{"start_index": 9, "end_index": 21}]

def test_link_plan_missing_text_is_ambiguous():
    plan = _link_plan(_fake_doc(_para(1, "hello\n")), "zzz", "http://e")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_match"

def test_link_plan_ambiguous_then_all_occurrences():
    doc = _fake_doc(_para(1, "go go\n"))
    plan = _link_plan(doc, "go", "http://e")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "multiple_matches"
    allowed = _link_plan(doc, "go", "http://e", all_occurrences=True)
    assert allowed["kind"] == "ok"
    assert len(allowed["requests"]) == 2


# ---------------------------------------------------------------------------
# previously-untested pure helpers
# ---------------------------------------------------------------------------

def test_utf16_len_counts_surrogate_pairs():
    assert _utf16_len("abc") == 3
    assert _utf16_len("café") == 4          # é is BMP, one code unit
    assert _utf16_len("😀") == 2             # astral, surrogate pair
    assert _utf16_len("a😀b") == 4

def test_parse_table_row_splits_and_strips_cells():
    assert _parse_table_row("| a | b | c |") == ["a", "b", "c"]
    with pytest.raises(ValueError):
        _parse_table_row("not a table row")

def test_is_table_separator_detects_separator_rows():
    assert _is_table_separator("| --- | --- |", 2) is True
    assert _is_table_separator("| :-- | --: |", 2) is True
    assert _is_table_separator("| a | b |", 2) is False     # not dashes
    assert _is_table_separator("| --- |", 2) is False        # wrong column count

def test_chapter_filename_is_zero_padded():
    assert _chapter_filename(7) == "chapter-07.xhtml"
    assert _chapter_filename(12) == "chapter-12.xhtml"

def test_text_from_elements_concatenates_and_strips_trailing_newline():
    assert _text_from_elements([{"textRun": {"content": "Hello\n"}}]) == "Hello"
    assert _text_from_elements(
        [{"textRun": {"content": "a"}}, {"textRun": {"content": "b\n"}}]
    ) == "ab"

def test_doc_text_runs_and_index_mapping():
    doc = _fake_doc(_para(1, "abc\n"), _para(5, "de\n"))
    runs = _doc_text_runs(doc)
    assert runs == [(1, "abc\n"), (5, "de\n")]
    assert _doc_index_at(runs, 0) == 1     # first char -> doc index 1
    assert _doc_index_at(runs, 2) == 3
    assert _doc_index_at(runs, 4) == 5     # into the second run

def test_doc_index_at_out_of_range_raises():
    with pytest.raises(IndexError):
        _doc_index_at([(1, "ab\n")], 99)

def test_inline_object_ids_extracts_image_refs():
    para = {"paragraph": {"elements": [
        {"inlineObjectElement": {"inlineObjectId": "kix.a"}},
        {"textRun": {"content": "x"}},
        {"inlineObjectElement": {"inlineObjectId": "kix.b"}},
    ]}}
    assert _inline_object_ids(para) == ["kix.a", "kix.b"]
    assert _inline_object_ids({"paragraph": {"elements": [{"textRun": {"content": "x"}}]}}) == []

def test_block_html_prefers_html_else_escapes_text():
    assert _block_html({"html": "<em>x</em>"}) == "<em>x</em>"
    assert _block_html({"text": "a & b"}) == "a &amp; b"

def test_epub_nav_lists_chapter_links():
    nav = _epub_nav("My Book", [{"filename": "chapter-01.xhtml", "title": "Chapter One"}])
    assert '<a href="chapter-01.xhtml">Chapter One</a>' in nav
    assert "My Book" in nav

def test_default_epub_title_single_vs_multiple():
    assert _default_epub_title(["Solo Chapter"]) == "Solo Chapter"
    assert _default_epub_title(["A", "B"]) == "Ploma Vermella Export"


# ---------------------------------------------------------------------------
# structural check: every CLI subcommand is dispatched in main()
# ---------------------------------------------------------------------------

def test_every_subcommand_is_dispatched():
    parser = _build_parser()
    sub_actions = [a for a in parser._actions if a.__class__.__name__ == "_SubParsersAction"]
    assert sub_actions, "no subparsers found"
    names = list(sub_actions[0].choices)
    src = inspect.getsource(main)
    # "note" is intentionally the else/default branch in main().
    missing = [n for n in names if f'"{n}"' not in src and n != "note"]
    assert not missing, f"subcommands not dispatched in main(): {missing}"


# ---------------------------------------------------------------------------
# EPUB cover / title page / author metadata
# ---------------------------------------------------------------------------

def test_title_page_xhtml_includes_title_subtitle_author():
    xhtml = _title_page_xhtml("My Book", "A Subtitle", "Chris Ford")
    assert '<h1 class="title">My Book</h1>' in xhtml
    assert '<p class="subtitle">A Subtitle</p>' in xhtml
    assert '<p class="author">Chris Ford</p>' in xhtml

def test_title_page_xhtml_omits_missing_fields():
    xhtml = _title_page_xhtml("Only Title")
    assert '<h1 class="title">Only Title</h1>' in xhtml
    assert "subtitle" not in xhtml
    assert "author" not in xhtml

def test_cover_page_xhtml_references_image():
    xhtml = _cover_page_xhtml("images/cover.jpg")
    assert '<img class="cover" src="images/cover.jpg" alt="Cover"/>' in xhtml
    assert 'epub:type="cover"' in xhtml

def test_epub_package_includes_author_creator():
    package = _epub_package(
        "Book", "uuid-1", [{"filename": "chapter-01.xhtml", "title": "Ch1"}],
        author="Chris Ford",
    )
    assert "<dc:creator>Chris Ford</dc:creator>" in package

def test_epub_package_marks_cover_image_and_meta():
    package = _epub_package(
        "Book", "uuid-1", [{"filename": "chapter-01.xhtml", "title": "Ch1"}],
        media_items=[{"id": "cover-image", "href": "images/cover.jpg", "media_type": "image/jpeg"}],
        cover_image_id="cover-image",
    )
    assert 'properties="cover-image"' in package
    assert '<meta name="cover" content="cover-image"/>' in package

def test_epub_package_front_matter_leads_spine():
    package = _epub_package(
        "Book", "uuid-1", [{"filename": "chapter-01.xhtml", "title": "Ch1"}],
        front_matter=[{"id": "titlepage", "href": "title.xhtml"}],
    )
    assert package.index('idref="titlepage"') < package.index('idref="chap1"')


def test_map_comments_flattens_and_filters_resolved():
    raw = [
        {"id": "a", "author": {"displayName": "X"}, "content": "c1",
         "quotedFileContent": {"value": "q1"}, "resolved": False},
        {"id": "b", "content": "c2", "resolved": True},
    ]
    assert _map_comments(raw, include_resolved=False) == [
        {
            "id": "a", "author": "X", "content": "c1",
            "quoted_text": "q1", "resolved": False, "replies": [],
        },
    ]
    both = _map_comments(raw, include_resolved=True)
    assert len(both) == 2
    assert both[1] == {
        "id": "b", "author": "", "content": "c2",
        "quoted_text": "", "resolved": True, "replies": [],
    }


def test_parse_hex_color():
    assert _parse_hex_color("#000000") == {"red": 0.0, "green": 0.0, "blue": 0.0}
    c = _parse_hex_color("d3002d")
    assert round(c["red"], 3) == round(211 / 255, 3)
    assert c["green"] == 0.0
    assert round(c["blue"], 3) == round(45 / 255, 3)
    with pytest.raises(ValueError):
        _parse_hex_color("#fff")

def test_style_plan_builds_request_with_chosen_fields():
    doc = _fake_doc(_para(1, "see Lean Startup now\n"))
    plan = _style_plan(doc, "Lean Startup", italic=True, color="#d3002d")
    assert plan["kind"] == "ok"
    requests, spans = plan["requests"], plan["spans"]
    style = requests[0]["updateTextStyle"]
    assert style["range"] == {"startIndex": 5, "endIndex": 17}
    assert style["textStyle"]["italic"] is True
    assert "rgbColor" in style["textStyle"]["foregroundColor"]["color"]
    assert set(style["fields"].split(",")) == {"italic", "foregroundColor"}
    assert spans == [{"start_index": 5, "end_index": 17}]

def test_style_plan_requires_a_style():
    with pytest.raises(ValueError):
        _style_plan(_fake_doc(_para(1, "text\n")), "text")

def test_style_plan_missing_text_is_ambiguous():
    plan = _style_plan(_fake_doc(_para(1, "hello\n")), "zzz", italic=True)
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_match"


def test_normalize_quotes_folds_curly_to_straight():
    assert _normalize_quotes("don’t say “hi”") == 'don\'t say "hi"'

def test_find_matches_is_quote_agnostic():
    # curly in the doc, straight in the query
    doc = _fake_doc(_para(1, "It’s a “test”\n"))
    m = _find_matches(doc, 'It\'s a "test"')
    assert len(m) == 1
    assert m[0]["start_index"] == 1
    # straight in the doc, curly in the query
    doc2 = _fake_doc(_para(1, "a 'b'\n"))
    assert len(_find_matches(doc2, "a ‘b’")) == 1


def _png_bytes(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def test_downscale_image_resizes_wide_image():
    big = _png_bytes(Image.frombytes("RGB", (2400, 300), os.urandom(2400 * 300 * 3)))
    out, _mt = _downscale_image(big, "image/png", 1600)
    assert Image.open(io.BytesIO(out)).size == (1600, 200)
    assert len(out) < len(big)

def test_downscale_image_photo_becomes_jpeg():
    photo = _png_bytes(Image.frombytes("RGB", (1200, 1000), os.urandom(1200 * 1000 * 3)))
    out, mt = _downscale_image(photo, "image/png", 1600)
    assert mt == "image/jpeg"
    assert len(out) < len(photo)

def test_downscale_image_flat_diagram_stays_png():
    flat = _png_bytes(Image.new("RGB", (1000, 800), (20, 40, 60)))
    _out, mt = _downscale_image(flat, "image/png", 1600)
    assert mt == "image/png"

def test_downscale_image_preserves_alpha_as_png():
    rgba = _png_bytes(Image.new("RGBA", (1200, 1000), (10, 20, 30, 128)))
    _out, mt = _downscale_image(rgba, "image/png", 1600)
    assert mt == "image/png"

def test_downscale_image_passes_through_non_image():
    out, mt = _downscale_image(b"not an image", "image/png", 1600)
    assert out == b"not an image"
    assert mt == "image/png"


# ---------------------------------------------------------------------------
# _outline_from_doc / pv outline
# ---------------------------------------------------------------------------
def _ol_para(text, style="NORMAL_TEXT", start=0, end=0, bullet=False, image_id=None):
    if image_id is not None:
        elements = [{"startIndex": start, "inlineObjectElement": {"inlineObjectId": image_id}}]
    else:
        elements = [{"startIndex": start, "textRun": {"content": text}}]
    paragraph = {"elements": elements, "paragraphStyle": {"namedStyleType": style}}
    if bullet:
        paragraph["bullet"] = {"listId": "L1"}
    return {"startIndex": start, "endIndex": end, "paragraph": paragraph}


OUTLINE_DOC = {"body": {"content": [
    {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
    _ol_para("My Title\n", "TITLE", 1, 10),
    _ol_para("Intro paragraph.\n", "NORMAL_TEXT", 10, 27),
    _ol_para("Section One\n", "HEADING_1", 27, 39),
    _ol_para("Body text.\n", "NORMAL_TEXT", 39, 50),
    _ol_para("\n", "NORMAL_TEXT", 50, 52, image_id="kix.abc"),
    _ol_para("Figure 1-1. A caption.\n", "NORMAL_TEXT", 52, 75),
    _ol_para("A bullet item.\n", "NORMAL_TEXT", 75, 90, bullet=True),
]}}


def test_outline_default_returns_headings_and_images_only():
    items = _outline_from_doc(OUTLINE_DOC)
    kinds = {(it["kind"], it.get("style")) for it in items}
    assert ("heading", "TITLE") in kinds
    assert ("heading", "HEADING_1") in kinds
    assert any(it["kind"] == "image" for it in items)
    assert all(it["kind"] != "paragraph" for it in items)


def test_outline_image_exposes_inline_object_id_and_index():
    img = next(it for it in _outline_from_doc(OUTLINE_DOC) if it["kind"] == "image")
    assert img["inline_object_id"] == "kix.abc"
    assert img["body_index"] == 5
    assert img["start_index"] == 50


def test_outline_full_includes_paragraphs_and_flags_bullets():
    items = _outline_from_doc(OUTLINE_DOC, full=True)
    assert any(it["kind"] == "paragraph" for it in items)
    bullet = next(it for it in items if it.get("bullet"))
    assert bullet["text"] == "A bullet item."
    assert all("start_index" in it for it in items)


def test_outline_full_skips_section_breaks():
    items = _outline_from_doc(OUTLINE_DOC, full=True)
    assert min(it["body_index"] for it in items) == 1


def test_build_parser_outline_full_flag():
    args = _build_parser().parse_args(["outline", "DOC", "--full"])
    assert args.command == "outline"
    assert args.full is True


# ---------------------------------------------------------------------------
# pv heading / pv bullets
# ---------------------------------------------------------------------------
STYLE_DOC = {"body": {"content": [
    _ol_para("Intro.\n", "NORMAL_TEXT", 1, 8),
    _ol_para("My Section\n", "NORMAL_TEXT", 8, 19),
    _ol_para("First point.\n", "NORMAL_TEXT", 19, 32),
    _ol_para("Second point.\n", "NORMAL_TEXT", 32, 46),
    _ol_para("Third point.\n", "NORMAL_TEXT", 46, 59),
    _ol_para("Outro.\n", "NORMAL_TEXT", 59, 66),
]}}


def test_named_style_for_level_maps_levels():
    assert _named_style_for_level("1") == "HEADING_1"
    assert _named_style_for_level("3") == "HEADING_3"
    assert _named_style_for_level("normal") == "NORMAL_TEXT"
    assert _named_style_for_level("Title") == "TITLE"


def test_named_style_for_level_rejects_unknown():
    with pytest.raises(ValueError):
        _named_style_for_level("banner")


def test_heading_plan_sets_named_style_over_paragraph_range():
    plan = _heading_plan(STYLE_DOC, "My Section", "HEADING_1")
    assert plan["kind"] == "ok"
    req = plan["request"]["updateParagraphStyle"]
    assert req["paragraphStyle"]["namedStyleType"] == "HEADING_1"
    assert req["range"] == {"startIndex": 8, "endIndex": 19}
    assert req["fields"] == "namedStyleType"


def test_heading_plan_ambiguous_anchor_reports_options():
    plan = _heading_plan(STYLE_DOC, "point", "HEADING_2")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "multiple_matches"
    assert len(plan["result"]["options"]) == 3


def test_heading_plan_missing_anchor_is_ambiguous():
    plan = _heading_plan(STYLE_DOC, "no such paragraph", "HEADING_1")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_match"


def test_heading_plan_occurrence_selects_one():
    plan = _heading_plan(STYLE_DOC, "point", "HEADING_2", occurrence=2)
    assert plan["kind"] == "ok"
    assert plan["request"]["updateParagraphStyle"]["range"]["startIndex"] == 32


def test_bullets_plan_spans_anchor_range():
    plan = _bullets_plan(STYLE_DOC, "First point", "Third point")
    assert plan["kind"] == "ok"
    req = plan["request"]["createParagraphBullets"]
    assert req["range"] == {"startIndex": 19, "endIndex": 59}
    assert req["bulletPreset"] == "BULLET_DISC_CIRCLE_SQUARE"


def test_bullets_plan_single_paragraph_when_no_end():
    plan = _bullets_plan(STYLE_DOC, "First point")
    rng = plan["request"]["createParagraphBullets"]["range"]
    assert rng == {"startIndex": 19, "endIndex": 32}


def test_bullets_plan_ordered_uses_numbered_preset():
    plan = _bullets_plan(STYLE_DOC, "First point", "Second point", ordered=True)
    preset = plan["request"]["createParagraphBullets"]["bulletPreset"]
    assert preset == "NUMBERED_DECIMAL_ALPHA_ROMAN"


def test_bullets_plan_normalizes_reversed_anchors():
    plan = _bullets_plan(STYLE_DOC, "Third point", "First point")
    assert plan["request"]["createParagraphBullets"]["range"] == {"startIndex": 19, "endIndex": 59}


def test_build_parser_heading_and_bullets():
    a = _build_parser().parse_args(["heading", "DOC", "anchor", "2"])
    assert a.command == "heading" and a.level == "2"
    b = _build_parser().parse_args(["bullets", "DOC", "start", "end", "--ordered"])
    assert b.command == "bullets" and b.ordered is True and b.end == "end"


# ---------------------------------------------------------------------------
# pv replace-image
# ---------------------------------------------------------------------------
FIGURE_DOC = {"body": {"content": [
    _ol_para("Body text before.\n", "NORMAL_TEXT", 1, 19),
    _ol_para("\n", "NORMAL_TEXT", 19, 21, image_id="kix.fig1"),
    _ol_para("Figure 1-1. The first figure.\n", "NORMAL_TEXT", 21, 51),
    _ol_para("More body text.\n", "NORMAL_TEXT", 51, 67),
    _ol_para("\n", "NORMAL_TEXT", 67, 69, image_id="kix.fig2"),
    _ol_para("Figure 1-2. The second figure.\n", "NORMAL_TEXT", 69, 100),
]}}


def test_preceding_image_id_finds_image_above_caption():
    content = FIGURE_DOC["body"]["content"]
    assert _preceding_image_id(content, 2) == "kix.fig1"
    assert _preceding_image_id(content, 5) == "kix.fig2"


def test_preceding_image_id_none_when_text_intervenes():
    content = FIGURE_DOC["body"]["content"]
    assert _preceding_image_id(content, 3) is None


def test_replace_image_plan_resolves_caption_to_object_id():
    plan = _replace_image_plan(FIGURE_DOC, "Figure 1-2.")
    assert plan["kind"] == "ok"
    assert plan["object_id"] == "kix.fig2"
    assert plan["caption_body_index"] == 5


def test_replace_image_plan_missing_caption_is_ambiguous():
    plan = _replace_image_plan(FIGURE_DOC, "Figure 9-9.")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_match"


def test_replace_image_plan_caption_without_image_is_ambiguous():
    doc = {"body": {"content": [
        _ol_para("Figure 3-1. Orphan caption.\n", "NORMAL_TEXT", 1, 30),
    ]}}
    plan = _replace_image_plan(doc, "Figure 3-1.")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_image"


def test_build_parser_replace_image():
    a = _build_parser().parse_args(
        ["replace-image", "DOC", "Figure 1-1.", "DECK", "g123", "--size", "MEDIUM"]
    )
    assert a.command == "replace-image"
    assert a.caption == "Figure 1-1." and a.slide_id == "g123" and a.size == "MEDIUM"
