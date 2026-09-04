"""Tests for pv.py — pure functions."""

import inspect
import io
import os
from datetime import datetime
from pathlib import Path

import pytest
from PIL import Image

import pv
from pv import (
    _RUNAWAY_TOLERANCE,
    _assign_parts,
    _block_html,
    _blocks_to_xhtml,
    _body_element_at,
    _body_paragraphs,
    _build_parser,
    _bullets_plan,
    _chapter_filename,
    _cite_plan,
    _cover_title_page_xhtml,
    _default_epub_output_path,
    _default_epub_title,
    _default_pdf_output_path,
    _doc_index_at,
    _doc_text_runs,
    _document_code_font,
    _downscale_image,
    _ebook_convert_command,
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
    _insert_before_plan,
    _insert_table_plan,
    _is_code_paragraph,
    _is_image_paragraph,
    _is_table_separator,
    _italic_spans,
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
    _parse_part_spec,
    _parse_table_row,
    _parse_terms,
    _place_figure_requests,
    _plan_edit_matches,
    _preceding_image_id,
    _prose_check_from_doc,
    _prose_structure_checks,
    _prose_text,
    _prose_text_checks,
    _referenced_figures,
    _replace_body_range_plan,
    _replace_image_plan,
    _replace_section_plan,
    _review_copy_title,
    _same_chapter,
    _sentences,
    _shade_plan,
    _shape_text,
    _slugify,
    _style_plan,
    _suggestions_from_doc,
    _table_cell_starts,
    _table_update_plan,
    _terms_italics_scope,
    _text_from_elements,
    _title_page_xhtml,
    _toc_entries_html,
    _toc_page_xhtml,
    _utf16_len,
    _word_count_summary,
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
        {"type": "paragraph", "text": "Body paragraph.", "html": "Body paragraph.",
             "code": False},
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
        {"type": "paragraph", "text": "Figure 1-1. A caption.",
         "html": "Figure 1-1. A caption.", "code": False},
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


def test_default_pdf_output_path_includes_date_suffix():
    path = _default_pdf_output_path(
        "Example Book",
        stamp=datetime(2026, 3, 24, 10, 30),
    )
    assert str(path) == "dist/example-book-20260324.pdf"


def test_ebook_convert_command_includes_paper_size():
    command = _ebook_convert_command(
        Path("/tmp/book.epub"), Path("/tmp/book.pdf"), paper_size="a4",
    )
    assert command[:3] == ["ebook-convert", "/tmp/book.epub", "/tmp/book.pdf"]
    assert "--paper-size" in command
    assert command[command.index("--paper-size") + 1] == "a4"
    assert "--pdf-no-cover" in command
    assert "--pdf-add-toc" in command
    assert command[command.index("--toc-title") + 1] == "Contents"


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

def test_document_code_font_finds_the_doc_s_own_monospace():
    doc = _fake_doc(
        _para(1, "Prose here.\n"),
        _para(13, "x = 1\n", font="Roboto Mono"),
        _para(20, "y = 2\n", font="Roboto Mono"),
        _para(27, "z\n", font="Courier New"),
    )
    assert _document_code_font(doc) == "Roboto Mono"

def test_document_code_font_falls_back_when_no_code_present():
    assert _document_code_font(_fake_doc(_para(1, "All prose.\n"))) == "Roboto Mono"
    doc = _fake_doc(_para(1, "All prose.\n"))
    assert _document_code_font(doc, default="Menlo") == "Menlo"

def test_insert_after_plan_code_font_styles_only_the_inserted_text():
    doc = _fake_doc(_para(1, "Intro line.\n"), _para(13, "Anchor paragraph here.\n"))
    plan = _insert_after_plan(doc, "Anchor paragraph", "x = 1", code_font="Roboto Mono")
    assert plan["kind"] == "ok"
    insert, style = plan["requests"]
    assert insert is plan["request"]
    assert insert["insertText"]["text"] == "\nx = 1"
    # The separator keeps the block off its anchor and must stay unstyled.
    assert style["updateTextStyle"]["range"] == {"startIndex": 36, "endIndex": 41}
    family = style["updateTextStyle"]["textStyle"]["weightedFontFamily"]["fontFamily"]
    assert family == "Roboto Mono"

def test_insert_after_plan_without_code_font_sends_one_request():
    doc = _fake_doc(_para(1, "Anchor paragraph here.\n"))
    plan = _insert_after_plan(doc, "Anchor paragraph", "prose")
    assert plan["requests"] == [plan["request"]]

def test_insert_before_plan_code_font_styles_the_inserted_text():
    doc = _fake_doc(_para(1, "Target paragraph.\n"))
    plan = _insert_before_plan(doc, "Target paragraph", "x = 1", code_font="Menlo")
    insert, style = plan["requests"]
    assert insert["insertText"]["text"] == "x = 1\n"
    assert style["updateTextStyle"]["range"] == {"startIndex": 1, "endIndex": 6}
    fam = style["updateTextStyle"]["textStyle"]["weightedFontFamily"]["fontFamily"]
    assert fam == "Menlo"

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
    assert style["fields"] == "link,foregroundColor"
    assert spans == [{"start_index": 9, "end_index": 21}]


def test_link_plan_applies_the_house_red_and_can_be_told_not_to():
    """Links are set in the O'Reilly red by default; --no-color leaves them blue."""
    doc = _fake_doc(_para(1, "See the Lean Startup here.\n"))
    red = _link_plan(doc, "Lean Startup", "http://e")["requests"][0]["updateTextStyle"]
    rgb = red["textStyle"]["foregroundColor"]["color"]["rgbColor"]
    assert round(rgb["red"], 3) == round(0xD3 / 255, 3)
    assert rgb["green"] == 0.0
    assert round(rgb["blue"], 3) == round(0x2D / 255, 3)
    plain = _link_plan(doc, "Lean Startup", "http://e", color=None)["requests"][0]
    assert plain["updateTextStyle"]["fields"] == "link"
    assert "foregroundColor" not in plain["updateTextStyle"]["textStyle"]

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

def test_doc_text_runs_descends_into_table_cells():
    def cell(start, text):
        return {"content": [_para(start, text)]}
    doc = _fake_doc(
        _para(1, "Before.\n"),
        {"table": {"tableRows": [
            {"tableCells": [cell(12, "Kiro\n"), cell(20, "Spec-kit\n")]},
        ]}},
        _para(40, "After.\n"),
    )
    runs = _doc_text_runs(doc)
    assert runs == [(1, "Before.\n"), (12, "Kiro\n"), (20, "Spec-kit\n"), (40, "After.\n")]
    # a match inside a table cell resolves to that cell's document index
    flat = "".join(t for _, t in runs)
    assert _doc_index_at(runs, flat.index("Spec-kit")) == 20

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

def test_cover_title_page_xhtml_combines_image_and_title():
    xhtml = _cover_title_page_xhtml("My Book", "images/cover.jpg", "A Subtitle", "Chris Ford")
    assert '<img class="cover" src="images/cover.jpg" alt="Cover"/>' in xhtml
    assert 'epub:type="cover titlepage"' in xhtml
    assert '<h1 class="title">My Book</h1>' in xhtml
    assert '<p class="subtitle">A Subtitle</p>' in xhtml
    assert '<p class="author">Chris Ford</p>' in xhtml

def test_toc_page_xhtml_lists_chapter_links():
    xhtml = _toc_page_xhtml([{"filename": "chapter-01.xhtml", "title": "Ch1"}])
    assert '<a href="chapter-01.xhtml">Ch1</a>' in xhtml
    assert 'epub:type="toc"' in xhtml

def test_toc_entries_html_groups_chapters_under_their_part():
    chapters = [
        {"filename": "chapter-01.xhtml", "title": "Preface", "part": None},
        {"filename": "chapter-02.xhtml", "title": "Ch1", "part": "Part I: Reverse Engineering"},
        {"filename": "chapter-03.xhtml", "title": "Ch2", "part": "Part I: Reverse Engineering"},
        {"filename": "chapter-04.xhtml", "title": "Ch3", "part": "Part II: Forward Engineering"},
    ]
    html_out = _toc_entries_html(chapters)
    assert '<li><a href="chapter-01.xhtml">Preface</a></li>' in html_out
    assert '<li class="toc-part">Part I: Reverse Engineering' in html_out
    assert '<li class="toc-part">Part II: Forward Engineering' in html_out
    assert '<a href="chapter-02.xhtml">Ch1</a>' in html_out
    assert html_out.index("Part I") < html_out.index("Ch1") < html_out.index("Part II")

def test_parse_part_spec_splits_title_and_doc_id():
    title, doc_id = _parse_part_spec(
        "Part I: Reverse Engineering=https://docs.google.com/document/d/abc123/edit"
    )
    assert title == "Part I: Reverse Engineering"
    assert doc_id == "abc123"

def test_parse_part_spec_requires_equals():
    with pytest.raises(ValueError):
        _parse_part_spec("Part I: Reverse Engineering")

def test_assign_parts_runs_until_next_start_doc():
    doc_ids = ["preface", "ch1", "ch2", "ch3"]
    part_specs = [("Part I: Reverse Engineering", "ch1"), ("Part II: Forward Engineering", "ch3")]
    assert _assign_parts(doc_ids, part_specs) == [
        None, "Part I: Reverse Engineering", "Part I: Reverse Engineering",
        "Part II: Forward Engineering",
    ]

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

def test_style_plan_can_turn_a_style_off():
    doc = _fake_doc(_para(1, "Term\n"))
    plan = _style_plan(doc, "Term", bold=False)
    assert plan["kind"] == "ok"
    style = plan["requests"][0]["updateTextStyle"]
    assert style["textStyle"] == {"bold": False}
    assert style["fields"] == "bold"

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
# _suggestions_from_doc / pv suggestions
# ---------------------------------------------------------------------------

def _suggested_run(text, insert=None, delete=None):
    run = {"textRun": {"content": text}}
    if insert:
        run["textRun"]["suggestedInsertionIds"] = insert
    if delete:
        run["textRun"]["suggestedDeletionIds"] = delete
    return run

# A title with a suggested "Chapter 7. " insertion, a replacement (delete "or",
# insert "and"), an unchanged paragraph, and a table cell with a suggested deletion.
SUGGESTIONS_DOC = {
    "title": "Chapter 07: Technology Iteration",
    "body": {
        "content": [
            {"sectionBreak": {}},
            {"paragraph": {"elements": [
                _suggested_run("Chapter 7. ", insert=["sug.a"]),
                _suggested_run("Technology Iteration\n"),
            ]}},
            {"paragraph": {"elements": [
                _suggested_run("different"),
                _suggested_run(" and", insert=["sug.b"]),
                _suggested_run(", or", delete=["sug.b"]),
                _suggested_run(" to keep it the same.\n"),
            ]}},
            {"paragraph": {"elements": [_suggested_run("Untouched paragraph.\n")]}},
            {"table": {"tableRows": [{"tableCells": [{"content": [
                {"paragraph": {"elements": [
                    _suggested_run("cell before"),
                    _suggested_run(" extra", delete=["sug.c"]),
                ]}},
            ]}]}]}},
            {"paragraph": {
                "elements": [_suggested_run("Restyled heading\n")],
                "suggestedParagraphStyleChanges": {"sug.d": {}},
            }},
            # Style-only: italicize a term as a term, and relink an attribution.
            # Neither run carries an insertion or deletion, so a text-only diff
            # misses both entirely.
            {"paragraph": {"elements": [
                {"textRun": {
                    "content": "setpoint",
                    "textStyle": {},
                    "suggestedTextStyleChanges": {"sug.e": {
                        "textStyle": {"italic": True},
                        "textStyleSuggestionState": {"italicSuggested": True},
                    }},
                }},
                {"textRun": {
                    "content": " describes",
                    "textStyle": {"link": {"url": "https://old.example/x"}},
                    "suggestedTextStyleChanges": {"sug.f": {
                        "textStyle": {
                            "link": {"url": "https://new.example/y"},
                            "foregroundColor": {"color": {"rgbColor": {
                                "red": 0.06666667, "green": 0.33333334, "blue": 0.8,
                            }}},
                        },
                        "textStyleSuggestionState": {
                            "linkSuggested": True, "foregroundColorSuggested": True,
                        },
                    }},
                }},
                _suggested_run(" the loop.\n"),
            ]}},
            {"paragraph": {
                "elements": [_suggested_run("Rebulleted item\n")],
                "suggestedBulletChanges": {"sug.g": {}},
            }},
        ]
    },
}


def test_suggestions_reports_only_changed_paragraphs():
    result = _suggestions_from_doc(SUGGESTIONS_DOC)
    # Untouched paragraph is excluded; the other six carry suggestions.
    assert result["changed_paragraphs"] == 6
    assert all("Untouched" not in p["before"] for p in result["paragraphs"])


def test_suggestions_before_after_and_marked_for_insertion():
    title = _suggestions_from_doc(SUGGESTIONS_DOC)["paragraphs"][0]
    assert title["before"] == "Technology Iteration"
    assert title["after"] == "Chapter 7. Technology Iteration"
    assert title["marked"] == "{+Chapter 7. +}Technology Iteration"


def test_suggestions_replacement_splits_into_delete_and_insert():
    para = _suggestions_from_doc(SUGGESTIONS_DOC)["paragraphs"][1]
    assert para["before"] == "different, or to keep it the same."
    assert para["after"] == "different and to keep it the same."
    assert para["marked"] == "different{+ and+}[-, or-] to keep it the same."


def test_suggestions_walks_into_table_cells():
    result = _suggestions_from_doc(SUGGESTIONS_DOC)
    cell = next(p for p in result["paragraphs"] if p["before"].startswith("cell before"))
    assert cell["after"] == "cell before"
    assert cell["marked"] == "cell before[- extra-]"


def test_suggestions_counts_and_style_changes():
    result = _suggestions_from_doc(SUGGESTIONS_DOC)
    assert result["insertion_count"] == 2  # sug.a, sug.b
    assert result["deletion_count"] == 2   # sug.b, sug.c
    assert result["paragraph_style_change_count"] == 1  # sug.d
    styled = next(p for p in result["paragraphs"] if p["style_change"])
    assert styled["style_change"] == ["sug.d"]


def test_suggestions_total_is_distinct_ids_not_the_sum():
    """A replacement is one suggestion spanning a deleted and an inserted run.

    sug.b sits in both buckets, so insertion_count + deletion_count double-counts it.
    Distinct IDs: a, b, c, d, e, f, g."""
    result = _suggestions_from_doc(SUGGESTIONS_DOC)
    assert result["insertion_count"] + result["deletion_count"] == 4
    assert result["total_suggestion_count"] == 7


def test_suggestions_reports_style_only_paragraphs():
    """A run can carry formatting suggestions with no text change at all."""
    result = _suggestions_from_doc(SUGGESTIONS_DOC)
    para = next(p for p in result["paragraphs"] if p["before"].startswith("setpoint"))
    # No text moved, so before and after match.
    assert para["before"] == para["after"] == "setpoint describes the loop."
    assert [e["text"] for e in para["text_style_edits"]] == ["setpoint", "describes"]
    assert result["text_style_change_count"] == 2  # sug.e, sug.f


def test_suggestions_text_style_edit_reports_property_deltas():
    result = _suggestions_from_doc(SUGGESTIONS_DOC)
    para = next(p for p in result["paragraphs"] if p["before"].startswith("setpoint"))
    italic, relink = para["text_style_edits"]
    assert italic["changes"] == [{"property": "italic", "from": None, "to": True}]
    # Links and colors are flattened to the URL and a hex code.
    assert relink["changes"] == [
        {"property": "foregroundColor", "from": None, "to": "#1155cc"},
        {"property": "link", "from": "https://old.example/x", "to": "https://new.example/y"},
    ]


def test_suggestions_reports_bullet_changes():
    result = _suggestions_from_doc(SUGGESTIONS_DOC)
    para = next(p for p in result["paragraphs"] if p["before"] == "Rebulleted item")
    assert para["bullet_change"] == ["sug.g"]
    assert result["bullet_change_count"] == 1


def test_suggestions_empty_doc_has_no_paragraphs():
    result = _suggestions_from_doc({})
    assert result["changed_paragraphs"] == 0
    assert result["paragraphs"] == []
    assert result["total_suggestion_count"] == 0


def test_build_parser_suggestions():
    args = _build_parser().parse_args(["suggestions", "DOC"])
    assert args.command == "suggestions"
    assert args.doc == "DOC"


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

def test_bullets_plan_remove_uses_delete_request():
    plan = _bullets_plan(STYLE_DOC, "First point", "Third point", remove=True)
    assert plan["kind"] == "ok"
    assert "createParagraphBullets" not in plan["request"]
    assert plan["request"]["deleteParagraphBullets"]["range"] == {"startIndex": 19, "endIndex": 59}


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


# ---------------------------------------------------------------------------
# pv place-figure
# ---------------------------------------------------------------------------
def test_place_figure_requests_builds_centered_image_and_caption():
    reqs = _place_figure_requests(100, "https://img", "Figure 2-1. A caption.", 300.0, 200.0)
    assert reqs[0]["insertText"]["location"]["index"] == 100
    assert reqs[0]["insertText"]["text"] == "\n\n\nFigure 2-1. A caption.\n"
    assert reqs[1]["insertInlineImage"]["location"]["index"] == 102
    assert reqs[1]["insertInlineImage"]["uri"] == "https://img"
    size = reqs[1]["insertInlineImage"]["objectSize"]
    assert size["width"]["magnitude"] == 300.0
    assert size["height"]["unit"] == "PT"
    style = reqs[2]["updateParagraphStyle"]
    assert style["paragraphStyle"]["alignment"] == "CENTER"
    assert style["range"] == {"startIndex": 102, "endIndex": 103}
    caption_len = len("Figure 2-1. A caption.")
    caption_range = {"startIndex": 104, "endIndex": 104 + caption_len}
    caption_align = reqs[3]["updateParagraphStyle"]
    assert caption_align["paragraphStyle"]["alignment"] == "CENTER"
    assert caption_align["range"] == caption_range
    caption_italic = reqs[4]["updateTextStyle"]
    assert caption_italic["textStyle"]["italic"] is True
    assert caption_italic["range"] == caption_range


def test_build_parser_place_figure():
    a = _build_parser().parse_args(
        ["place-figure", "DOC", "anchor", "DECK", "g1", "--caption", "Figure 2-1."]
    )
    assert a.command == "place-figure"
    assert a.caption == "Figure 2-1." and a.slide_id == "g1"


# ---------------------------------------------------------------------------
# pv cite
# ---------------------------------------------------------------------------
def test_cite_plan_applies_italic_and_link():
    doc = {"body": {"content": [_ol_para("See The Coal Question now.\n", "NORMAL_TEXT", 1, 28)]}}
    plan = _cite_plan(doc, "The Coal Question", "https://oreilly/x")
    assert plan["kind"] == "ok"
    req = plan["requests"][0]["updateTextStyle"]
    assert req["textStyle"]["italic"] is True
    assert req["textStyle"]["link"]["url"] == "https://oreilly/x"
    assert req["fields"] == "italic,link,foregroundColor"
    assert req["textStyle"]["foregroundColor"]["color"]["rgbColor"]["green"] == 0.0


def test_cite_plan_ambiguous_when_title_repeats():
    doc = {"body": {"content": [_ol_para("Cloud FinOps vs Cloud FinOps.\n", "NORMAL_TEXT", 1, 31)]}}
    plan = _cite_plan(doc, "Cloud FinOps", "https://x")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "multiple_matches"


def test_cite_plan_occurrence_selects_one():
    doc = {"body": {"content": [_ol_para("Cloud FinOps vs Cloud FinOps.\n", "NORMAL_TEXT", 1, 31)]}}
    plan = _cite_plan(doc, "Cloud FinOps", "https://x", occurrence=2)
    assert plan["kind"] == "ok"
    assert len(plan["requests"]) == 1


def test_build_parser_cite():
    a = _build_parser().parse_args(["cite", "DOC", "Title", "URL", "--occurrence", "2"])
    assert a.command == "cite" and a.occurrence == 2


# ---------------------------------------------------------------------------
# pv insert-before
# ---------------------------------------------------------------------------
def test_insert_before_plan_inserts_at_paragraph_start():
    doc = {"body": {"content": [
        _ol_para("First.\n", "NORMAL_TEXT", 1, 8),
        _ol_para("Target paragraph.\n", "NORMAL_TEXT", 8, 26),
    ]}}
    plan = _insert_before_plan(doc, "Target paragraph", "New line.")
    assert plan["kind"] == "ok"
    req = plan["request"]["insertText"]
    assert req["location"]["index"] == 8
    assert req["text"] == "New line.\n"
    assert plan["body_index"] == 1


def test_insert_before_plan_ambiguous_anchor():
    doc = {"body": {"content": [
        _ol_para("a point\n", "NORMAL_TEXT", 1, 9),
        _ol_para("a point\n", "NORMAL_TEXT", 9, 17),
    ]}}
    plan = _insert_before_plan(doc, "point", "x")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "multiple_matches"


def test_build_parser_insert_before():
    a = _build_parser().parse_args(["insert-before", "DOC", "anchor", "text", "--occurrence", "3"])
    assert a.command == "insert-before" and a.occurrence == 3


# ---------------------------------------------------------------------------
# pv replace-section
# ---------------------------------------------------------------------------
SECTION_DOC = {"body": {"content": [
    _ol_para("Title\n", "TITLE", 1, 7),
    _ol_para("Alpha\n", "HEADING_1", 7, 13),
    _ol_para("Old alpha body.\n", "NORMAL_TEXT", 13, 29),
    _ol_para("More old.\n", "NORMAL_TEXT", 29, 39),
    _ol_para("Sub\n", "HEADING_2", 39, 43),
    _ol_para("Sub body.\n", "NORMAL_TEXT", 43, 53),
    _ol_para("Beta\n", "HEADING_1", 53, 58),
    _ol_para("Beta body.\n", "NORMAL_TEXT", 58, 69),
]}}


def test_replace_section_plan_spans_to_next_same_level_heading():
    plan = _replace_section_plan(SECTION_DOC, "Alpha", "New body.")
    assert plan["kind"] == "ok"
    assert plan["section_start"] == 13   # after the Alpha heading
    assert plan["section_end"] == 53     # start of Beta — includes the H2 subsection
    assert plan["insert_text"] == "New body.\n\n"


def test_replace_section_plan_last_section_preserves_final_newline():
    plan = _replace_section_plan(SECTION_DOC, "Beta", "New beta.")
    assert plan["kind"] == "ok"
    assert plan["section_start"] == 58
    assert plan["section_end"] == 68     # last endIndex (69) minus the final newline
    assert plan["insert_text"] == "New beta."


def test_replace_section_plan_only_matches_headings():
    plan = _replace_section_plan(SECTION_DOC, "Old alpha", "x")
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_match"


# ---------------------------------------------------------------------------
# pv replace-block
# ---------------------------------------------------------------------------
RANGE_DOC = {"body": {"content": [
    {"startIndex": 0, "endIndex": 1, "sectionBreak": {}},
    _ol_para("Title\n", "TITLE", 1, 7),
    _ol_para("Old opener.\n", "NORMAL_TEXT", 7, 19),
    _ol_para("More opener.\n", "NORMAL_TEXT", 19, 32),
    _ol_para("Heading\n", "HEADING_1", 32, 40),
    _ol_para("Body.\n", "NORMAL_TEXT", 40, 46),
]}}


def test_replace_body_range_plan_appends_trailing_newline():
    # Without this the last new paragraph merges into the paragraph that follows.
    plan = _replace_body_range_plan(RANGE_DOC, 2, 3, "New opener.")
    assert plan["insert_text"] == "New opener.\n"
    delete, insert, _style = plan["requests"]
    assert delete["deleteContentRange"]["range"] == {"startIndex": 7, "endIndex": 32}
    assert insert["insertText"]["location"]["index"] == 7


def test_replace_body_range_plan_keeps_existing_trailing_newline():
    plan = _replace_body_range_plan(RANGE_DOC, 2, 3, "One.\n\nTwo.\n")
    assert plan["insert_text"] == "One.\n\nTwo.\n"


def test_replace_body_range_plan_reapplies_replaced_block_style():
    # The insert lands at the start of the HEADING_1 that follows, so the new
    # paragraphs would inherit HEADING_1 without this.
    plan = _replace_body_range_plan(RANGE_DOC, 2, 3, "New opener.")
    style = plan["requests"][2]["updateParagraphStyle"]
    assert style["paragraphStyle"]["namedStyleType"] == "NORMAL_TEXT"
    assert style["range"] == {"startIndex": 7, "endIndex": 7 + len("New opener.\n")}


def test_replace_body_range_plan_preserves_a_replaced_heading_style():
    plan = _replace_body_range_plan(RANGE_DOC, 4, 4, "New heading")
    style = plan["requests"][2]["updateParagraphStyle"]
    assert style["paragraphStyle"]["namedStyleType"] == "HEADING_1"


def test_replace_body_range_plan_delete_only_emits_no_insert():
    plan = _replace_body_range_plan(RANGE_DOC, 2, 2, "")
    assert plan["insert_text"] == ""
    assert len(plan["requests"]) == 1


def test_replace_body_range_plan_at_end_preserves_final_newline():
    plan = _replace_body_range_plan(RANGE_DOC, 5, 5, "Last.")
    assert plan["insert_text"] == "Last."
    assert plan["requests"][0]["deleteContentRange"]["range"]["endIndex"] == 45


def test_replace_body_range_plan_rejects_out_of_range():
    with pytest.raises(ValueError):
        _replace_body_range_plan(RANGE_DOC, 2, 99, "x")


def test_build_parser_replace_block():
    a = _build_parser().parse_args(["replace-block", "DOC", "2", "5", "text"])
    assert a.command == "replace-block" and a.start_body_index == 2


def test_build_parser_replace_section():
    a = _build_parser().parse_args(["replace-section", "DOC", "Alpha", "new text"])
    assert a.command == "replace-section" and a.heading == "Alpha"


def test_build_parser_no_command_allowed():
    # Bare `pv` must parse (command None) so main() can print help, not error out.
    args = _build_parser().parse_args([])
    assert args.command is None


# ---------------------------------------------------------------------------
# _prose_check_from_doc / pv prose-check
# ---------------------------------------------------------------------------

def _named(check_list, name):
    return next(c for c in check_list if c["check"] == name)


def _styled_run(text, italic=False):
    run = {"textRun": {"content": text, "textStyle": {}}}
    if italic:
        run["textRun"]["textStyle"]["italic"] = True
    return run


def _check_para(elements, style="NORMAL_TEXT"):
    return {"paragraph": {
        "elements": elements, "paragraphStyle": {"namedStyleType": style},
    }}


def _cell(text, bold=False):
    style = {"bold": True} if bold else {}
    return {"content": [{
        "startIndex": 0, "endIndex": 0,
        "paragraph": {"elements": [{"textRun": {"content": text, "textStyle": style}}]},
    }]}


def _table_doc(grid, bold_header=True, start=100):
    """A doc holding one table, with plausible cell indices."""
    idx = start + 2
    rows = []
    for r, row in enumerate(grid):
        cells = []
        for text in row:
            content = text + "\n"
            cell = _cell(text, bold=bold_header and r == 0)
            cell["content"][0]["startIndex"] = idx + 1
            cell["content"][0]["endIndex"] = idx + 1 + len(content)
            idx += 2 + len(content)
            cells.append(cell)
        rows.append({"tableCells": cells})
    table = {"rows": len(grid), "columns": len(grid[0]), "tableRows": rows}
    return {"body": {"content": [
        {"startIndex": start, "endIndex": idx, "table": table},
    ]}}


def test_table_update_plan_rewrites_every_cell_back_to_front():
    doc = _table_doc([["Section", "Words"], ["Preface", "3,093"]])
    plan = _table_update_plan(
        doc, "Section", [["Section", "Words"], ["Preface", "3,138"]]
    )
    assert plan["kind"] == "ok"
    assert (plan["rows"], plan["columns"]) == (2, 2)
    # Later cells are rewritten first, so earlier cells' indices stay valid.
    inserts = [r["insertText"] for r in plan["requests"] if "insertText" in r]
    starts = [i["location"]["index"] for i in inserts]
    assert starts == sorted(starts, reverse=True)
    assert [i["text"] for i in inserts] == ["3,138", "Preface", "Words", "Section"]


def test_table_update_plan_reapplies_a_bold_header():
    doc = _table_doc([["Section", "Words"], ["Preface", "3,093"]])
    plan = _table_update_plan(
        doc, "Section", [["Section", "Words"], ["Preface", "3,138"]]
    )
    styled = [r["updateTextStyle"] for r in plan["requests"] if "updateTextStyle" in r]
    # Only the header row carried bold, so only its two cells get it back.
    assert len(styled) == 2
    assert all(s["textStyle"]["bold"] is True for s in styled)
    assert all(s["fields"] == "bold" for s in styled)


def test_table_update_plan_refuses_to_reshape_the_table():
    doc = _table_doc([["Section", "Words"], ["Preface", "3,093"]])
    plan = _table_update_plan(doc, "Section", [["Section", "Words"]])
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "shape_mismatch"
    assert "2x2" in plan["result"]["message"] and "1x2" in plan["result"]["message"]


def test_table_update_plan_no_match_is_ambiguous():
    doc = _table_doc([["Section", "Words"]])
    plan = _table_update_plan(doc, "Nothing like this", [["a", "b"]])
    assert plan["kind"] == "ambiguous"
    assert plan["result"]["reason"] == "no_match"


def test_prose_check_surfaces_the_rate_denominators():
    """The Conclusion's metrics table needs re-weighting across chapters, so the
    counts behind each percentage have to be readable, not inferred from the rate."""
    doc = {"title": "t", "body": {"content": [
        _check_para([_styled_run("Heading here\n")], style="HEADING_1"),
        _check_para([_styled_run("One two three four five six. Seven eight nine.\n")]),
        _check_para([_styled_run("Ten eleven twelve thirteen fourteen.\n")]),
        _check_para([_styled_run("Short one.\n")]),
    ]}}
    result = _prose_check_from_doc(doc)
    # The two denominators do not filter alike, which is the reason to report both:
    # the heading is excluded from prose entirely, but "Short one." is a sentence
    # while being too short to count as a running-prose paragraph.
    assert result["sentence_count"] == 4
    assert result["paragraph_count"] == 2
    over = _named(result["checks"], "paragraphs_over_120_words")["value"]
    assert over.startswith("0 (")


def test_prose_check_word_counts_separate_prose_from_tables():
    """`word_count` is the whole document; `prose_word_count` is what the sentence
    mean is actually computed over, so an overall mean can be re-derived exactly."""
    doc = {"title": "t", "body": {"content": [
        _check_para([_styled_run("Alpha beta gamma delta epsilon zeta.\n")]),
        {"table": {"tableRows": [{"tableCells": [
            {"content": [_check_para([_styled_run("cell one\n")])]},
            {"content": [_check_para([_styled_run("cell two\n")])]},
        ]}]}},
    ]}}
    result = _prose_check_from_doc(doc)
    assert result["prose_word_count"] == 6
    assert result["word_count"] > result["prose_word_count"]


def test_unapplied_statuses_cover_the_no_change_results():
    """A structured "nothing happened" result must be distinguishable from success.

    `pv edit` returns {"status": "ambiguous"} when the anchor does not match. That is a
    deliberate result shape, but it printed and exited 0, so a batch driver testing the
    exit code reported silent no-ops as applied edits.
    """
    from pv import _UNAPPLIED_STATUSES
    assert "ambiguous" in _UNAPPLIED_STATUSES
    assert "not_found" in _UNAPPLIED_STATUSES
    assert "edited" not in _UNAPPLIED_STATUSES
    assert "replaced" not in _UNAPPLIED_STATUSES
    assert "linked" not in _UNAPPLIED_STATUSES


def test_referenced_figures_expands_a_range():
    """"Figures 11-5 through 11-7" names three figures but only two numbers."""
    refs = _referenced_figures("See Figures 11-5 through 11-7 for the breakdown.")
    assert refs == {"11-5", "11-6", "11-7"}


def test_referenced_figures_handles_and_and_dashes():
    assert _referenced_figures("Figures 3-1 and 3-2") == {"3-1", "3-2"}
    assert _referenced_figures("Figures 3-1\u20133-3") == {"3-1", "3-2", "3-3"}


def test_referenced_figures_does_not_confuse_a_prefix():
    """"Figure 1-1" must not count as a reference to "Figure 1-10"."""
    assert "1-10" not in _referenced_figures("As shown in Figure 1-1, the loop closes.")


def test_harness_roles_flag_a_lowercase_enumeration():
    """The measured defect is whole runs set in lower case, not isolated slips."""
    text = "It provides services that cross-cut guides, guards, sensors, and checks."
    check = _named(_prose_text_checks(text), "harness_roles_capitalized")
    assert check["status"] == "review"
    assert "guides, guards, sensors, and checks" in check["detail"][0]


def test_harness_roles_accept_a_capitalized_enumeration():
    text = "Your combination of Guides, Guards, Sensors, and Checks lowers the cost."
    assert _named(
        _prose_text_checks(text), "harness_roles_capitalized")["status"] == "ok"


def test_harness_roles_ignore_the_ordinary_verbs():
    """"checks" and "guides" are common words; a bare search for them is unusable."""
    text = (
        "The linter checks the code before it lands. An architecture diagram guides "
        "the developer to the right module. The pipeline guards against regressions."
    )
    assert _named(
        _prose_text_checks(text), "harness_roles_capitalized")["status"] == "ok"


def test_harness_roles_need_two_distinct_roles_in_one_run():
    """One role repeated is prose about that role, not the taxonomy being named."""
    text = "The checks ran, and further checks confirmed it. Later checks passed too."
    assert _named(
        _prose_text_checks(text), "harness_roles_capitalized")["status"] == "ok"


def test_harness_roles_flag_a_partly_lowercased_run():
    """A mixed run is the slip mid-sentence the editor flagged."""
    text = "Unlike Guides, guards come with a verdict on what is acceptable."
    assert _named(
        _prose_text_checks(text), "harness_roles_capitalized")["status"] == "review"


def test_inclusive_we_reports_a_density_not_a_ban():
    """A stray "we" in a long chapter is fine; a chapter built on it is not."""
    # One instance in a chapter-sized run of prose stays under the threshold.
    clean = ("You should tell the agent what good looks like. " * 130) + "We agreed."
    assert _named(_prose_text_checks(clean), "inclusive_we")["status"] == "ok"

    heavy = "We carve our problem into pieces so we can iterate on it ourselves. " * 8
    check = _named(_prose_text_checks(heavy), "inclusive_we")
    assert check["status"] == "review"
    assert "per 1,000 words" in check["value"]


def test_inclusive_we_catches_contractions_and_let_us():
    """"we're" and "let's" are the same voice problem as a bare "we"."""
    text = "We\u2019re going to ship it, so let\u2019s agree on our approach."
    check = _named(_prose_text_checks(text), "inclusive_we")
    assert check["status"] == "review"
    # we're, let's, our
    assert check["value"].startswith("3 ")


def test_inclusive_we_gives_context_so_a_reviewer_can_classify():
    """The disambiguation rule differs per instance, so a bare count is not enough."""
    text = "We saw in Part III that agents drift. " * 6
    detail = _named(_prose_text_checks(text), "inclusive_we")["detail"]
    assert detail and all("saw in Part III" in d for d in detail)


def test_prose_text_checks_flags_long_sentences_and_mean():
    checks = _prose_text_checks(" ".join(["word"] * 40) + ".")
    assert _named(checks, "sentences_over_35_words")["value"].startswith("1 ")
    assert _named(checks, "sentences_over_35_words")["status"] == "review"
    assert _named(checks, "sentence_length_mean")["value"] == 40.0
    assert _named(checks, "sentence_length_mean")["status"] == "review"


def test_long_sentences_are_budgeted_not_banned():
    """Sarah signed off on a chapter with 46 of them, so zero would cry wolf."""
    long_one = " ".join(["word"] * 40) + "."
    padding = " ".join("A short ordinary sentence here." for _ in range(40))
    assert _named(_prose_text_checks(long_one + " " + padding),
                  "sentences_over_35_words")["status"] == "ok"


def test_prose_text_checks_mean_within_target_passes():
    text = " ".join("A sentence of about eleven plain ordinary words here." for _ in range(5))
    assert _named(_prose_text_checks(text), "sentence_length_mean")["status"] == "ok"


def test_prose_text_checks_em_dash_density_and_nested_asides():
    text = "A sentence — with one aside — and a second clause. " + " ".join(["filler"] * 20)
    checks = _prose_text_checks(text)
    assert _named(checks, "sentences_with_two_em_dashes")["value"].startswith("1 ")
    assert _named(checks, "sentences_with_two_em_dashes")["status"] == "review"
    assert _named(checks, "em_dash_density")["status"] == "review"


def test_insert_after_matches_the_document_paragraph_spacing():
    """A doc that separates paragraphs with a blank line gets one before the insert too."""
    spaced = _fake_doc(_para(1, "Anchor here.\n"), _para(15, "\n"), _para(16, "Next.\n"))
    plan = _insert_after_plan(spaced, "Anchor here.", "New text.")
    assert plan["request"]["insertText"]["text"] == "\n\nNew text."
    tight = _fake_doc(_para(1, "Anchor here.\n"), _para(15, "Next.\n"))
    plan = _insert_after_plan(tight, "Anchor here.", "New text.")
    assert plan["request"]["insertText"]["text"] == "\nNew text."


def test_prose_text_checks_em_dash_density_flags_only_crowding():
    """Author's call 2026-09-03: sparse em-dashes are plainer prose, not a defect.

    The upper bound was removed after it failed 6 of 14 sections for writing that
    reads fine. Crowding is still a finding.
    """
    dense = " ".join(["word"] * 80) + " — one aside."
    assert _named(_prose_text_checks(dense), "em_dash_density")["status"] == "review"
    sparse = " ".join(["word"] * 400) + " — one aside."
    assert _named(_prose_text_checks(sparse), "em_dash_density")["status"] == "ok"
    none_at_all = " ".join(["word"] * 400) + "."
    assert _named(_prose_text_checks(none_at_all), "em_dash_density")["status"] == "ok"


def test_sentences_do_not_run_across_paragraph_breaks():
    """Fragment bullets take no full stop, so a list must not read as one sentence."""
    bullets = "\n".join(f"A bullet item number {i} without a full stop" for i in range(1, 7))
    assert len(_sentences(bullets)) == 6
    assert max(len(s.split()) for s in _sentences(bullets)) == 9


def test_sentence_ceiling_gates_separately_from_the_budget():
    """45 words is the gate; the 35-word measure stays a budget, as with code health."""
    short = "A short sentence. " * 12
    runaway = " ".join(["word"] * 50) + "."
    checks = _prose_text_checks(short + runaway)
    gate = _named(checks, "sentences_over_45_words")
    assert gate["value"].startswith("1 ") and gate["status"] == "review"
    assert _named(_prose_text_checks(short), "sentences_over_45_words")["status"] == "ok"


def test_runaway_gate_is_a_tolerance_not_a_ceiling():
    """A rate, not a ceiling. The rate itself is _RUNAWAY_TOLERANCE, currently 5%."""
    one_long = " ".join(["word"] * 50) + "."
    inside = "A short sentence here. " * 99
    ok = _named(_prose_text_checks(inside + one_long), "sentences_over_45_words")
    assert ok["status"] == "ok", "1 in 100 is inside the tolerance"
    # Just over the line: 1 of 20 is 5.0%, and the gate is strictly-under.
    outside = "A short sentence here. " * 19
    bad = _named(_prose_text_checks(outside + one_long), "sentences_over_45_words")
    assert bad["status"] == "review", "1 in 20 is exactly 5%, which does not pass"
    # And comfortably inside at the old 2% bar, to pin that this is a loosening.
    assert _RUNAWAY_TOLERANCE == 5.0
    mid = "A short sentence here. " * 39
    assert _named(_prose_text_checks(mid + one_long),
                  "sentences_over_45_words")["status"] == "ok", "1 in 40 = 2.5%"


def test_sentence_split_handles_closing_delimiters_and_links():
    """A sentence can end behind a bracket, a quote, or a linked title.

    Measured on the manuscript 2026-09-03: 15 of 146 runaway sentences were two
    sentences welded by this, not long prose. Both shapes are common in cited text.
    """
    assert len(_sentences("Ends inside.) The next one starts here.")) == 2
    assert len(_sentences(
        "See [Who Needs an Architect?](http://example.com/a) At least for now.")) == 2
    assert len(_sentences('He said "stop." Then he left.')) == 2
    assert len(_sentences("Plain one. Plain two.")) == 2
    # A mid-sentence link must not split, or every citation becomes two sentences.
    assert len(_sentences("As [Morris](http://example.com/b) writes, it is so.")) == 1


def test_prose_metrics_ignore_a_rendered_table():
    """A table renders as 'a | b' with no full stop, so counting it wrecks the mean."""
    prose = "Short sentence here. Another short one follows it now.\n"
    rows = "\n".join(f"Ch {i} | {i}000 | 20.{i}" for i in range(1, 13))
    table = "Chapter | Words | Mean\n" + rows
    full = prose + table
    # A line break now ends a sentence, so rows are no longer swallowed into one
    # giant sentence. They are still not prose, and counting thirteen of them as
    # sentences distorts any per-sentence measure.
    assert len(_sentences(full)) == 15, "two sentences plus thirteen table rows"
    assert len(_sentences(prose)) == 2
    counted = _named(_prose_text_checks(full), "sentence_length_mean")["value"]
    honest = _named(_prose_text_checks(full, prose=prose), "sentence_length_mean")["value"]
    assert counted != honest, "the table must not be able to move the prose measure"


def test_extract_blocks_keeps_tables():
    """Tables have no `paragraph` key, so the loop used to skip them entirely."""
    def cell(text):
        return {"content": [{"paragraph": {"elements": [
            {"textRun": {"content": text, "textStyle": {}}}]}}]}
    doc = {"body": {"content": [
        {"paragraph": {"elements": [{"textRun": {"content": "Before.\n", "textStyle": {}}}]}},
        {"table": {"tableRows": [
            {"tableCells": [cell("Thread"), cell("Chapters")]},
            {"tableCells": [cell("Control theory"), cell("4, 5, 6")]},
        ]}},
    ]}}
    blocks = _extract_blocks(doc)
    kinds = [b["type"] for b in blocks]
    assert kinds == ["paragraph", "table"]
    assert blocks[1]["rows"] == [["Thread", "Chapters"], ["Control theory", "4, 5, 6"]]


def test_blocks_to_xhtml_renders_a_table_with_a_header_row():
    blocks = [{"type": "table", "rows": [["Thread", "Chapters"], ["Control theory", "4, 5"]]}]
    out = _blocks_to_xhtml("Conclusion", blocks)
    assert "<table>" in out and "</table>" in out
    assert "<th>Thread</th><th>Chapters</th>" in out
    bold = [{"type": "table", "rows": [["<strong>Thread</strong>"], ["a"]]}]
    assert "<th>Thread</th>" in _blocks_to_xhtml("C", bold), "th is already emphasised"
    assert "<td>Control theory</td><td>4, 5</td>" in out


def test_insert_table_plan_places_and_validates():
    doc = _fake_doc(_para(1, "Anchor paragraph.\n"), _para(20, "After.\n"))
    plan = _insert_table_plan(doc, "Anchor", [["a", "b"], ["c", "d"]])
    assert plan["kind"] == "ok"
    assert plan["rows"] == 2 and plan["columns"] == 2
    assert plan["index"] == 19, "defaults to just after the anchor paragraph"
    before = _insert_table_plan(doc, "Anchor", [["a", "b"]], before=True)
    assert before["index"] == 1


def test_insert_table_plan_rejects_ragged_rows():
    doc = _fake_doc(_para(1, "Anchor.\n"))
    with pytest.raises(ValueError, match="same number of cells"):
        _insert_table_plan(doc, "Anchor", [["a", "b"], ["c"]])
    with pytest.raises(ValueError, match="at least one row"):
        _insert_table_plan(doc, "Anchor", [])


def test_table_cell_starts_reads_in_reading_order():
    """Cells must come back in reading order so the fill can run in reverse."""
    def cell(i):
        return {"content": [{"startIndex": i, "paragraph": {"elements": []}}]}
    table = {"tableRows": [{"tableCells": [cell(5), cell(9)]},
                           {"tableCells": [cell(13), cell(17)]}]}
    assert _table_cell_starts(table) == [5, 9, 13, 17]


def test_resolve_all_uses_the_paginating_fetcher(monkeypatch):
    """Resolve-all must see every page, not just the API's default first 20."""
    pages = [
        {"comments": [{"id": "a", "resolved": False}], "nextPageToken": "t2"},
        {"comments": [{"id": "b", "resolved": False}, {"id": "c", "resolved": True}]},
    ]
    calls = {"list": 0, "resolved": []}

    class FakeReplies:
        def create(self, **kwargs):
            calls["resolved"].append(kwargs["commentId"])
            return self

        def execute(self):
            return {"id": "r"}

    class FakeComments:
        def list(self, **kwargs):
            self._page = pages[calls["list"]]
            calls["list"] += 1
            return self

        def execute(self):
            return self._page

    class FakeService:
        def comments(self):
            return FakeComments()

        def replies(self):
            return FakeReplies()

    monkeypatch.setattr(pv, "_drive_service", lambda: FakeService())
    result = pv.resolve_all_comments("https://docs.google.com/document/d/abc123")
    assert calls["list"] == 2, "should have followed nextPageToken"
    assert sorted(calls["resolved"]) == ["a", "b"]
    assert result["count"] == 2


def test_needles_match_on_word_boundaries():
    """'able to' must not fire on 'unable to', nor 'analyse' on the US plural 'analyses'."""
    checks = _prose_text_checks("The team was unable to run the analyses it had planned.")
    assert _named(checks, "tic_phrases")["value"] == 0
    assert _named(checks, "uk_spellings")["value"] == 0
    checks = _prose_text_checks("The team was able to run it, and analysed the artefact.")
    assert "able tox1" in _named(checks, "tic_phrases")["detail"]
    assert sorted(_named(checks, "uk_spellings")["detail"]) == ["analysedx1", "artefactx1"]


def test_acronym_check_ignores_part_numbers():
    """Part II and Part IV are numbers, not acronyms to spell out."""
    checks = _prose_text_checks("Part II covers this, Part IV the rest, and the VPC stays.")
    assert _named(checks, "acronyms_to_verify")["detail"] == ["VPC"]


def test_spelling_checks_skip_code_paragraphs_and_captions():
    """A UK spelling inside a code snippet is the snippet's, not a correction to make."""
    def para(text, mono=False):
        style = {"weightedFontFamily": {"fontFamily": "Courier New"}} if mono else {}
        return {"paragraph": {"paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                              "elements": [{"textRun": {"content": text, "textStyle": style}}]}}
    doc = {"body": {"content": [
        para("The order was CANCELLED and the colour analysed.\n", mono=True),
        para("Figure 8-1. A caption mentioning an artefact in passing.\n"),
        para("The team shipped the feature on time.\n"),
    ]}}
    prose = _prose_text(doc)
    assert "CANCELLED" not in prose and "artefact" not in prose
    assert "shipped the feature" in prose
    checks = _prose_text_checks(_extract_text(doc), prose=prose)
    assert _named(checks, "uk_spellings")["value"] == 0


def test_prose_text_checks_finds_passives_tics_and_uk_forms():
    text = "The report was written in order to explain the artefact."
    checks = _prose_text_checks(text)
    assert _named(checks, "passive_constructions")["value"].startswith("1 (")
    assert "in order tox1" in _named(checks, "tic_phrases")["detail"]
    assert "artefactx1" in _named(checks, "uk_spellings")["detail"]


def test_passive_check_gates_on_the_share_of_sentences_not_the_count():
    """Author's call 2026-09-03: pin the gate at 12.5% of sentences (goal 10%).

    Chapter 7 is the chapter Sarah signed off and it measures 10.7%, so a gate
    below that would fail the chapter the targets are calibrated on.
    """
    clean = " ".join(["The team ships it."] * 9 + ["The report was written."])
    passive_check = _named(_prose_text_checks(clean), "passive_constructions")
    assert passive_check["value"].startswith("1 (10.0% of sentences)")
    assert passive_check["status"] == "ok"

    crowded = " ".join(["The report was written."] * 2 + ["The team ships it."] * 6)
    crowded_check = _named(_prose_text_checks(crowded), "passive_constructions")
    assert crowded_check["value"].startswith("2 (25.0% of sentences)")
    assert crowded_check["status"] == "review"


def test_prose_text_checks_flags_ly_ordinals_and_spares_the_bare_form():
    """Sarah recasts Firstly/Secondly to First/Second; the bare form must not flag."""
    text = "Firstly, it limits work. Secondly, it traces back. First, check the bare form."
    detail = _named(_prose_text_checks(text), "ordinal_adverbs")["detail"]
    assert detail == ["firstlyx1", "secondlyx1"]


def test_uk_watchlist_catches_doubled_l_but_spares_final_stress():
    """UK doubles a final L; US only when the stress is on the last syllable."""
    checks = _prose_text_checks("Threat modelling of a controlled, compelled system.")
    detail = _named(checks, "uk_spellings")["detail"]
    assert detail == ["modellingx1"]


def test_flagged_phrases_counts_and_is_quote_agnostic():
    """The list is per-work; smart and straight apostrophes must match alike."""
    text = "In today’s world it shows up, and it shows up again. Here’s the thing."
    checks = _prose_text_checks(text, ["shows up", "in today's", "here's the thing"])
    detail = _named(checks, "flagged_phrases")["detail"]
    assert "shows up x2" in detail
    assert "in today's x1" in detail
    assert "here's the thing x1" in detail


def test_flagged_phrases_check_absent_without_a_list():
    names = [c["check"] for c in _prose_text_checks("Plain prose.")]
    assert "flagged_phrases" not in names


def test_flagged_phrases_does_not_match_inside_a_longer_word():
    checks = _prose_text_checks("The realms of possibility.", ["realm"])
    assert _named(checks, "flagged_phrases")["status"] == "ok"


def test_prose_text_checks_acronyms_skip_the_assumed_list():
    checks = _prose_text_checks(
        "The API and the LLM talk to the VPC over OIDC, and the CTO saw the UI."
    )
    detail = _named(checks, "acronyms_to_verify")["detail"]
    assert detail == ["OIDC", "VPC"]


def test_prose_text_checks_spelled_numbers_and_double_spaces():
    checks = _prose_text_checks("It ran for twenty years.  Then it stopped.")
    assert _named(checks, "spelled_numbers_over_nine")["detail"] == ["twenty"]
    assert _named(checks, "double_spaces")["value"] == 1


def test_italic_spans_merges_adjacent_runs():
    doc = {"body": {"content": [_check_para([
        _styled_run("A "),
        _styled_run("set", italic=True),
        _styled_run("point", italic=True),
        _styled_run(" is a target.\n"),
    ])]}}
    assert _italic_spans(doc) == [(2, "setpoint")]


def test_parse_terms_reads_bare_terms_and_home_chapters():
    lines = ["# comment", "", "setpoint = 06", "plain term", "  spaced = 05  "]
    assert _parse_terms(lines) == [
        ("setpoint", "06"), ("plain term", None), ("spaced", "05"),
    ]


def test_terms_italics_scope_reports_only_unitalicized_first_use():
    text = "A setpoint and a controller."
    spans = [(2, "setpoint")]
    missing, away = _terms_italics_scope(
        text, spans, [("setpoint", None), ("controller", None), ("absent term", None)],
    )
    # setpoint is italic at its first use; controller is not; the third never appears.
    assert missing == ["controller"]
    assert away == []


def test_terms_italics_scope_skips_terms_introduced_in_another_chapter():
    """The rule is per-book: a borrowed term stays plain outside its home chapter."""
    text = "A setpoint appears here too."
    missing, away = _terms_italics_scope(
        text, [], [("setpoint", "06")], chapter="07",
    )
    assert missing == []
    assert away == []


def test_terms_italics_scope_flags_a_borrowed_term_set_in_italics():
    text = "A setpoint appears here too."
    missing, away = _terms_italics_scope(
        text, [(2, "setpoint")], [("setpoint", "06")], chapter="07",
    )
    assert away == ["setpoint"]


def test_style_plan_sets_a_monospace_family():
    """Added 2026-09-04: repairing a code line's lost typeface needed the raw API.

    Google Docs sometimes leaves the first character of a code paragraph in the body
    font. `_is_code_paragraph` then rejects the whole paragraph and the EPUB renders
    the listing in the body serif, so the repair has to set a font family.
    """
    doc = _fake_doc(_para(1, "check(\n"))
    plan = _style_plan(doc, "check(", monospace="Roboto Mono")
    style = plan["requests"][0]["updateTextStyle"]
    assert style["textStyle"]["weightedFontFamily"] == {"fontFamily": "Roboto Mono"}
    assert style["fields"] == "weightedFontFamily"


def test_style_plan_combines_monospace_with_other_styles():
    plan = _style_plan(_fake_doc(_para(1, "check(\n")), "check(",
                       italic=True, monospace="Menlo")
    style = plan["requests"][0]["updateTextStyle"]
    assert style["textStyle"]["italic"] is True
    assert style["textStyle"]["weightedFontFamily"] == {"fontFamily": "Menlo"}
    assert set(style["fields"].split(",")) == {"italic", "weightedFontFamily"}


def test_style_plan_still_requires_at_least_one_style():
    with pytest.raises(ValueError, match="monospace"):
        _style_plan(_fake_doc(_para(1, "text\n")), "text")


def test_blocks_to_xhtml_renders_code_paragraphs_as_one_pre_block():
    """Code has to keep its typeface in the EPUB and PDF.

    Before 2026-09-04 `_extract_blocks` never carried `_is_code_paragraph` through, so
    every listing rendered as one ordinary <p> per line in the body serif, with no
    monospace rule anywhere in the stylesheet.
    """
    out = _blocks_to_xhtml("Ch", [
        {"type": "paragraph", "text": "Prose.", "html": "Prose.", "code": False},
        {"type": "paragraph", "text": "check(", "html": "check(", "code": True},
        {"type": "paragraph", "text": "  x == 1", "html": "  x == 1", "code": True},
        {"type": "paragraph", "text": ")", "html": ")", "code": True},
        {"type": "paragraph", "text": "After.", "html": "After.", "code": False},
    ])
    assert out.count("<pre><code>") == 1          # consecutive lines coalesce
    assert out.count("</code></pre>") == 1
    assert "check(\n  x == 1\n)\n" in out         # order and indentation kept
    assert "<p>Prose.</p>" in out and "<p>After.</p>" in out


def test_inline_html_wraps_monospace_runs_in_code():
    """Inline code loses its typeface unless the run's font family is honoured."""
    out = _inline_html([
        {"textRun": {"content": "Set ", "textStyle": {}}},
        {"textRun": {"content": "AGENTS.md", "textStyle": {
            "weightedFontFamily": {"fontFamily": "Roboto Mono"}}}},
        {"textRun": {"content": " first.", "textStyle": {}}},
    ])
    assert out == "Set <code>AGENTS.md</code> first."


def test_blocks_to_xhtml_renders_sidebar_markers_as_a_styled_aside():
    """`<SIDEBAR>`…`</SIDEBAR>` are structure, not prose.

    Before 2026-09-04 the build passed them through and 20 literal markers appeared in
    the rendered EPUB and PDF.
    """
    blocks = [
        {"type": "paragraph", "text": "Before."},
        {"type": "paragraph", "text": "<SIDEBAR>"},
        {"type": "heading", "level": 2, "text": "Example: Knowledge graph"},
        {"type": "paragraph", "text": "Inside."},
        {"type": "paragraph", "text": "</SIDEBAR>"},
        {"type": "paragraph", "text": "After."},
    ]
    out = _blocks_to_xhtml("Ch", blocks)
    assert "SIDEBAR" not in out
    assert '<aside class="sidebar">' in out
    assert out.count("</aside>") == 1
    assert out.index('<aside class="sidebar">') < out.index("Inside.") < out.index("</aside>")
    assert out.index("Before.") < out.index('<aside class="sidebar">')
    assert out.index("</aside>") < out.index("After.")


def test_blocks_to_xhtml_closes_an_unterminated_sidebar():
    """A missing `</SIDEBAR>` must not produce unbalanced XHTML."""
    out = _blocks_to_xhtml("Ch", [
        {"type": "paragraph", "text": "<sidebar>"},
        {"type": "paragraph", "text": "Inside."},
    ])
    assert out.count('<aside class="sidebar">') == 1
    assert out.count("</aside>") == 1


def test_find_matches_reports_italic_and_bold():
    """`pv find` has to surface character styling, not just paragraph style.

    Fixing "italicized more than once" and "italicized away from home chapter" means
    removing italics from a *particular* later occurrence, and until 2026-09-04 there
    was no way to see which occurrence carried them. `_italic_spans` cannot answer it:
    it measures in rendered-text offsets, which drift from document indices by the
    markup of every preceding link.
    """
    doc = _fake_doc({
        "startIndex": 1,
        "paragraph": {"elements": [
            {"startIndex": 1, "textRun": {"content": "plain ", "textStyle": {}}},
            {"startIndex": 7, "textRun": {"content": "damping", "textStyle": {"italic": True}}},
            {"startIndex": 14, "textRun": {"content": " and ", "textStyle": {}}},
            {"startIndex": 19, "textRun": {"content": "bold", "textStyle": {"bold": True}}},
            {"startIndex": 23, "textRun": {"content": ".\n", "textStyle": {}}},
        ]},
    })
    assert _find_matches(doc, "damping")[0]["italic"] is True
    assert _find_matches(doc, "damping")[0]["bold"] is False
    assert _find_matches(doc, "bold")[0]["bold"] is True
    assert _find_matches(doc, "plain")[0]["italic"] is False


def test_terms_italics_scope_tolerates_zero_padded_chapter_numbers():
    """terms.txt writes `= 05`; the CLI is naturally given `--chapter 5`.

    Measured 2026-09-04: a raw string compare reported 26 of Chapter 5's own terms as
    "italicized away from their home chapter" under `--chapter 5`, and 0 under
    `--chapter 05`. Both forms must agree.
    """
    assert _same_chapter("05", "5")
    assert _same_chapter("5", "05")
    assert _same_chapter("13", "13")
    assert not _same_chapter("05", "6")
    assert not _same_chapter("11", "1")


def test_terms_italics_scope_still_checks_a_term_in_its_home_chapter():
    text = "A setpoint appears here."
    missing, away = _terms_italics_scope(
        text, [], [("setpoint", "06")], chapter="06",
    )
    assert missing == ["setpoint"]


def test_prose_check_flags_term_italicized_only_on_a_later_use():
    doc = {"body": {"content": [_check_para([
        _styled_run("The damping idea matters. Later we call it "),
        _styled_run("damping", italic=True),
        _styled_run(" again.\n"),
    ])]}}
    checks = _prose_check_from_doc(doc)["checks"]
    assert _named(checks, "italics_not_on_first_use")["detail"] == ["damping"]


def test_repeated_italics_ignores_stressed_function_words():
    """Italics mark a term once, but also carry emphasis — "and" may recur."""
    doc = {"body": {"content": [_check_para([
        _styled_run("Design "), _styled_run("and", italic=True),
        _styled_run(" build, not design "), _styled_run("and", italic=True),
        _styled_run(" hope.\n"),
    ])]}}
    checks = _prose_check_from_doc(doc)["checks"]
    assert _named(checks, "terms_italicized_more_than_once")["value"] == 0


def test_repeated_italics_still_flags_a_real_term():
    doc = {"body": {"content": [_check_para([
        _styled_run("A "), _styled_run("setpoint", italic=True),
        _styled_run(" and later another "), _styled_run("setpoint", italic=True),
        _styled_run(".\n"),
    ])]}}
    checks = _prose_check_from_doc(doc)["checks"]
    assert _named(checks, "terms_italicized_more_than_once")["detail"] == ["setpoint"]


def test_paragraph_length_skips_headings_bullets_and_fragments():
    """Only running prose counts: a heading, a list item and a stub are not paragraphs."""
    def para(text, style="NORMAL_TEXT", bullet=False):
        p = {"paragraphStyle": {"namedStyleType": style},
             "elements": [{"textRun": {"content": text, "textStyle": {}}}]}
        if bullet:
            p["bullet"] = {"listId": "l1"}
        return {"paragraph": p}
    body = " ".join(["word"] * 50) + "\n"
    doc = {"body": {"content": [
        para("A heading that runs to several words here\n", style="HEADING_1"),
        para(body),
        para("a bulleted item of some length here\n", bullet=True),
        para("Too short\n"),
    ]}}
    assert _body_paragraphs(doc) == [body.strip()]
    checks = _prose_structure_checks(doc, body)
    assert _named(checks, "paragraph_length_mean")["value"] == 50.0
    assert _named(checks, "paragraphs_over_120_words")["value"].startswith("0 ")


def test_paragraph_check_gates_on_the_120_word_ceiling_not_the_mean():
    """The work's rule is a ceiling per paragraph; a good mean can hide a long one."""
    def para(text):
        return {"paragraph": {"paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                              "elements": [{"textRun": {"content": text, "textStyle": {}}}]}}
    short = " ".join(["word"] * 10) + "\n"
    long_one = " ".join(["word"] * 150) + "\n"
    doc = {"body": {"content": [para(short)] * 5 + [para(long_one)]}}
    checks = _prose_structure_checks(doc, short * 5 + long_one)
    assert _named(checks, "paragraph_length_mean")["status"] == "ok", "mean is reported only"
    over = _named(checks, "paragraphs_over_120_words")
    assert over["value"].startswith("1 ") and over["status"] == "review"


def test_italic_spans_line_up_with_the_rendered_text_after_a_link():
    """A link renders as [text](url); italic offsets must count those characters."""
    doc = {"body": {"content": [{"paragraph": {"elements": [
        {"textRun": {"content": "See ", "textStyle": {}}},
        {"textRun": {"content": "Tidy First", "textStyle": {"link": {"url": "http://x"}}}},
        {"textRun": {"content": " on ", "textStyle": {}}},
        {"textRun": {"content": "accretive", "textStyle": {"italic": True}}},
        {"textRun": {"content": " growth.\n", "textStyle": {}}},
    ]}}]}}
    text = _extract_text(doc)
    spans = _italic_spans(doc)
    assert [text[start:start + len(raw)] for start, raw in spans] == ["accretive"]


def test_prose_check_ignores_headings_when_locating_first_use():
    """A term named in a heading is not its first use; the body italics aren't late."""
    doc = {"body": {"content": [
        {"paragraph": {
            "paragraphStyle": {"namedStyleType": "HEADING_1"},
            "elements": [{"textRun": {"content": "S-type systems\n", "textStyle": {}}}],
        }},
        {"paragraph": {
            "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
            "elements": [
                {"textRun": {"content": "An ", "textStyle": {}}},
                {"textRun": {"content": "S-type", "textStyle": {"italic": True}}},
                {"textRun": {"content": " program is derivable.\n", "textStyle": {}}},
            ],
        }},
    ]}}
    text = "S-type systems\nAn S-type program is derivable.\n"
    checks = _prose_structure_checks(doc, text, terms=[("S-type", None)])
    assert _named(checks, "terms_missing_italics_on_first_use")["value"] == 0
    assert _named(checks, "italics_not_on_first_use")["value"] == 0


def test_prose_check_flags_stacked_headings():
    doc = {"body": {"content": [
        _check_para([_styled_run("Chapter 5. Guides\n")], style="HEADING_1"),
        _check_para([_styled_run("Guides\n")], style="HEADING_2"),
        _check_para([_styled_run("Body text at last.\n")]),
    ]}}
    checks = _prose_check_from_doc(doc)["checks"]
    assert _named(checks, "headings_with_no_body_between")["value"] == 1


def test_prose_check_accepts_a_heading_followed_by_body():
    doc = {"body": {"content": [
        _check_para([_styled_run("Guides\n")], style="HEADING_2"),
        _check_para([_styled_run("Body text.\n")]),
        _check_para([_styled_run("Guards\n")], style="HEADING_2"),
    ]}}
    checks = _prose_check_from_doc(doc)["checks"]
    assert _named(checks, "headings_with_no_body_between")["status"] == "ok"


def test_prose_check_flags_figure_caption_with_no_earlier_reference():
    doc = {"body": {"content": [
        _check_para([_styled_run("Some lead-in prose with no pointer.\n")]),
        _check_para([_styled_run("Figure 5-1. The harness quadrant.\n")]),
    ]}}
    checks = _prose_check_from_doc(doc)["checks"]
    assert _named(checks, "figures_not_referenced_before_caption")["detail"] == ["Figure 5-1"]


def test_prose_check_accepts_figure_referenced_before_its_caption():
    doc = {"body": {"content": [
        _check_para([_styled_run("The quadrant is shown in Figure 5-1.\n")]),
        _check_para([_styled_run("Figure 5-1. The harness quadrant.\n")]),
    ]}}
    checks = _prose_check_from_doc(doc)["checks"]
    assert _named(checks, "figures_not_referenced_before_caption")["status"] == "ok"


def test_prose_check_omits_the_terms_check_when_no_list_is_given():
    doc = {"body": {"content": [_check_para([_styled_run("Plain prose.\n")])]}}
    names = [c["check"] for c in _prose_check_from_doc(doc)["checks"]]
    assert "terms_missing_italics_on_first_use" not in names
    names_with = [
        c["check"] for c in _prose_check_from_doc(doc, [("prose", None)])["checks"]
    ]
    assert "terms_missing_italics_on_first_use" in names_with


def test_prose_check_reports_flagged_count_and_judgement_list():
    doc = {"body": {"content": [_check_para([_styled_run("The artefact was written.\n")])]}}
    result = _prose_check_from_doc(doc)
    assert result["flagged"] == sum(1 for c in result["checks"] if c["status"] == "review")
    assert result["flagged"] > 0
    assert any("citation" in item for item in result["needs_a_reader"])


# ---------------------------------------------------------------------------
# _shade_plan / pv shade
# ---------------------------------------------------------------------------

def _marked_doc(*texts, start=1):
    content, idx = [], start
    for t in texts:
        content.append({
            "startIndex": idx, "endIndex": idx + len(t),
            "paragraph": {"elements": [{"textRun": {"content": t}}]},
        })
        idx += len(t)
    return {"body": {"content": content}}


SHADE_DOC = _marked_doc(
    "Intro paragraph\n", "<sidebar>\n", "Sidebar body\n", "</sidebar>\n",
    "Middle prose\n", "<sidebar>\n", "Second body\n", "</sidebar>\n", "Outro\n",
)


def test_shade_plan_covers_the_inclusive_marker_range():
    plan = _shade_plan(SHADE_DOC, "<sidebar>", "</sidebar>", "#efefef")
    assert plan["kind"] == "ok"
    assert plan["count"] == 1
    first = SHADE_DOC["body"]["content"][1]
    closer = SHADE_DOC["body"]["content"][3]
    assert plan["ranges"][0] == {
        "startIndex": first["startIndex"], "endIndex": closer["endIndex"],
    }


def test_shade_plan_all_pairs_shades_every_block():
    plan = _shade_plan(SHADE_DOC, "<sidebar>", "</sidebar>", "#efefef", all_pairs=True)
    assert plan["count"] == 2
    assert plan["ranges"][0]["endIndex"] < plan["ranges"][1]["startIndex"]


def test_shade_plan_does_not_treat_a_closer_as_an_opener():
    """'</sidebar>' contains '<sidebar>' as a substring; openers must exclude closers."""
    plan = _shade_plan(SHADE_DOC, "<sidebar>", "</sidebar>", all_pairs=True)
    assert plan["count"] == 2


def test_shade_plan_refuses_unbalanced_markers():
    doc = _marked_doc("<sidebar>\n", "Body\n", "<sidebar>\n", "More\n", "</sidebar>\n")
    plan = _shade_plan(doc, "<sidebar>", "</sidebar>", all_pairs=True)
    assert plan["kind"] == "unbalanced"
    assert plan["start_matches"] == 2 and plan["end_matches"] == 1


def test_shade_plan_refuses_a_closer_before_its_opener():
    doc = _marked_doc("</sidebar>\n", "Body\n", "<sidebar>\n")
    plan = _shade_plan(doc, "<sidebar>", "</sidebar>")
    assert plan["kind"] == "crossed"


def test_shade_plan_reports_missing_markers():
    plan = _shade_plan(_marked_doc("Just prose\n"), "<sidebar>", "</sidebar>")
    assert plan["kind"] == "not_found"


def test_shade_plan_remove_clears_the_background():
    plan = _shade_plan(SHADE_DOC, "<sidebar>", "</sidebar>", remove=True)
    style = plan["requests"][0]["updateParagraphStyle"]["paragraphStyle"]
    assert style["shading"] == {"backgroundColor": {}}


def test_shade_plan_defaults_to_light_grey():
    plan = _shade_plan(SHADE_DOC, "<sidebar>", "</sidebar>")
    rgb = (plan["requests"][0]["updateParagraphStyle"]["paragraphStyle"]
           ["shading"]["backgroundColor"]["color"]["rgbColor"])
    assert round(rgb["red"], 4) == round(0xEF / 255, 4)


def test_shade_plan_is_case_insensitive_about_markers():
    doc = _marked_doc("<SIDEBAR>\n", "Body\n", "</SIDEBAR>\n")
    assert _shade_plan(doc, "<sidebar>", "</sidebar>")["kind"] == "ok"


def test_build_parser_shade():
    args = _build_parser().parse_args(
        ["shade", "DOC", "<sidebar>", "</sidebar>", "--color", "#eeeeee", "--all"],
    )
    assert args.command == "shade" and args.all_pairs is True
    assert args.color == "#eeeeee"


# ---------------------------------------------------------------------------
# _word_count_summary / pv words
# ---------------------------------------------------------------------------

WORD_ENTRIES = [
    ("Chapter 01: Legacy", "one two three four five"),
    ("Chapter 02: Context", "six seven eight"),
    ("[Old] Chapter 01: Superseded", "stale words here that should not count"),
    ("Table of Contents", "toc"),
]


def test_word_count_summary_totals_every_document_by_default():
    r = _word_count_summary(WORD_ENTRIES)
    assert r["document_count"] == 4
    assert r["total_words"] == 5 + 3 + 7 + 1
    assert r["excluded"] == []


def test_word_count_summary_excludes_by_case_insensitive_substring():
    r = _word_count_summary(WORD_ENTRIES, ["[old]", "table of contents"])
    assert r["document_count"] == 2
    assert r["total_words"] == 8
    assert [d["name"] for d in r["excluded"]] == [
        "[Old] Chapter 01: Superseded", "Table of Contents",
    ]


def test_word_count_summary_reports_what_it_skipped():
    """A total that silently drops documents is worse than no total."""
    r = _word_count_summary(WORD_ENTRIES, ["[old]"])
    assert r["excluded_words"] == 7
    assert r["excluded"][0]["words"] == 7


def test_word_count_summary_handles_an_empty_folder():
    r = _word_count_summary([])
    assert r == {"documents": [], "document_count": 0, "total_words": 0,
                 "excluded": [], "excluded_words": 0}


def test_build_parser_words():
    args = _build_parser().parse_args(
        ["words", "FOLDER", "--exclude", "[Old]", "--exclude", "Index"],
    )
    assert args.command == "words"
    assert args.exclude == ["[Old]", "Index"]


def test_build_parser_prose_check():
    args = _build_parser().parse_args(
        ["prose-check", "DOC", "--terms", "t.txt", "--chapter", "07"],
    )
    assert args.command == "prose-check"
    assert args.doc == "DOC"
    assert args.terms == "t.txt"
    assert args.chapter == "07"
