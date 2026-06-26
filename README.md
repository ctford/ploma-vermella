# Ploma Vermella

CLI tool for reviewing Google Doc book chapters against a style guide, posting feedback as an in-document review section.

---

## Setup

### 1. Python environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
sh install-hooks.sh
```

### 2. Google Cloud — one-time

1. Go to [console.cloud.google.com](https://console.cloud.google.com) and create a project (e.g. `ploma-vermella`).
2. Enable APIs:
   - **Google Docs API**
   - **Google Drive API**
   - **Google Slides API**
   - **Google Sheets API**
3. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**.
   - Application type: **Desktop app**
   - Download the JSON and save it as `credentials/client_secret.json`.
4. Under **OAuth consent screen**, add your Google account as a test user.

On first use, a browser window will open for OAuth authorisation. The token is cached at `credentials/token.json`.

If you add APIs or OAuth scopes later, move `credentials/token.json` aside and rerun a `pv` command to force reauthorisation.

---

## Usage

Tell Claude Code (or any AGENTS.md-aware agent) in plain English:

> "Review chapter 3 at `<doc-url>`."

The agent identifies which **work** the chapter belongs to, reads that work's context (its `context/<work>/README.md` indexes the rest) plus anything in `references/`, fetches the document, and appends a **🪶 Ploma Vermella Review** section with dated, bulleted, located feedback. Agent instructions live in [AGENTS.md](AGENTS.md).

The `pv` CLI is available for direct use (with the virtualenv active):

```bash
source .venv/bin/activate
pv list <folder-url>                        # list docs in a Drive folder
pv fetch <doc-url>                          # fetch title + text of a doc
pv note <doc-url> <quoted-text> <comment>   # append a review note
pv slides-fetch <presentation-url>          # fetch slide text from a deck
pv slides-thumb <presentation-url> <page-id> # get a slide thumbnail URL
pv sheet-fetch <sheet-url> --range ...      # read sheet rows by range
pv sheet-update <sheet-url> <range> ...     # write sheet rows from JSON
pv figure-map <doc-url>                     # inspect image neighborhoods in a doc
pv outline <doc-url> [--full]               # structural map: headings + images, with indices
pv replace-block <doc-url> <start> <end> ... # replace one body-element block safely
pv replace-section <doc-url> <heading> <text> # replace a heading's body up to the next heading
pv insert-image <doc-url> <body-index> ...  # restore an inline image at a body index
pv replace-image <doc-url> <caption> <deck-url> <slide-id>  # re-export a figure from a slide thumbnail
pv place-figure <doc-url> <anchor> <deck-url> <slide-id> --caption ...  # insert a centered figure + caption
pv find <doc-url> <text>                    # locate text: indices, paragraph style, is_code, context
pv insert-after <doc-url> <anchor> <text>   # insert paragraph(s) after an anchor paragraph
pv insert-before <doc-url> <anchor> <text>  # insert paragraph(s) before an anchor paragraph
pv link <doc-url> <text> <url>              # hyperlink a span of text
pv cite <doc-url> <title> <url>             # italicize + hyperlink a work title (citation)
pv heading <doc-url> <anchor> <level>       # set a paragraph's style (1-6, normal, title)
pv bullets <doc-url> <start> [end]          # make a paragraph range a bulleted/numbered list
pv build-epub <doc-url> <doc-url> ...       # build an EPUB (figures preserved) into dist/ with a date suffix
#
# Run `pv -h` or `pv <command> -h` for the full command list.
```

## Figure Workflow

Use the Slides deck as the source of truth for image assets, the Google Doc as the source of truth for current placement, and (optionally) a Google Sheet as the source of truth for logged figure metadata.

Recommended workflow for figure cleanup:

1. `pv figure-map <doc-url>` to identify each inline image and the surrounding prose/caption block.
2. `pv slides-fetch <presentation-url>` to inspect slide text and locate the matching source slide.
3. `pv slides-thumb <presentation-url> <page-id>` if you need a stable image URL for reinsertion.
4. Fix one figure block at a time with `pv replace-block`, then verify by re-running `pv figure-map`.
5. If an image is missing, restore it with `pv insert-image <doc-url> <body-index> <image-url>`.
6. Once the chapter is stable, update any publisher-facing figure-tracking sheet with `pv sheet-update`.

This is intentionally slower than a bulk edit, but it avoids index drift in Google Docs.

---

## Context

Context lives under `context/`, one subdirectory per **work** (a book or other long-running effort). Each work is its own git repo and is self-contained: a `README.md` indexes the directory, a `style_guide.md` holds that work's prose rules, and other files (`outline.md`, `documents.md`, `figures.md`, `notes.md`, …) are described in the README. Longer-lived local reference material (e.g. publisher style guides) goes in `references/`.

The agent loads a work's `README.md` first and follows it to whatever else the task needs — see [AGENTS.md](AGENTS.md) for the full workflow. No code changes needed to add a work.

---

## Limitations (Google Docs API)

Two things the Google Docs / Drive APIs simply do not expose, which shape how `pv` works:

- **No suggested edits.** There is no public API to create tracked "Suggesting mode" changes. Every edit `pv` makes (`pv edit`, `pv replace-block`, `pv insert-after`, `pv link`, etc.) is applied directly, as if in Editing mode. The API can *read* existing suggestions in a document, but it cannot *create* them. If you want a change to land as a suggestion, make it by hand in the Google Docs UI.
- **No text-anchored comments.** There is no public API (Drive API, Docs API, or Apps Script) to create a comment anchored to a specific text range with the yellow in-document highlight. `pv comment` posts to the sidebar with the quoted text (`quotedFileContent`) for context, but without the highlight tying it to a location in the body.

Because of these limits, `pv` favours the in-document **🪶 Ploma Vermella Review** section for located feedback — each note quotes the exact text it refers to, so the reader can find it without a highlight.

---

## Development

```bash
pytest tests/ -v        # run tests
ruff check pv.py tests/  # lint
```

Tests and lint run automatically as a pre-commit hook and on every push via GitHub Actions.
