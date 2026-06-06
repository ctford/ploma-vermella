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

Tell Claude Code in plain English:

> "Review chapter 3 at `<doc-url>`."

Claude will read all files in `context/` and `references/`, fetch the document, and append a **🪶 Ploma Vermella Review** section at the end of the doc with dated, bulleted, located feedback.

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
pv replace-block <doc-url> <start> <end> ... # replace one body-element block safely
pv insert-image <doc-url> <body-index> ...  # restore an inline image at a body index
pv build-epub <doc-url> <doc-url> ...       # build an EPUB from multiple docs into dist/ with a date suffix
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

Drop working review material into `context/` and longer-lived local reference material into `references/` — style guide, chapter outline, author notes. Claude reads both directories on each review. No code changes needed.

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
