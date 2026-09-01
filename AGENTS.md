# Ploma Vermella — Agent Instructions

CLI tool that reviews Google Doc book chapters and posts feedback as comments.

> This is the single source of truth for agent instructions in this repo. `CLAUDE.md`
> is a thin pointer to this file so that Claude Code (and any other AGENTS.md-aware
> tool) loads the same content.

## Development workflow

- **Trunk-based development.** Commit directly to `main` and keep history linear — no
  pull requests, no feature branches, no merge commits. The end state should read as a
  single sequential history.
- **One change at a time.** Make a change, run the tests, commit it, push it, then start
  the next. Don't batch unrelated changes into one commit.
- The pre-commit hook runs `ruff check pv.py tests/` and `pytest tests/`; keep both green.
  Tests cover pure functions and must use fake IDs/data only.
- Never commit secrets or private content (see *What Not to Commit*).

## Files

- `pv.py` — Google Docs/Drive API logic and `pv` CLI entry point
- `tests/test_pv.py` — unit tests (run with `.venv/bin/pytest`, lint with `.venv/bin/ruff`)
- `install-hooks.sh` — installs pre-commit hook that runs tests before each commit
- `context/` — gitignored; one subdir per **work** (a book or other long-running effort under review)
  - Each work subdir is its own independent git repo, versioned locally (no shared remote).
  - `context/<work-slug>/README.md` — **index of that work's context. Read this first;** it
    describes the work and points to everything else worth loading (progressive loading).
  - `context/<work-slug>/style_guide.md` — prose rules for that work. Each work is
    self-contained, so style guides may diverge between works.
  - Other per-work files (`outline.md`, `documents.md`, `keystone-chapters.md`, `folders.md`,
    `notes.md`, `figures.md`, `talking-chapters/`, etc.) are listed and described in the
    work's `README.md`. A work's `notes.md` records where its conventions diverge from its
    own style guide (e.g. a multi-author book that uses "we" intentionally).
- `references/` — gitignored long-lived local reference material such as publisher style guides
- `credentials/` — gitignored; contains OAuth credentials and token

## Working durably with a work's context

Conversation state is summarized away at compaction; only the repo survives. So:

- **Don't assume a chapter's status.** Whether a chapter is drafted or a stub is recorded in
  the work's `documents.md` (and is derivable from its draft/manuscript folder). Check it —
  don't guess from memory, and don't drop planning bullets into a drafted chapter.
- **Persist stabilized models.** When a conceptual model or decision firms up in discussion,
  write it into the work's context repo before moving on, so it outlives the conversation.

## CLI (`pv`)

Invoke via `.venv/bin/pv` — no need to activate the virtualenv first.

```
.venv/bin/pv list <folder-url>                         # list docs in a Drive folder
.venv/bin/pv fetch <doc-url>                           # fetch title + text of a doc
.venv/bin/pv slides-fetch <presentation-url>           # fetch slide text from a deck
.venv/bin/pv slides-thumb <presentation-url> <page-id> # get a slide thumbnail URL
.venv/bin/pv sheet-fetch <sheet-url> [--range A1...]   # read sheet metadata or rows
.venv/bin/pv sheet-update <sheet-url> <range> <json>   # write rows to a sheet range
.venv/bin/pv figure-map <doc-url>                      # list image neighborhoods in a doc
.venv/bin/pv replace-block <doc-url> <start> <end> ... # replace an inclusive body-element range
                                                       # (the replacement keeps the replaced block's paragraph style)
.venv/bin/pv insert-image <doc-url> <body-index> ...   # insert an inline image
.venv/bin/pv replace-image <doc-url> <caption> <deck-url> <slide-id>  # re-export a figure from a slide thumbnail
.venv/bin/pv place-figure <doc-url> <anchor> <deck-url> <slide-id> --caption ...  # insert a centered figure + caption after an anchor
.venv/bin/pv note <doc-url> <quoted-text> <comment>    # append to the review section
.venv/bin/pv clear <doc-url>                           # remove the review section
.venv/bin/pv mv <doc-url> <folder-url>                 # move a doc into a folder
.venv/bin/pv cp <doc-url> <folder-url> [--name NAME]   # copy a doc into a folder
.venv/bin/pv upload <local-path> <folder-url> [--name NAME]  # upload a local file (e.g. a built EPUB/PDF) into a folder
.venv/bin/pv review-copy <doc-url> <folder-url>        # copy with dated title, clear review section
.venv/bin/pv edit <doc-url> <old> <new> [--all]        # replace text in the doc body
.venv/bin/pv replace-section <doc-url> <heading> <text>  # replace a heading's body up to the next heading
.venv/bin/pv find <doc-url> <text>                     # locate text: indices, style, is_code, context
.venv/bin/pv outline <doc-url> [--full]                # structural map: headings + images (indices, object IDs)
.venv/bin/pv words <doc-or-folder-url> [--exclude SUBSTRING ...]  # word count for a doc, or every doc
                                                       # in a folder, with a total; --exclude skips by name and reports what it skipped
.venv/bin/pv shade <doc-url> <start> <end> [--color HEX] [--all] [--remove]  # shade a delimited
                                                       # block's paragraphs (e.g. <sidebar>…</sidebar>); --all does every pair in the doc
.venv/bin/pv prose-check <doc-url> [--terms PATH] [--chapter CH] [--phrases PATH]  # mechanical style sweep: sentence length, em-dash density,
                                                       # passives, tics, UK forms, -ly ordinals, paragraph length, first-use italics, stacked headings, figure refs.
                                                       # --chapter scopes the italics rule per-book: a term introduced in another
                                                       # chapter should be plain here, not italicized again
.venv/bin/pv suggestions <doc-url>                     # list an editor's suggested edits as per-paragraph before→after deltas
                                                       # reports text AND style-only suggestions; total_suggestion_count is
                                                       # distinct IDs (matches the Docs UI), not insertions + deletions
.venv/bin/pv insert-after <doc-url> <anchor> <text>    # insert paragraph(s) after an anchor paragraph
.venv/bin/pv insert-before <doc-url> <anchor> <text>   # insert paragraph(s) before an anchor paragraph
.venv/bin/pv link <doc-url> <text> <url> [--all] [--color HEX|--no-color]  # hyperlink a span of text
.venv/bin/pv cite <doc-url> <title> <url> [--all] [--color HEX|--no-color] # italicize + hyperlink a work title
                                                       # links are set in the O'Reilly red (d3002d) by default,
                                                       # not the default blue; --no-color opts out
.venv/bin/pv style <doc-url> <text> [--italic|--bold|--underline|--no-italic|--no-bold|--no-underline|--color HEX] [--all]  # character styling (turn styles on, or off with --no-*)
.venv/bin/pv heading <doc-url> <anchor> <level>        # set a paragraph's style (1-6, normal, title) by anchor
.venv/bin/pv bullets <doc-url> <start> [end] [--ordered] [--remove]  # make a paragraph range a bulleted/numbered list (or strip bullets)
.venv/bin/pv comment <doc-url> <quoted-text> <text>    # anchored sidebar comment
.venv/bin/pv comments <doc-url> [--include-resolved]   # list comments: id, content, quoted text, resolved
.venv/bin/pv resolve <doc-url> <comment-id>            # resolve a comment
.venv/bin/pv resolve-all <doc-url>                     # resolve every unresolved comment
.venv/bin/pv build-epub <doc-url> ... [-o OUT] [--title T] [--subtitle S] [--author A] [--cover IMG] [--max-image-width N] [--no-optimize] [--part "TITLE=START_DOC_URL" ...]
                                                       # build an EPUB; book metadata comes from context, not hardcoded; images downscaled to 1600px by default; --part groups chapters under a Part heading in the Contents/nav
.venv/bin/pv build-pdf <doc-url> ... [-o OUT] [--title T] [--subtitle S] [--author A] [--cover IMG] [--paper-size letter|a4|...] [--keep-epub] [--part "TITLE=START_DOC_URL" ...]
                                                       # build a paginated PDF via the same pipeline as build-epub, rendered with Calibre's ebook-convert
```

All commands output JSON. Use `.venv/bin/pv -h` or `.venv/bin/pv <command> -h` for help.

`pv list` only returns Google **Docs** (it filters on `mimeType='application/vnd.google-apps.document'`)
— it silently returns an empty list for a folder that holds only Slides decks, Sheets, or other file
types. Don't read an empty `pv list` result as "this folder is empty"; if you need every item in a
folder regardless of type, query the Drive API directly instead.

Text matching (`edit`, `find`, `link`, `style`, `insert-after`) is **quote-agnostic**: curly and
straight quotes/apostrophes match interchangeably, so you don't have to reproduce smart quotes exactly.

`edit` and `find` also match text **inside table cells** (they walk the body in reading order,
descending into tables). Editing a table cell's text works like any other edit; inserting or
deleting whole table rows/columns is still out of scope.

**Extend `pv` when the work turns manual.** If a task has you repeatedly dropping to the raw
Docs/Slides/Sheets API (ad-hoc `pv._docs_service()` Python) or hand-running the same multi-step edit,
treat that as a signal that `pv` is missing a primitive — each manual workaround is a command that
should exist. Prefer adding or updating a `pv` command (a pure, testable plan function plus a thin
wrapper, following the one-change-at-a-time discipline above, and documented here and in `README.md`)
over accumulating one-off scripts. Extend the existing text-anchor + structured "ambiguous result"
model (see `edit`, `outline`, `heading`, `replace-section`) rather than inventing a new shape.

## Review Workflow

Run `pv prose-check <doc-url> --terms context/<work>/terms.txt` first — it reports the
mechanical style rules (sentence length, em-dash density, passives, tic phrases, UK forms,
first-use italics, stacked headings, unreferenced figures) with the work's target numbers, so
those never depend on a reviewer noticing them. It deliberately does *not* judge citations,
verbatim quotes, metaphors, or chapter openings; those are listed under `needs_a_reader` in its
output and are your job. A work supplies its own `terms.txt` (its terms of art) and the targets
live in its `style_guide.md`.

When asked to review a chapter:
1. Identify which **work** the chapter belongs to (one subdir under `context/`). Read
   `context/<work-slug>/README.md` first — it indexes the work and points to `folders.md`
   (the Drive folder URL, if present), the style guide, the outline, and so on.
2. Run `pv list <folder-url>` to list available documents, or use a doc URL/ID directly if given
3. Run `pv fetch <doc-url>` to get the chapter text and any existing comments
4. Read the work's `style_guide.md` plus the other files its `README.md` points to, and any
   relevant files in `references/`, via the Read tool. The work's `notes.md` flags any
   divergence from its style guide.
5. Run `pv clear <doc-url>` to remove any previous review section
6. For each issue found, run `pv note <doc-url> <quoted-text> <comment>` — `quoted_text` must be an exact substring of the document text. The `comment` should be self-contained: when the problem is a specific word or phrase, quote it inside the comment (e.g. `"Draft placeholder: \"Something something\" — expand or cut"`) so the reader knows exactly what to fix without needing to hunt for the highlighted text.

## LLM-speak

Agent-assisted drafting is the likeliest way machine-register prose enters a manuscript, and
the agent writing it is the least likely to notice. So:

- **Check your own prose before proposing it.** Any drafted sentence, definition, sidebar or
  rewrite gets checked against the work's phrase list (`context/<work>/llm-speak.txt`) *before*
  it reaches the user or the document. This applies to prose written into the work's context
  repo too, not just into chapters.
- **Correct it during editing.** Run `pv prose-check <doc> --phrases context/<work>/llm-speak.txt`
  as part of any review or sweep, and treat hits as mechanical fixes.
- **Add to the list when the user names a new one**, and note in the file whether it was
  measured in the manuscript or is there to prevent ingress.
- **Never edit the author's own voice on this basis — but check provenance properly.**
  `git log -S` is **useless** for this: the agent commits as the user, so everything in a
  work's context repo reads as the author's regardless of who wrote it. Much of that repo is
  agent-written, so presence there is not evidence of anything.
  The reliable test is the work's **`talking-chapters/` transcripts**, which are unambiguously
  the author speaking:
  - Present in the transcripts → the author's word. Leave it alone.
  - Absent from every transcript, but present in agent-written docs → very likely the agent's.
    Say so, and check with the author before adding it to the list.
  A phrase inside a quotation is always the source's, whoever typed it.
- **Keep the list conservative.** Measure before adding: a phrase with legitimate use at volume
  (`leverage` as a noun, `carries`, `land`, `sits`) makes the check noisy and then ignored.
  Record rejections with their counts so the list is not re-padded on intuition later.

## Acting on Review Feedback

When working through review notes (PV bullets or sidebar comments), classify each into one of three tiers and act accordingly:

- **Mechanical fix** (typo, US-vs-UK spelling, missing comma, hyphenation, term-rename, unambiguous text substitution): batch with sibling mechanical fixes and apply via `pv edit`. Show the user a single report listing every fix that landed (or skipped, with reason). Don't ask per-fix permission.
- **Judgement fix** (single text edit that requires per-instance judgement — sentence rewrite, citation reformat, voice recast): propose one at a time as a `before → after` diff in chat. Apply only after the user says so. Do not batch.
- **Discussion fix** (structural, ambiguous, or open-ended — section reorganisation, "consider splitting", "feels long", or notes the reviewer flagged as questions): leave alone unless the user explicitly asks. If a discussion fix can be reframed as a mechanical or judgement fix, propose the conversion; otherwise it stays as an open comment.

## Figure Editing Workflow

When fixing figures, avoid broad Google Docs `batchUpdate` edits over multiple figures. Those are fragile because document indices shift after each insertion or deletion.

Use this workflow instead:
1. Run `pv figure-map <doc-url>` and work from the reported `body_index` neighborhoods rather than cached raw indices.
2. Inspect one figure at a time: lead-in paragraph, image paragraph, caption paragraph, following paragraph.
3. Use `pv replace-block` for local prose/caption cleanup around that single figure block.
4. Use `pv slides-fetch` and `pv slides-thumb` to match the source slide and retrieve a reinsertion URL when needed.
5. Use `pv insert-image` only when the image itself is missing.
6. Re-run `pv figure-map` after each figure to verify the local block before moving on.

For work-specific resources (deck URLs, figure logs, publisher-facing tracking sheets, intake conventions for third-party images), check `context/<work-slug>/figures.md` if present (the work's `README.md` will say).

## What Not to Commit

Never commit:
- Anything under `credentials/` (gitignored — OAuth secrets, tokens, URLs)
- Anything under `references/` unless you are certain you have rights to redistribute it
- Anything under `context/` to *this* repo (gitignored — book content lives in each work's own repo)
- Absolute paths specific to this machine (e.g. `/Users/yourname/...`)
- Document IDs, folder IDs, or deployment URLs — these belong in `credentials/` or context files
- API keys or client secrets of any kind

## Known Limitations

Google Docs does not expose text-anchored comment creation via any public API (Drive API, Docs API, or Apps Script). Comments post to the sidebar with `quotedFileContent` visible but without yellow text highlighting. This is a deliberate Google restriction.

## Setup

```bash
python3 -m venv .venv && source .venv/bin/activate && pip install -e ".[dev]" && sh install-hooks.sh
```

Credentials: Google Cloud project with Docs API, Drive API, Slides API, and Sheets API enabled, OAuth 2.0 Desktop credentials saved as `credentials/client_secret.json`. First run opens a browser for authorisation; token cached at `credentials/token.json`.

`pv build-pdf` shells out to Calibre's `ebook-convert` CLI (not a Python dependency) to
render the intermediate EPUB to PDF: `brew install --cask calibre`. `pv build-epub` doesn't
need it.

## OAuth Reauth

If `pv` fails with `google.auth.exceptions.RefreshError` and `invalid_grant: Token has been expired or revoked.`, the token refresh path will not recover automatically. Trigger a fresh OAuth flow by moving `credentials/token.json` aside and rerunning a `pv` command, for example:

```bash
mv credentials/token.json credentials/token.json.bak-$(date +%Y%m%d) && .venv/bin/pv list <folder-url>
```

`pv` will print a Google authorisation URL and start a local callback server. Complete the browser login, let the command finish, and a fresh `credentials/token.json` will be written. Do not delete `client_secret.json`.
