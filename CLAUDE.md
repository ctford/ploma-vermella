# Ploma Vermella — Claude Code Instructions

CLI tool that reviews Google Doc book chapters and posts feedback as comments.

## Files

- `pv.py` — Google Docs/Drive API logic and `pv` CLI entry point
- `tests/test_pv.py` — unit tests (run with `.venv/bin/pytest`, lint with `.venv/bin/ruff`)
- `install-hooks.sh` — installs pre-commit hook that runs tests before each commit
- `context/` — gitignored; one subdir per book under review, plus shared rules
  - `context/style_guide.md` — shared prose rules (applies to every book)
  - `context/<book-slug>/` — per-book context: `outline.md`, `documents.md`, `keystone-chapters.md`, `folders.md`, `notes.md`, `talking-chapters/`, etc.
  - A book's `notes.md` records where its conventions diverge from the shared style guide (e.g. a multi-author book that uses "we" intentionally)
- `references/` — gitignored long-lived local reference material such as publisher style guides
- `credentials/` — gitignored; contains OAuth credentials and token

## CLI (`pv`)

Invoke via `.venv/bin/pv` — no need to activate the virtualenv first.

```
.venv/bin/pv list <folder-url>                         # list docs in a Drive folder
.venv/bin/pv fetch <doc-url>                           # fetch title + text of a doc
.venv/bin/pv note <doc-url> <quoted-text> <comment>    # append to the review section
.venv/bin/pv clear <doc-url>                           # remove the review section
.venv/bin/pv mv <doc-url> <folder-url>                 # move a doc into a folder
.venv/bin/pv cp <doc-url> <folder-url> [--name NAME]   # copy a doc into a folder
.venv/bin/pv review-copy <doc-url> <folder-url>        # copy with dated title, clear review section
.venv/bin/pv edit <doc-url> <old> <new> [--all]        # replace text in the doc body
.venv/bin/pv resolve <doc-url> <comment-id>            # resolve a comment
```

All commands output JSON. Use `.venv/bin/pv -h` or `.venv/bin/pv <command> -h` for help.

## Review Workflow

When asked to review a chapter:
1. Identify which book the chapter belongs to (one subdir under `context/`). Read `context/<book-slug>/folders.md` for the Drive folder URL if present.
2. Run `pv list <folder-url>` to list available documents, or use a doc URL/ID directly if given
3. Run `pv fetch <doc-url>` to get the chapter text and any existing comments
4. Read `context/style_guide.md` (shared rules) plus everything under `context/<book-slug>/` and any relevant files in `references/` via the Read tool. The book's `notes.md` will flag any divergence from the shared style guide.
5. Run `pv clear <doc-url>` to remove any previous review section
6. For each issue found, run `pv note <doc-url> <quoted-text> <comment>` — `quoted_text` must be an exact substring of the document text. The `comment` should be self-contained: when the problem is a specific word or phrase, quote it inside the comment (e.g. `"Draft placeholder: \"Something something\" — expand or cut"`) so the reader knows exactly what to fix without needing to hunt for the highlighted text.

## What Not to Commit

Never commit:
- Anything under `credentials/` (gitignored — OAuth secrets, tokens, URLs)
- Anything under `references/` unless you are certain you have rights to redistribute it
- Absolute paths specific to this machine (e.g. `/Users/yourname/...`)
- Document IDs, folder IDs, or deployment URLs — these belong in `credentials/` or context files
- API keys or client secrets of any kind

## Known Limitations

Google Docs does not expose text-anchored comment creation via any public API (Drive API, Docs API, or Apps Script). Comments post to the sidebar with `quotedFileContent` visible but without yellow text highlighting. This is a deliberate Google restriction.

## Setup

```bash
python3 -m venv .venv && source .venv/bin/activate && pip install -e ".[dev]" && sh install-hooks.sh
```

Credentials: Google Cloud project with Docs API + Drive API enabled, OAuth 2.0 Desktop credentials saved as `credentials/client_secret.json`. First run opens a browser for authorisation; token cached at `credentials/token.json`.

## OAuth Reauth

If `pv` fails with `google.auth.exceptions.RefreshError` and `invalid_grant: Token has been expired or revoked.`, the token refresh path will not recover automatically. Trigger a fresh OAuth flow by moving `credentials/token.json` aside and rerunning a `pv` command, for example:

```bash
mv credentials/token.json credentials/token.json.bak-$(date +%Y%m%d) && .venv/bin/pv list <folder-url>
```

`pv` will print a Google authorisation URL and start a local callback server. Complete the browser login, let the command finish, and a fresh `credentials/token.json` will be written. Do not delete `client_secret.json`.
