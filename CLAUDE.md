# Ploma Vermella — Claude Code Instructions

CLI tool that reviews Google Doc book chapters and posts feedback as comments.

## Files

- `pv.py` — Google Docs/Drive API logic and `pv` CLI entry point
- `tests/test_pv.py` — unit tests (run with `pytest`, lint with `ruff`)
- `install-hooks.sh` — installs pre-commit hook that runs tests before each commit
- `context/style_guide.md` — working prose rules for the current review
- `context/outline.md` — working chapter outline for the current review
- `references/` — gitignored long-lived local reference material such as publisher style guides
- `credentials/` — gitignored; contains OAuth credentials and token

## CLI (`pv`)

Invoke via `.venv/bin/pv` — no need to activate the virtualenv first.

```
.venv/bin/pv list <folder-url>                         # list docs in a Drive folder
.venv/bin/pv fetch <doc-url>                           # fetch title + text of a doc
.venv/bin/pv note <doc-url> <quoted-text> <comment>    # append to the review section
.venv/bin/pv clear <doc-url>                           # remove the review section
```

All commands output JSON. Use `.venv/bin/pv -h` or `.venv/bin/pv <command> -h` for help.

## Review Workflow

When asked to review a chapter:
1. Read `context/folders.md` to find the Drive folder URL (gitignored, user-maintained)
2. Run `pv list <folder-url>` to list available documents, or use a doc URL/ID directly if given
3. Run `pv fetch <doc-url>` to get the chapter text and any existing comments
4. Read all files in `context/` and `references/` via the Read tool
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
