# Ploma Vermella — Claude Code Instructions

CLI tool that reviews Google Doc book chapters and posts feedback as comments.

## Files

- `gdocs.py` — Google Docs/Drive API logic; also the `pv` CLI entry point
- `tests/test_gdocs.py` — unit tests (run with `pytest`, lint with `ruff`)
- `install-hooks.sh` — installs pre-commit hook that runs tests before each commit
- `context/style_guide.md` — prose rules (edit to update review criteria)
- `context/outline.md` — chapter outline (edit to update structure expectations)
- `credentials/` — gitignored; contains OAuth credentials and token

## CLI (`pv`)

```
pv list <folder-url>                          # list docs in a Drive folder
pv fetch <doc-url>                            # fetch title + text of a doc
pv comments <doc-url>                         # list existing comments
pv comment <doc-url> <quoted-text> <comment>  # post a sidebar comment
pv note <doc-url> <quoted-text> <comment>     # append to the review section
```

All commands output JSON. Use `pv -h` or `pv <command> -h` for help.

## Review Workflow

When asked to review a chapter:
1. Read `context/folders.md` to find the Drive folder URL (gitignored, user-maintained)
2. Run `pv list <folder-url>` to list available documents, or use a doc URL/ID directly if given
3. Run `pv fetch <doc-url>` to get the chapter text
4. Run `pv comments <doc-url>` to see existing comments — skip any `quoted_text` already commented on
5. Read `context/style_guide.md` and `context/outline.md` via the Read tool
6. For each issue found, run `pv comment <doc-url> <quoted-text> <comment>` — `quoted_text` must be an exact substring of the document text
7. Run `pv note` once per issue to also append it to the in-document review section

## What Not to Commit

Never commit:
- Anything under `credentials/` (gitignored — OAuth secrets, tokens, URLs)
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
