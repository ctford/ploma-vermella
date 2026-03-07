# Ploma Vermella — Claude Code Instructions

MCP server that reviews Google Doc book chapters and posts feedback as comments.

## Files

- `server.py` — FastMCP server, exposes 4 tools
- `gdocs.py` — Google Docs/Drive API logic
- `tests/test_gdocs.py` — unit tests (run with `pytest`, lint with `ruff`)
- `install-hooks.sh` — installs pre-commit hook that runs tests before each commit
- `context/style_guide.md` — prose rules (edit to update review criteria)
- `context/outline.md` — chapter outline (edit to update structure expectations)
- `credentials/` — gitignored; contains OAuth credentials and token

## MCP Tools

- `gdocs_list_folder(folder_id_or_url)` → `[{id, name, url}]`
- `gdocs_fetch_document(doc_id_or_url)` → `{title, text}`
- `gdocs_fetch_comments(doc_id_or_url)` → `[{id, author, content, quoted_text}]`
- `gdocs_post_comment(doc_id_or_url, quoted_text, comment)` → `{status, ...}`

## Review Workflow

When asked to review a chapter:
1. Read `context/folders.md` to find the Drive folder URL (gitignored, user-maintained)
2. Call `gdocs_list_folder` to list available documents, or use a doc URL/ID directly if given
3. Call `gdocs_fetch_document` to get the chapter text
4. Call `gdocs_fetch_comments` to see existing comments — skip any `quoted_text` already commented on
5. Read `context/style_guide.md` and `context/outline.md` via the Read tool
6. Post one comment per issue via `gdocs_post_comment` — `quoted_text` must be an exact substring of the document text

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

Register with Claude Code:
```bash
claude mcp add gdocs -- python $(pwd)/server.py
```
