# Ploma Vermella — Claude Code Instructions

MCP server that reviews Google Doc book chapters and posts feedback as comments.

## Files

- `server.py` — FastMCP server, exposes 4 tools
- `gdocs.py` — Google Docs/Drive API logic
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
1. List the folder to find the document
2. Fetch the document text
3. Fetch existing comments to avoid duplicates
4. Read `context/style_guide.md` and `context/outline.md`
5. Post one comment per issue — `quoted_text` must be an exact substring of the document

## Known Limitations

Google Docs does not expose text-anchored comment creation via any public API (Drive API, Docs API, or Apps Script). Comments post to the sidebar with `quotedFileContent` visible but without yellow text highlighting. This is a deliberate Google restriction.

## Setup

```bash
python3 -m venv .venv && source .venv/bin/activate && pip install -e .
```

Credentials: Google Cloud project with Docs API + Drive API enabled, OAuth 2.0 Desktop credentials saved as `credentials/client_secret.json`. First run opens a browser for authorisation; token cached at `credentials/token.json`.

Register with Claude Code:
```bash
claude mcp add gdocs -- python $(pwd)/server.py
```
