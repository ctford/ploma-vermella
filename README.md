# ploma-vermella

MCP server for reviewing Google Doc book chapters against a style guide and outline, posting feedback as inline comments.

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
3. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**.
   - Application type: **Desktop app**
   - Download the JSON and save it as:
     ```
     credentials/client_secret.json
     ```
4. Under **OAuth consent screen**, add your Google account as a test user.

On first use, a browser window will open for OAuth authorization. The token is cached at `credentials/token.json`.

### 3. Register MCP server in Claude Code

```bash
claude mcp add gdocs -- python $(pwd)/server.py
```

Or manually add to `~/.claude/claude.json`:

```json
{
  "mcpServers": {
    "gdocs": {
      "command": "python",
      "args": ["/path/to/ploma-vermella/server.py"]
    }
  }
}
```

Restart Claude Code after registering.

---

## Verification

```
> List the tools available from gdocs
```
Should show: `gdocs_fetch_document`, `gdocs_fetch_comments`, `gdocs_post_comment`.

```
> Fetch the content of <doc-url>
```
Should return the chapter title and text.

```
> Review <doc-url> using context/style_guide.md and context/outline.md — post comments, skip anything already commented
```
Comments appear on the Google Doc. Running again posts no duplicates.

---

## Usage

Tell Claude Code in plain English:

> "Review chapter 3 at `<doc-url>` against our style guide. Post inline comments for each issue. Skip anything that already has a comment."

Claude will:
1. Call `gdocs_fetch_document` to get the chapter text
2. Call `gdocs_fetch_comments` to see existing comments
3. Read `context/style_guide.md` and `context/outline.md` via its `Read` tool
4. Review the chapter and call `gdocs_post_comment` for each issue found

---

## Editing the style guide and outline

Just edit `context/style_guide.md` and `context/outline.md` directly — no code changes needed. Claude reads them fresh on every review.

---

## Development

```bash
pytest tests/ -v        # run tests
ruff check *.py tests/  # lint
```

Tests and lint also run automatically as a pre-commit hook (installed via `install-hooks.sh`) and on every push via GitHub Actions.

## Project structure

```
ploma-vermella/
├── server.py            # MCP server (FastMCP)
├── gdocs.py             # Google Docs / Drive API logic
├── tests/
│   └── test_gdocs.py    # unit tests
├── context/
│   ├── style_guide.md   # Prose rules and O'Reilly conventions
│   └── outline.md       # Chapter-by-chapter outline
├── credentials/         # gitignored — OAuth credentials
│   └── client_secret.json
├── install-hooks.sh     # installs pre-commit hook
├── .github/workflows/ci.yml
├── pyproject.toml
└── README.md
```
