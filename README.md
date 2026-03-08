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
3. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**.
   - Application type: **Desktop app**
   - Download the JSON and save it as `credentials/client_secret.json`.
4. Under **OAuth consent screen**, add your Google account as a test user.

On first use, a browser window will open for OAuth authorisation. The token is cached at `credentials/token.json`.

---

## Usage

Tell Claude Code in plain English:

> "Review chapter 3 at `<doc-url>`."

Claude will read all files in `context/`, fetch the document, and append a **🪶 Ploma Vermella Review** section at the end of the doc with dated, bulleted, located feedback.

The `pv` CLI is available for direct use:

```bash
pv list <folder-url>                        # list docs in a Drive folder
pv fetch <doc-url>                          # fetch title + text of a doc
pv note <doc-url> <quoted-text> <comment>   # append a review note
```

---

## Context

Drop any reference material into `context/` — style guide, chapter outline, author notes. Claude reads everything in the directory on each review. No code changes needed.

---

## Development

```bash
pytest tests/ -v        # run tests
ruff check gdocs.py tests/  # lint
```

Tests and lint run automatically as a pre-commit hook and on every push via GitHub Actions.
