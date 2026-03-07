"""MCP server exposing Google Docs review tools to Claude Code."""

from fastmcp import FastMCP
import gdocs

mcp = FastMCP(
    name="gdocs",
    instructions=(
        "Tools for fetching Google Docs content and managing comments. "
        "When reviewing a chapter: fetch the document, fetch existing comments to avoid duplicates, "
        "read context/style_guide.md and context/outline.md via the Read tool, "
        "then post comments for each issue found — skipping passages already commented on."
    ),
)


@mcp.tool(
    description=(
        "List all Google Docs in a Google Drive folder. "
        "Accepts either a full Drive folder URL or a bare folder ID. "
        "Returns a list of {id, name, url} — one entry per document."
    )
)
def gdocs_list_folder(folder_id_or_url: str) -> list[dict]:
    """List all Google Docs in a Drive folder."""
    return gdocs.list_folder(folder_id_or_url)


@mcp.tool(
    description=(
        "Fetch the full plain-text content of a Google Doc. "
        "Accepts either a full Google Docs URL or a bare document ID. "
        "Returns {title: str, text: str}."
    )
)
def gdocs_fetch_document(doc_id_or_url: str) -> dict:
    """Fetch title and plain text of a Google Doc."""
    return gdocs.fetch_document(doc_id_or_url)


@mcp.tool(
    description=(
        "Fetch all existing comments on a Google Doc. "
        "Accepts either a full Google Docs URL or a bare document ID. "
        "Returns a list of {id, author, content, quoted_text}. "
        "Use this before posting to avoid duplicate comments on the same passage."
    )
)
def gdocs_fetch_comments(doc_id_or_url: str) -> list[dict]:
    """List existing comments on a Google Doc."""
    return gdocs.fetch_comments(doc_id_or_url)


@mcp.tool(
    description=(
        "Post a review comment anchored to a specific passage in a Google Doc. "
        "Accepts either a full Google Docs URL or a bare document ID. "
        "quoted_text must be an exact substring of the document text. "
        "Performs a duplicate check: if the same quoted_text already has a comment, returns a no-op. "
        "Returns {status: 'posted'|'skipped', ...}."
    )
)
def gdocs_post_comment(doc_id_or_url: str, quoted_text: str, comment: str) -> dict:
    """Post a comment anchored to quoted_text; skips if already commented."""
    return gdocs.post_comment(doc_id_or_url, quoted_text, comment)


if __name__ == "__main__":
    mcp.run()
