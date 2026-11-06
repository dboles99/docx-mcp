"""The MCP server: six read-only DOCX tools over stdio.

Run it with ``docx-mcp`` or ``python -m docx_mcp.server``. No tool here writes to or
changes a DOCX file.
"""
from __future__ import annotations

import re
import zipfile
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

from lxml import etree
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

from .docx_tools import (
    check_term_variants,
    find_term,
    get_docx_stats,
    list_paragraph_range,
    read_accepted_text,
    scan_meta_notes,
)

mcp = MCPServer("docx-mcp")


@contextmanager
def _reported_errors() -> Iterator[None]:
    """Turn an expected failure into a tool error the agent can read.

    The SDK withholds the text of any other exception from the client, which would
    leave an agent unable to tell a wrong path from a damaged file.
    """
    try:
        yield
    except (OSError, ValueError, zipfile.BadZipFile, re.error, etree.XMLSyntaxError) as exc:
        raise ToolError(f"{type(exc).__name__}: {exc}") from exc


@mcp.tool()
def read_docx_accepted_text(
    path: str,
    start: int = 0,
    limit: int = 50,
    min_length: int = 10,
) -> list[dict[str, Any]]:
    """Read paragraphs from a DOCX as accepted text, with tracked deletions removed.

    Returns a list of {para_index, text, has_tracked_change, char_count}.
    Pagination: start=0, limit=50 examines paragraphs 0 to 49. Set min_length=0 to
    include short and empty paragraphs.
    """
    with _reported_errors():
        return read_accepted_text(path, start=start, limit=limit, min_length=min_length)


@mcp.tool()
def find_docx_term(
    path: str,
    pattern: str,
    context_chars: int = 150,
    flags: str = "IGNORECASE",
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Search the accepted text of a DOCX for a regular expression.

    pattern is a Python regex. flags is a comma-separated list of re flag names, such
    as "IGNORECASE,DOTALL". Returns a list of
    {para_index, match, context, has_tracked_change}.
    """
    with _reported_errors():
        return find_term(path, pattern, context_chars=context_chars, flags=flags, limit=limit)


@mcp.tool()
def check_docx_term_variants(
    path: str,
    terms: list[str],
    case_sensitive: bool = False,
) -> dict[str, Any]:
    """Count occurrences of each term across all accepted text in a DOCX.

    Useful for consistency checks, for example passing
    ["cyber security", "cybersecurity"] to see which spelling is used and where.
    """
    with _reported_errors():
        return check_term_variants(path, terms, case_sensitive=case_sensitive)


@mcp.tool()
def get_docx_structure_stats(path: str) -> dict[str, Any]:
    """Return quick statistics for a DOCX: paragraph and word counts, tracked-change
    paragraph count, comment count and file size.
    """
    with _reported_errors():
        return get_docx_stats(path)


@mcp.tool()
def scan_docx_meta_notes(
    path: str,
    patterns: list[str] | None = None,
    check_first_n: int = 20,
) -> list[dict[str, Any]]:
    """Scan the first N paragraphs of a DOCX for working notes left in by mistake,
    such as "TODO:" or "[internal]". Pass patterns to replace the defaults.
    Returns a list of {para_index, text, matched_pattern}.
    """
    with _reported_errors():
        return scan_meta_notes(path, patterns=patterns, check_first_n=check_first_n)


@mcp.tool()
def list_docx_paragraph_range(
    path: str,
    start: int,
    end: int,
    include_empty: bool = False,
) -> list[dict[str, Any]]:
    """Return accepted text for paragraphs [start, end) by index, to inspect a known
    region of a document.
    """
    with _reported_errors():
        return list_paragraph_range(path, start, end, include_empty=include_empty)


def run() -> None:
    """Entry point for the ``docx-mcp`` console script (stdio transport)."""
    mcp.run()


if __name__ == "__main__":
    run()
