"""Read-only DOCX inspection tools, and an MCP server that exposes them."""
from .docx_tools import (
    check_term_variants,
    find_term,
    get_docx_stats,
    list_paragraph_range,
    read_accepted_text,
    scan_meta_notes,
)

__version__ = "0.1.0"
__all__ = [
    "__version__",
    "check_term_variants",
    "find_term",
    "get_docx_stats",
    "list_paragraph_range",
    "read_accepted_text",
    "scan_meta_notes",
]
