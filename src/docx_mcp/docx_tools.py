"""Read-only inspection of DOCX files, using zipfile and lxml directly.

Every function here reads ``word/document.xml``, or ``word/comments.xml`` for the
comment count, and never modifies the file.

What "accepted text" means here: what the document would say if every tracked change
were accepted. Insertions count. Deletions and moved-away text do not.

Pitfalls this module handles:

- Split runs. Word divides a sentence across runs for arbitrary reasons, so text is
  joined per paragraph before any matching.
- Tracked deletions. Text inside ``w:del`` and ``w:moveFrom`` is excluded at any depth.
- Text boxes. A paragraph inside a text box is its own paragraph, not part of the
  paragraph anchoring it, and the ``mc:Fallback`` copy Word writes is not counted again.
- lxml element truthiness. Always test ``is not None``, never truthiness.
- Untrusted input. The XML parser does not resolve entities, load a DTD or use the
  network, and an oversized part is refused before it is decompressed.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

_DOCUMENT_PART = "word/document.xml"
_COMMENTS_PART = "word/comments.xml"

_DEL_TAG = f"{{{W}}}del"
_INS_TAG = f"{{{W}}}ins"
_MOVE_FROM_TAG = f"{{{W}}}moveFrom"
_MOVE_TO_TAG = f"{{{W}}}moveTo"
_T_TAG = f"{{{W}}}t"
_P_TAG = f"{{{W}}}p"
_PPR_TAG = f"{{{W}}}pPr"
_TAB_TAG = f"{{{W}}}tab"
_BR_TAG = f"{{{W}}}br"
_CR_TAG = f"{{{W}}}cr"
_COMMENT_TAG = f"{{{W}}}comment"
_FALLBACK_TAG = f"{{{MC}}}Fallback"

_TRACKED_TAGS = (_DEL_TAG, _INS_TAG, _MOVE_FROM_TAG, _MOVE_TO_TAG)
_NOT_ACCEPTED_TAGS = (_DEL_TAG, _MOVE_FROM_TAG)

# A limit on the decompressed size of one XML part, so that a small crafted file
# cannot exhaust memory.
MAX_PART_BYTES = 256 * 1024 * 1024

# All quote variants, ASCII and Unicode typographic, for use inside a search pattern.
QUOTE_VARIANTS = r"""[\"'"“”‘’„‹›]"""

DEFAULT_NOTE_PATTERNS = (
    r"^\s*Note\s*:",
    r"^\s*TODO\s*[:\-]",
    r"^\s*FIXME\s*[:\-]",
    r"^\s*REVIEW\s*[:\-]",
    r"^\s*\[internal\]",
)


# -- Low-level helpers --------------------------------------------------------


def _parser() -> etree.XMLParser:
    return etree.XMLParser(
        resolve_entities=False, no_network=True, load_dtd=False, huge_tree=False
    )


def _read_part(archive: zipfile.ZipFile, name: str) -> bytes:
    size = archive.getinfo(name).file_size
    if size > MAX_PART_BYTES:
        raise ValueError(f"{name} is {size} bytes, over the {MAX_PART_BYTES} byte limit")
    return archive.read(name)


def _load_root(path: str) -> etree._Element:
    with zipfile.ZipFile(path) as archive:
        if _DOCUMENT_PART not in archive.namelist():
            raise ValueError(f"{path} has no {_DOCUMENT_PART}, so it is not a Word document")
        return etree.fromstring(_read_part(archive, _DOCUMENT_PART), _parser())


def _collect_accepted(element: Any, parts: list[str]) -> None:
    for child in element:
        tag = child.tag
        if tag in _NOT_ACCEPTED_TAGS or tag == _P_TAG or tag == _PPR_TAG:
            continue
        if tag == _T_TAG:
            if child.text:
                parts.append(child.text)
        elif tag == _TAB_TAG:
            parts.append("\t")
        elif tag in (_BR_TAG, _CR_TAG):
            parts.append("\n")
        else:
            _collect_accepted(child, parts)


def _para_accepted_text(para: Any) -> str:
    """The accepted text of one paragraph, excluding any paragraphs nested inside it."""
    parts: list[str] = []
    _collect_accepted(para, parts)
    return "".join(parts)


def _para_has_tracked_change(para: Any) -> bool:
    return any(child.tag in _TRACKED_TAGS for child in para.iter())


def _comment_count(path: str) -> int:
    """Count comments. They live in their own part, not in document.xml."""
    with zipfile.ZipFile(path) as archive:
        if _COMMENTS_PART not in archive.namelist():
            return 0
        root = etree.fromstring(_read_part(archive, _COMMENTS_PART), _parser())
    return sum(1 for _ in root.iter(_COMMENT_TAG))


def _in_fallback(para: Any) -> bool:
    return any(ancestor.tag == _FALLBACK_TAG for ancestor in para.iterancestors())


def _get_body_paragraphs(root: Any) -> list[tuple[int, Any]]:
    """Every paragraph in the body, with a stable index.

    Includes table cells and text boxes. The index is the paragraph's position in
    this list, so it stays usable as an address across paginated calls.
    """
    body = root.find(f"{{{W}}}body")
    if body is None:
        return []
    paragraphs = (el for el in body.iter(_P_TAG) if not _in_fallback(el))
    return list(enumerate(paragraphs))


def _regex_flags(flags: str) -> int:
    known = re.RegexFlag.__members__
    combined = 0
    for name in flags.split(","):
        member = known.get(name.strip().upper())
        if member is not None:
            combined |= member
    return combined


# -- Public tools -------------------------------------------------------------


def read_accepted_text(
    path: str,
    start: int = 0,
    limit: int = 50,
    min_length: int = 10,
) -> list[dict[str, Any]]:
    """Read paragraphs from a DOCX as accepted text.

    Returns a list of {para_index, text, has_tracked_change, char_count}.
    Pagination: start=0, limit=50 examines paragraphs 0 to 49. Set min_length=0 to
    include short and empty paragraphs.
    """
    root = _load_root(path)
    paras = _get_body_paragraphs(root)
    result = []
    for idx, para in paras[start : start + limit]:
        text = _para_accepted_text(para)
        if len(text) < min_length:
            continue
        result.append(
            {
                "para_index": idx,
                "text": text,
                "has_tracked_change": _para_has_tracked_change(para),
                "char_count": len(text),
            }
        )
    return result


def find_term(
    path: str,
    pattern: str,
    context_chars: int = 150,
    flags: str = "IGNORECASE",
    limit: int = 100,
) -> list[dict[str, Any]]:
    """Search the accepted text of a DOCX for a regular expression.

    pattern is a Python regex. Use QUOTE_VARIANTS for a quote-agnostic search.
    flags is a comma-separated list of re flag names, such as "IGNORECASE,DOTALL".
    Unknown names are ignored.
    Returns a list of {para_index, match, context, has_tracked_change}.
    """
    compiled = re.compile(pattern, _regex_flags(flags))
    root = _load_root(path)
    results: list[dict[str, Any]] = []

    for idx, para in _get_body_paragraphs(root):
        text = _para_accepted_text(para)
        for match in compiled.finditer(text):
            start = max(0, match.start() - context_chars // 2)
            end = min(len(text), match.end() + context_chars // 2)
            results.append(
                {
                    "para_index": idx,
                    "match": match.group(0),
                    "context": text[start:end],
                    "has_tracked_change": _para_has_tracked_change(para),
                }
            )
            if len(results) >= limit:
                return results
    return results


def check_term_variants(
    path: str,
    terms: list[str],
    case_sensitive: bool = False,
) -> dict[str, Any]:
    """Count occurrences of each term across all accepted text.

    Useful for consistency checks: pass ['cyber security', 'cybersecurity'] to see
    which forms appear and where.

    Returns {"counts": {...}, "locations": {term: [para_index, ...]}, "total_paragraphs": N}.
    """
    flags = 0 if case_sensitive else re.IGNORECASE
    root = _load_root(path)
    paras = _get_body_paragraphs(root)

    counts: dict[str, int] = {term: 0 for term in terms}
    locations: dict[str, list[int]] = {term: [] for term in terms}

    for idx, para in paras:
        text = _para_accepted_text(para)
        for term in terms:
            hits = len(re.findall(re.escape(term), text, flags))
            if hits:
                counts[term] += hits
                locations[term].append(idx)

    return {
        "counts": counts,
        "locations": locations,
        "total_paragraphs": len(paras),
    }


def get_docx_stats(path: str) -> dict[str, Any]:
    """Return quick structural statistics for a DOCX file.

    Returns {path, paragraph_count, word_count_approx, tracked_change_paragraph_count,
    comment_count, has_tracked_changes, file_size_kb}.
    """
    file_path = Path(path)
    root = _load_root(path)
    paras = _get_body_paragraphs(root)

    word_count = 0
    tracked_paragraphs = 0
    for _, para in paras:
        word_count += len(_para_accepted_text(para).split())
        if _para_has_tracked_change(para):
            tracked_paragraphs += 1

    return {
        "path": str(file_path),
        "paragraph_count": len(paras),
        "word_count_approx": word_count,
        "tracked_change_paragraph_count": tracked_paragraphs,
        "comment_count": _comment_count(path),
        "has_tracked_changes": tracked_paragraphs > 0,
        "file_size_kb": round(file_path.stat().st_size / 1024, 1),
    }


def scan_meta_notes(
    path: str,
    patterns: list[str] | None = None,
    check_first_n: int = 20,
) -> list[dict[str, Any]]:
    """Scan the start of a document for working notes left in by mistake.

    The default patterns match lines starting with Note:, TODO, FIXME, REVIEW or
    [internal]. Pass your own patterns to match project identifiers. They replace
    the defaults.

    Returns a list of {para_index, text, matched_pattern}.
    """
    active = list(DEFAULT_NOTE_PATTERNS) if patterns is None else patterns
    compiled = [re.compile(pattern, re.IGNORECASE) for pattern in active]
    root = _load_root(path)
    results = []

    for idx, para in _get_body_paragraphs(root)[:check_first_n]:
        text = _para_accepted_text(para).strip()
        if not text:
            continue
        for pattern, regex in zip(active, compiled, strict=True):
            if regex.search(text):
                results.append(
                    {"para_index": idx, "text": text[:300], "matched_pattern": pattern}
                )
                break

    return results


def list_paragraph_range(
    path: str,
    start: int,
    end: int,
    include_empty: bool = False,
) -> list[dict[str, Any]]:
    """Return accepted text for paragraphs [start, end) by index.

    Useful for inspecting a known region, such as the paragraphs around a search hit.
    Returns a list of {para_index, text, has_tracked_change, char_count}.
    """
    root = _load_root(path)
    result = []
    for idx, para in _get_body_paragraphs(root)[start:end]:
        text = _para_accepted_text(para)
        if not include_empty and not text.strip():
            continue
        result.append(
            {
                "para_index": idx,
                "text": text,
                "has_tracked_change": _para_has_tracked_change(para),
                "char_count": len(text),
            }
        )
    return result
