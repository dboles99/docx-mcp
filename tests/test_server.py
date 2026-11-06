"""The MCP layer: tools are registered and reach the real DOCX logic."""
from __future__ import annotations

import asyncio
import subprocess
import sys
from pathlib import Path

from conftest import para
from mcp import Client, StdioServerParameters

from docx_mcp.server import mcp as server

ROOT = Path(__file__).resolve().parents[1]
ALL_TOOLS = {
    "read_docx_accepted_text",
    "find_docx_term",
    "check_docx_term_variants",
    "get_docx_structure_stats",
    "scan_docx_meta_notes",
    "list_docx_paragraph_range",
}


def structured(tool: str, args: dict) -> dict:
    result = asyncio.run(server.call_tool(tool, args))
    assert not result.is_error, f"{tool} failed: {result.content}"
    return result.structured_content


def test_all_six_tools_are_registered():
    assert {tool.name for tool in asyncio.run(server.list_tools())} == ALL_TOOLS


def test_stats_tool(docx):
    out = structured(
        "get_docx_structure_stats", {"path": docx(para("Hello world") + para("Second one here"))}
    )
    assert out["paragraph_count"] == 2
    assert out["word_count_approx"] == 5
    assert out["has_tracked_changes"] is False


def test_read_tool(docx):
    path = docx(para("First paragraph text") + para("Second one here"))
    rows = structured("read_docx_accepted_text", {"path": path, "limit": 10})["result"]
    assert [row["text"] for row in rows] == ["First paragraph text", "Second one here"]


def test_find_tool(docx):
    rows = structured(
        "find_docx_term", {"path": docx(para("The quick brown fox")), "pattern": "quick"}
    )["result"]
    assert [row["match"] for row in rows] == ["quick"]


def test_variants_tool(docx):
    out = structured(
        "check_docx_term_variants", {"path": docx(para("colour color")), "terms": ["colour"]}
    )
    assert out["counts"] == {"colour": 1}


def test_scan_tool(docx):
    rows = structured("scan_docx_meta_notes", {"path": docx(para("TODO: finish this."))})["result"]
    assert rows[0]["para_index"] == 0


def test_range_tool(docx):
    path = docx(para("Zero") + para("One") + para("Two"))
    rows = structured("list_docx_paragraph_range", {"path": path, "start": 1, "end": 2})["result"]
    assert [row["text"] for row in rows] == ["One"]


async def call_as_client(tool: str, args: dict):
    async with Client(server) as client:
        return await client.call_tool(tool, args)


def test_a_missing_file_is_an_error_the_agent_can_read(tmp_path: Path):
    result = asyncio.run(
        call_as_client("read_docx_accepted_text", {"path": str(tmp_path / "absent.docx")})
    )
    assert result.is_error
    assert "FileNotFoundError" in result.content[0].text


def test_a_damaged_file_is_an_error_the_agent_can_read(tmp_path: Path):
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"not a zip")
    result = asyncio.run(call_as_client("get_docx_structure_stats", {"path": str(bad)}))
    assert result.is_error
    assert "BadZipFile" in result.content[0].text


def test_a_bad_pattern_is_an_error_the_agent_can_read(docx):
    result = asyncio.run(call_as_client("find_docx_term", {"path": docx(para("x")), "pattern": "("}))
    assert result.is_error
    assert "error" in result.content[0].text.lower()


def test_the_server_works_over_real_stdio(docx):
    path = docx(para("Text sent through a subprocess."))
    parameters = StdioServerParameters(command=sys.executable, args=["-m", "docx_mcp.server"])

    async def run() -> tuple[set[str], str]:
        async with Client(parameters) as client:
            tools = await client.list_tools()
            result = await client.call_tool("read_docx_accepted_text", {"path": path})
            assert not result.is_error
            return {tool.name for tool in tools.tools}, result.structured_content["result"][0]["text"]

    names, text = asyncio.run(run())
    assert names == ALL_TOOLS
    assert text == "Text sent through a subprocess."


def test_the_console_script_starts_a_server():
    script = Path(sys.executable).with_name("docx-mcp")
    parameters = StdioServerParameters(command=str(script))

    async def run() -> int:
        async with Client(parameters) as client:
            return len((await client.list_tools()).tools)

    assert asyncio.run(run()) == len(ALL_TOOLS)


def test_the_example_runs():
    result = subprocess.run(
        [sys.executable, str(ROOT / "examples" / "call_over_stdio.py")],
        capture_output=True,
        text=True,
        check=True,
        timeout=120,
    )
    lines = result.stdout.splitlines()
    assert "accepted: The pilot covered twelve sites and met its target in the first year." in lines
    assert "search for deleted word: 0 hits" in lines
    assert 'variants: {"cyber security": 1, "cybersecurity": 1}' in lines
    assert lines[-1].startswith("notes: TODO:")


def test_the_docx_logic_does_not_import_the_sdk():
    code = "import sys, docx_mcp.docx_tools; print('mcp' in sys.modules)"
    result = subprocess.run(
        [sys.executable, "-c", code], capture_output=True, text=True, check=True, timeout=60
    )
    assert result.stdout.strip() == "False"
