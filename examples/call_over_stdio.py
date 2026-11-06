"""Start the server as a subprocess and call it the way an MCP client would.

    python examples/call_over_stdio.py

The document is invented and built in a temporary folder.
"""
from __future__ import annotations

import asyncio
import json
import sys
import tempfile
from pathlib import Path

from make_sample_docx import write_sample
from mcp import Client, StdioServerParameters


async def main() -> None:
    server = StdioServerParameters(command=sys.executable, args=["-m", "docx_mcp.server"])
    with tempfile.TemporaryDirectory() as tmp:
        path = str(write_sample(Path(tmp) / "sample.docx"))
        async with Client(server) as client:
            tools = await client.list_tools()
            print("tools:", ", ".join(sorted(tool.name for tool in tools.tools)))

            async def call(tool: str, **arguments: object) -> object:
                result = await client.call_tool(tool, {"path": path, **arguments})
                if result.is_error:
                    raise RuntimeError(f"{tool} failed: {result.content}")
                return result.structured_content

            stats = await call("get_docx_structure_stats")
            stats = {key: value for key, value in stats.items() if key != "path"}
            print("stats:", json.dumps(stats, sort_keys=True))

            rows = await call("list_docx_paragraph_range", start=1, end=2)
            print("accepted:", rows["result"][0]["text"])

            hits = await call("find_docx_term", pattern="failed")
            print("search for deleted word:", len(hits["result"]), "hits")

            variants = await call(
                "check_docx_term_variants", terms=["cyber security", "cybersecurity"]
            )
            print("variants:", json.dumps(variants["counts"], sort_keys=True))

            notes = await call("scan_docx_meta_notes")
            print("notes:", notes["result"][0]["text"])


if __name__ == "__main__":
    asyncio.run(main())
