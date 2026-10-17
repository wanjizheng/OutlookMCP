"""Real MCP stdio client acceptance; never prints message bodies."""

import asyncio
import json
import os
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parents[1]


async def main():
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "outlook_mcp.server"],
        cwd=str(ROOT),
        env=dict(os.environ),
    )
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = (await session.list_tools()).tools
            assert len(tools) == 5 and all(t.annotations.readOnlyHint for t in tools)
            calls = [
                ("outlook_status", {}),
                ("list_email", {"count": 2}),
                ("list_calendar", {"days": 7}),
                ("search_email", {"query": "research", "top": 2, "mode": "bm25"}),
            ]
            report = {"tools": [t.name for t in tools], "calls": {}}
            for name, args in calls:
                result = await session.call_tool(name, args)
                if result.isError:
                    raise RuntimeError(f"{name} failed: {result.content}")
                report["calls"][name] = "passed"
                if name == "list_email":
                    payload = result.structuredContent or json.loads(
                        result.content[0].text
                    )
                    item = payload["items"][0]
                    read_result = await session.call_tool(
                        "read_email",
                        {
                            "entry_id": item["EntryID"],
                            "store_id": item.get("StoreID", ""),
                            "source": "live",
                            "max_chars": 500,
                        },
                    )
                    if read_result.isError:
                        raise RuntimeError("Live read_email failed")
                    report["calls"]["read_email"] = "passed"
            forbidden = await session.call_tool("send_email", {})
            assert forbidden.isError
            report["write_tool_rejected"] = True
            print(json.dumps(report))


if __name__ == "__main__":
    asyncio.run(main())
