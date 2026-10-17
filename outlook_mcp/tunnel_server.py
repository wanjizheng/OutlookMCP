"""Private tunnel entry point. Never pass the tunnel credential to MCP tools."""

import os
from pathlib import Path

os.environ.pop("CONTROL_PLANE_API_KEY", None)
os.chdir(Path(__file__).resolve().parents[1])

from .server import mcp  # noqa: E402

if __name__ == "__main__":
    mcp.run(transport="stdio")
