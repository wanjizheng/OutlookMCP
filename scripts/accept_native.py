"""Bounded live native acceptance, with no email content in the report."""

import json
import os
from pathlib import Path

from outlook_mcp import bridge, rag

ROOT = Path(__file__).resolve().parents[1]
os.environ["OUTLOOK_MCP_DATA_DIR"] = str(ROOT / "artifacts/native-acceptance")
rows = bridge.list_mail(count=2)["items"]
rag.initialize()
with rag.connect(False) as c:
    rag.ingest(rows, c)
embedded = rag.embed_pending(8)
result = rag.search("research collaboration", top=2, mode="vector")
assert result["mode"] == "vector" and result["items"]
print(
    json.dumps(
        {
            "native_messages_indexed": len(rows),
            "vector_search_results": len(result["items"]),
            "embedding": embedded,
            "wsl_required": False,
        }
    )
)
