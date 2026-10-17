"""Exercise the first hybrid query in a fresh, real stdio server process."""

import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
import unittest
from pathlib import Path

from outlook_mcp import rag

ROOT = Path(__file__).resolve().parents[1]


class StdioTests(unittest.TestCase):
    def test_first_hybrid_query_completes_after_stdio_threads_start(self):
        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:
            path = Path(tmp) / "email_rag.db"
            rag.initialize(path)
            with rag.connect(False, path) as c:
                rag.ingest(
                    [
                        {
                            "EntryID": "stdio-fixture",
                            "Body": "attachmentregressionneedle",
                        }
                    ],
                    c,
                )
            with tempfile.TemporaryFile(mode="w+", encoding="utf-8") as errors:
                process = subprocess.Popen(
                    [sys.executable, "-m", "outlook_mcp.server"],
                    cwd=ROOT,
                    env=dict(os.environ, OUTLOOK_MCP_DATA_DIR=tmp),
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=errors,
                    text=True,
                    encoding="utf-8",
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                responses = queue.Queue()

                def read():
                    for line in process.stdout:
                        responses.put(json.loads(line))

                reader = threading.Thread(target=read, daemon=True)
                reader.start()

                def send(message):
                    process.stdin.write(json.dumps(message) + "\n")
                    process.stdin.flush()

                def request(identity, method, params):
                    send(
                        {
                            "jsonrpc": "2.0",
                            "id": identity,
                            "method": method,
                            "params": params,
                        }
                    )
                    response = responses.get(timeout=20)
                    self.assertEqual(response.get("id"), identity)
                    self.assertNotIn("error", response)
                    return response["result"]

                try:
                    request(
                        1,
                        "initialize",
                        {
                            "protocolVersion": "2025-06-18",
                            "capabilities": {},
                            "clientInfo": {"name": "regression", "version": "1"},
                        },
                    )
                    send({"jsonrpc": "2.0", "method": "notifications/initialized"})
                    result = request(
                        2,
                        "tools/call",
                        {
                            "name": "search_email",
                            "arguments": {
                                "query": "attachmentregressionneedle",
                                "mode": "hybrid",
                                "top": 1,
                            },
                        },
                    )
                    self.assertFalse(result.get("isError"))
                    payload = result.get("structuredContent") or json.loads(
                        result["content"][0]["text"]
                    )
                    self.assertEqual(payload["items"][0]["entry_id"], "stdio-fixture")
                finally:
                    process.stdin.close()
                    try:
                        process.wait(timeout=3)
                    except subprocess.TimeoutExpired:
                        if os.name == "nt":
                            subprocess.run(
                                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                                capture_output=True,
                                timeout=10,
                                check=False,
                            )
                        else:
                            process.kill()
                        process.wait(timeout=10)
                    reader.join(timeout=2)
                    process.stdout.close()


if __name__ == "__main__":
    unittest.main()
