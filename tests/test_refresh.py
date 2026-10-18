import asyncio
import datetime as dt
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from outlook_mcp import refresh, rag, cli

ROOT = Path(__file__).resolve().parents[1]


class RefreshTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=ROOT / "artifacts")
        self.root = Path(self.temp.name)
        self.environment = patch.dict(os.environ, OUTLOOK_MCP_DATA_DIR=str(self.root))
        self.environment.start()

    def tearDown(self):
        self.environment.stop()
        self.temp.cleanup()

    def state(self):
        state = {"job_id": "a" * 32, "status": "running", "created_epoch": time.time()}
        refresh.save(state)
        return state

    def test_fixed_start_and_repeat_requests_join_one_job(self):
        with patch.object(refresh.subprocess, "Popen") as spawn, patch.dict(os.environ, CONTROL_PLANE_API_KEY="fake-test-key"):
            first = refresh.start()
            second = refresh.start()
        spawn.assert_called_once()
        self.assertEqual(first["job_id"], second["job_id"])
        self.assertTrue(second["already_running"])
        command = spawn.call_args.args[0]
        self.assertEqual(command[1:3], ["-m", "outlook_mcp.refresh"])
        self.assertEqual(len(command), 4)
        self.assertNotIn("CONTROL_PLANE_API_KEY", spawn.call_args.kwargs["env"])
        self.assertEqual(spawn.call_args.kwargs["env"]["EMAIL_RAG_AUTO_INSTALL"], "0")
        self.assertFalse(spawn.call_args.kwargs.get("shell", False))
        self.assertEqual(first["scope"]["resume"], True)
        self.assertNotIn("created_epoch", first)

    def test_launch_failure_and_stale_or_reused_process_are_not_success(self):
        with patch.object(refresh.subprocess, "Popen", side_effect=OSError):
            result = refresh.start()
        self.assertEqual(result["status"], "failed")
        result.update(status="running", worker_pid=123, process_created=5)
        refresh.save(result)
        for created in (None, 100):
            with patch.object(refresh, "process_started", return_value=created):
                self.assertEqual(refresh.status(result["job_id"])["status"], "interrupted")
        with self.assertRaises(ValueError):
            refresh.status("../runtime-key.dpapi")

    def test_unknown_status_does_not_create_index_or_job_directory(self):
        with self.assertRaises(ValueError):
            refresh.status("b" * 32)
        self.assertFalse((self.root / "refresh-jobs").exists())
        self.assertFalse(rag.db_path().exists())

    def test_busy_writer_does_not_call_outlook_or_embed(self):
        with patch.object(cli, "writer_lock", side_effect=RuntimeError("busy")), \
                patch.object(refresh.synchronizer, "run") as sync, patch.object(rag, "embed_pending") as embed:
            state = self.state()
            refresh.execute(state)
        self.assertEqual(state["status"], "failed")
        sync.assert_not_called()
        embed.assert_not_called()





    def test_recent_vector_budget_skips_history_and_can_continue_backlog(self):
        rag.initialize()
        with rag.connect(False) as c:
            rag.ingest([{"EntryID": "history", "Body": "old", "ReceivedTime": "2020-01-01"},
                        {"EntryID": "recent1", "Body": "new1", "ReceivedTime": dt.datetime.now(dt.UTC).isoformat()},
                        {"EntryID": "recent2", "Body": "new2", "ReceivedTime": dt.datetime.now(dt.UTC).isoformat()}], c)
        with patch.object(rag, "embed", side_effect=lambda texts: [[1.0, 0.0] for _ in texts]):
            self.assertEqual(rag.embed_pending(1, recent_hours=72)["embedded"], 1)
            self.assertEqual(rag.embed_pending(1, recent_hours=72)["embedded"], 1)
            self.assertEqual(rag.embed_pending(1, recent_hours=72)["embedded"], 0)
        with rag.connect() as c:
            ids = {r[0] for r in c.execute("SELECT c.entry_id FROM email_chunks c JOIN email_chunk_vectors v ON c.chunk_id=v.chunk_id")}
        self.assertEqual(ids, {"recent1", "recent2"})

    def test_real_worker_timeout_kills_its_descendant_tree(self):
        state = self.state()
        code = '''import subprocess,sys,time,json
from outlook_mcp import refresh
refresh.TIMEOUT=1
def execute(state):
 p=subprocess.Popen([sys.executable,'-c','import time; time.sleep(60)'])
 (refresh.directory()/'test-child.json').write_text(json.dumps({'pid':p.pid}))
 time.sleep(60)
refresh.execute=execute
refresh.worker('a'*32)
'''
        result = subprocess.run([sys.executable, "-c", code], cwd=ROOT, env=dict(os.environ),
                                capture_output=True, timeout=15)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(refresh.status(state["job_id"])["status"], "timed_out")
        pid = json.loads((refresh.directory() / "test-child.json").read_text())["pid"]
        self.assertIsNone(refresh.process_started(pid))

    def test_opt_in_stdio_tools_and_readonly_status_without_starting_real_job(self):
        state = self.state()
        state["status"] = "completed"
        refresh.save(state)
        async def check():
            params = StdioServerParameters(command=sys.executable, args=["-m", "outlook_mcp.server"], cwd=str(ROOT),
                env=dict(os.environ, OUTLOOK_MCP_ENABLE_REFRESH="1", OUTLOOK_MCP_ENABLE_SEND="0"))
            async with stdio_client(params) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    tools = {t.name: t for t in (await session.list_tools()).tools}
                    self.assertEqual(len(tools), 8)
                    self.assertFalse(tools["refresh_email_index"].annotations.readOnlyHint)
                    self.assertFalse(tools["refresh_email_index"].annotations.destructiveHint)
                    self.assertEqual(tools["refresh_email_index"].inputSchema.get("properties", {}), {})
                    self.assertTrue(tools["get_refresh_status"].annotations.readOnlyHint)
                    self.assertTrue(tools["full_sync_email_index"].annotations.destructiveHint)
                    result = await session.call_tool("get_refresh_status", {"job_id": state["job_id"]})
                    self.assertFalse(result.isError)
                    data = result.structuredContent or json.loads(result.content[0].text)
                    self.assertEqual(data["status"], "completed")
                    rejected = await session.call_tool("get_refresh_status", {"job_id": "../secret"})
                    self.assertTrue(rejected.isError)
        asyncio.run(asyncio.wait_for(check(), timeout=30))
