import asyncio
import importlib.util
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from outlook_mcp import bridge, cli, rag
from outlook_mcp.server import mcp

ROOT = Path(__file__).resolve().parents[1]


class BackendTests(unittest.TestCase):
    def test_compact_import_preserves_text_and_only_valid_vectors(self):
        from scripts import compact_legacy

        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:
            source = Path(tmp) / "legacy.db"
            dest = Path(tmp) / "native.db"
            rag.initialize(source)
            with rag.connect(False, source) as c:
                rag.ingest(
                    [
                        {
                            "EntryID": "message",
                            "Body": "attachment research text",
                            "_AttachmentRecords": [
                                {
                                    "name": "native.txt",
                                    "status": "ok",
                                    "text": "nativecachedneedle",
                                }
                            ],
                        }
                    ],
                    c,
                )
                cid = c.execute("SELECT chunk_id FROM email_chunks").fetchone()[0]
                c.execute(
                    "INSERT INTO email_chunk_vectors VALUES(?,?,?,?)",
                    (cid, 2, b"12345678", "now"),
                )
                c.execute(
                    "INSERT INTO email_chunk_vectors VALUES(?,?,?,?)",
                    (999, 2, b"12345678", "now"),
                )
                c.execute("CREATE TABLE embedding_cache(payload TEXT)")
                c.execute("INSERT INTO embedding_cache VALUES(?)", ("redundant cache",))
            with mock.patch.object(
                sys,
                "argv",
                ["compact", "--source", str(source), "--destination", str(dest)],
            ):
                compact_legacy.main()
            with rag.connect(True, dest) as c:
                self.assertIn(
                    "nativecachedneedle",
                    c.execute(
                        "SELECT records_json FROM native_attachment_state"
                    ).fetchone()[0],
                )
                self.assertEqual(
                    c.execute("SELECT count(*) FROM email_chunks").fetchone()[0], 1
                )
                self.assertEqual(
                    c.execute("SELECT count(*) FROM email_chunk_vectors").fetchone()[0],
                    1,
                )
                self.assertEqual(
                    c.execute(
                        "SELECT count(*) FROM sqlite_master WHERE name='embedding_cache'"
                    ).fetchone()[0],
                    0,
                )
                self.assertEqual(
                    c.execute(
                        "SELECT count(*) FROM email_chunks_fts WHERE email_chunks_fts MATCH 'attachment'"
                    ).fetchone()[0],
                    1,
                )

    @unittest.skipUnless(os.name == "nt", "Native Windows writer lock")
    def test_writer_lock_prevents_overlapping_jobs(self):
        with (
            tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp,
            mock.patch.dict(os.environ, {"OUTLOOK_MCP_DATA_DIR": tmp}),
        ):
            with cli.writer_lock():
                with self.assertRaisesRegex(RuntimeError, "Another native"):
                    with cli.writer_lock():
                        self.fail("Second writer acquired the lock")

    def test_native_backfill_resumes_committed_page(self):
        with (
            tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp,
            mock.patch.dict(os.environ, {"OUTLOOK_MCP_DATA_DIR": tmp}),
        ):
            first = {
                "items": [{"EntryID": "a", "Body": "one"}],
                "next_cursor": "cursor-a",
            }
            second = {"items": [{"EntryID": "b", "Body": "two"}], "next_cursor": None}
            with mock.patch.object(bridge, "list_mail", return_value=first):
                self.assertTrue(cli._sync(["Inbox"], 0, 1, True)["Inbox"]["has_more"])
            with mock.patch.object(bridge, "list_mail", return_value=second) as read:
                self.assertFalse(cli._sync(["Inbox"], 0, 1, True)["Inbox"]["has_more"])
                self.assertEqual(read.call_args.args[3], "cursor-a")
            with rag.connect() as c:
                self.assertEqual(
                    c.execute("SELECT count(*) FROM emails").fetchone()[0], 2
                )

    def test_unknown_bridge_operations_rejected_before_execution(self):
        with mock.patch("subprocess.run") as run:
            with self.assertRaises(ValueError):
                bridge.invoke("send")
            run.assert_not_called()

    def test_limits_rejected_before_outlook(self):
        with mock.patch.object(bridge, "invoke") as invoke:
            for kwargs in (
                {"count": 0},
                {"count": 501},
                {"folder": "../../private"},
                {"max_chars": 100001},
            ):
                with self.assertRaises(ValueError):
                    bridge.list_mail(**kwargs)
            with self.assertRaises(ValueError):
                bridge.calendar(days=0)
            invoke.assert_not_called()

    def test_index_search_is_readonly_and_handles_fts_punctuation(self):
        with (
            tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp,
            mock.patch.dict(os.environ, {"OUTLOOK_MCP_DATA_DIR": tmp}),
        ):
            rag.initialize()
            with rag.connect(False) as c:
                rag.ingest(
                    [
                        dict(
                            EntryID="fixture",
                            Subject="Research project",
                            Body="A synthetic deadline for testing only.",
                            ReceivedTime="2026-10-17T12:00:00Z",
                        )
                    ],
                    c,
                )
            before = rag.db_path().read_bytes()
            result = rag.search('"research" OR deadline*', mode="bm25")
            self.assertEqual(result["items"][0]["entry_id"], "fixture")
            self.assertEqual(before, rag.db_path().read_bytes())
            with rag.connect() as c:
                with self.assertRaises(Exception):
                    c.execute("DELETE FROM emails")
            self.assertTrue(rag.read_indexed("fixture", max_chars=4)["truncated"])

    def test_hybrid_reports_fallback(self):
        with (
            tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp,
            mock.patch.dict(os.environ, {"OUTLOOK_MCP_DATA_DIR": tmp}),
            mock.patch.object(rag, "embed", side_effect=RuntimeError("missing")),
        ):
            rag.initialize()
            result = rag.search("research")
            self.assertEqual(result["mode"], "bm25")
            self.assertIsNotNone(result["fallback"])

    def test_reingest_updates_text_without_duplicate_chunks(self):
        with (
            tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp,
            mock.patch.dict(os.environ, {"OUTLOOK_MCP_DATA_DIR": tmp}),
        ):
            rag.initialize()
            row = dict(EntryID="one", Subject="Test", Body="old")
            with rag.connect(False) as c:
                rag.ingest([row], c)
                rag.ingest([row], c)
                self.assertEqual(
                    c.execute("SELECT count(*) FROM email_chunks").fetchone()[0], 1
                )
                row["Body"] = "new"
                rag.ingest([row], c)
            self.assertEqual(rag.read_indexed("one")["body"], "new")

    def test_only_read_tools_registered(self):
        tools = asyncio.run(mcp.list_tools())
        self.assertEqual(
            {t.name for t in tools},
            {
                "outlook_status",
                "search_email",
                "list_email",
                "read_email",
                "list_calendar",
            },
        )
        self.assertTrue(
            all(
                t.annotations.readOnlyHint and not t.annotations.destructiveHint
                for t in tools
            )
        )

    def test_registration_preserves_unrelated_settings(self):
        spec = importlib.util.spec_from_file_location(
            "register", ROOT / "scripts/register_client.py"
        )
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:
            p = Path(tmp) / "config.toml"
            original = b"# keep me\r\n[features]\r\nmemories = true\r\n"
            p.write_bytes(original)
            module.register(p)
            self.assertTrue(p.read_bytes().startswith(original))
            with self.assertRaises(RuntimeError):
                module.register(p)


if __name__ == "__main__":
    unittest.main()
