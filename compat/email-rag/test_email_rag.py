#!/usr/bin/env python3
import argparse
import contextlib
import io
import json
import os
import sqlite3
import subprocess
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

import email_rag


class EmailRagSafetyTests(unittest.TestCase):
    def test_normal_token_launch_translates_bridge_paths(self) -> None:
        with mock.patch.object(email_rag.subprocess,'run',return_value=mock.Mock(stdout='C:\\bridge\\outlook.ps1\n')) as run:
            if email_rag.os.name != 'nt':
                self.assertEqual(email_rag._windows_script_path(Path('/home/user/bridge.ps1')),'C:\\bridge\\outlook.ps1')
                self.assertEqual(run.call_args.args[0][:2],['wslpath','-w'])

    def test_cursor_preserves_tied_message_identity(self) -> None:
        import base64
        rows=[dict(EntryID=e,StoreID='store',ReceivedTime='2026-10-17T10:00:31.123456Z') for e in ('c','a','b')]
        cursor=email_rag._resolve_before_cursor(rows,None)
        decoded=json.loads(base64.b64decode(cursor[3:]))
        self.assertEqual(decoded['k'],'store|a')
        self.assertIn('.123456',decoded['t'])
        with self.assertRaisesRegex(RuntimeError,'did not advance'):
            email_rag._resolve_before_cursor(rows,cursor)

    def test_time_only_full_checkpoint_is_not_reused(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            conn=email_rag.ensure_db(Path(tmp)/'index.db')
            for key,value in {'signature':'s','step_index':'0','processed':'100','before_iso':'2026-10-17T10:00:31Z'}.items():
                email_rag.set_meta(conn,'full_safe.page.'+key,value)
            self.assertIsNone(email_rag._load_full_safe_page_checkpoint(conn,signature='s',step_index=0))
            conn.close()

    @staticmethod
    def _mail_row(entry_id: str, received_time: str) -> dict:
        return {
            "EntryID": entry_id,
            "Subject": f"Subject {entry_id}",
            "Sender": "Sender",
            "SenderEmailAddress": "sender@example.com",
            "ReceivedTime": received_time,
            "Body": f"Body {entry_id}",
            "Attachments": [],
        }

    def test_database_and_parent_are_private(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db_path = Path(tmp) / "nested" / "email_rag.db"
            conn = email_rag.ensure_db(db_path)
            conn.close()
            self.assertEqual(db_path.parent.stat().st_mode & 0o777, 0o700)
            self.assertEqual(db_path.stat().st_mode & 0o777, 0o600)

    def test_pipe_request_timeout_is_enforced(self) -> None:
        children: list[subprocess.Popen] = []

        def new_child() -> subprocess.Popen:
            child = subprocess.Popen(
                [
                    sys.executable,
                    "-c",
                    "import sys,time; sys.stdin.buffer.readline(); time.sleep(10)",
                ],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            children.append(child)
            return child

        def kill_latest() -> None:
            if children and children[-1].poll() is None:
                children[-1].kill()
                children[-1].wait(timeout=2)

        started = time.monotonic()
        try:
            with (
                mock.patch.object(email_rag, "_ensure_pipe_proc", side_effect=new_child),
                mock.patch.object(email_rag, "_force_kill_pipe", side_effect=kill_latest),
            ):
                with self.assertRaisesRegex(RuntimeError, "timed out"):
                    email_rag._pipe_request({"cmd": "ping"}, timeout_sec=0.05)
        finally:
            for child in children:
                if child.poll() is None:
                    child.kill()
                    child.wait(timeout=2)
        self.assertLess(time.monotonic() - started, 1.0)

    def test_failed_full_rebuild_keeps_existing_index_and_anchor(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            db_path = tmp_path / "email_rag.db"
            conn = email_rag.ensure_db(db_path)
            conn.execute(
                """
                INSERT INTO emails(
                    entry_id, stable_key, subject, sender, sender_email,
                    received_time, body, updated_at
                ) VALUES('existing', 'stable', 'subject', 'sender', '',
                         '2026-01-01T00:00:00Z', 'body', '2026-01-01T00:00:00Z')
                """
            )
            email_rag.set_meta(conn, "last_sync_at", "2026-01-01T00:00:00Z")
            conn.commit()
            conn.close()

            args = argparse.Namespace(
                db=db_path,
                no_embed=True,
                embed_backend="local",
                gpu_probe=False,
                page_size=200,
                disable_paged_read=False,
                attachments="none",
                attachments_dir=tmp_path / "attachments",
                attachments_cleanup="none",
                max_attachments_per_email=6,
                attachment_max_chars=12000,
                skip_existing=True,
                embed_base_url="http://127.0.0.1:8000/v1",
                embed_api_key="test",
                embed_model="test",
                embed_batch=1,
                embed_timeout=1,
            )
            plans = [
                {
                    "count": 1,
                    "folder": "Inbox",
                    "hours": 0,
                    "mode": "all",
                    "maxChars": 4000,
                }
            ]
            with mock.patch.object(
                email_rag,
                "run_outlook_mail_read",
                side_effect=RuntimeError("Outlook unavailable"),
            ):
                rc = email_rag._execute_sync(
                    args,
                    plans,
                    profile="test-full",
                    full_rebuild=True,
                    clear_embedding_cache=True,
                    continue_on_plan_error=False,
                )
            self.assertEqual(rc, 1)
            conn = sqlite3.connect(db_path)
            self.assertEqual(
                conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0], 1
            )
            self.assertEqual(
                conn.execute(
                    "SELECT value FROM meta WHERE key='last_sync_at'"
                ).fetchone()[0],
                "2026-01-01T00:00:00Z",
            )
            conn.close()

    def test_cleanup_refuses_outside_email_rag_root(self) -> None:
        report = email_rag._cleanup_exported_attachments(Path("/tmp"))
        self.assertIn("error", report)

    def test_pid_guard_rejects_unrelated_process(self) -> None:
        self.assertFalse(email_rag._is_expected_embed_server_pid(os.getpid()))

    def test_database_creates_compact_vector_table_and_delete_trigger(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db_path = Path(tmp) / "email_rag.db"
            conn = email_rag.ensure_db(db_path)
            now = email_rag.now_iso()
            cur = conn.execute(
                """
                INSERT INTO email_chunks(
                  entry_id, chunk_index, subject, sender, received_time,
                  chunk_text, text_hash, embedding_json, embedding_norm, updated_at
                ) VALUES('e1', 0, 's', 'x', '', 'body', 'hash', '[1,0]', 1, ?)
                """,
                (now,),
            )
            chunk_id = int(cur.lastrowid)
            email_rag._backfill_chunk_vectors(conn)
            row = conn.execute(
                "SELECT dimensions, length(embedding_f32) FROM email_chunk_vectors WHERE chunk_id=?",
                (chunk_id,),
            ).fetchone()
            self.assertEqual(row, (2, 8))
            conn.execute("DELETE FROM email_chunks WHERE chunk_id=?", (chunk_id,))
            self.assertEqual(
                conn.execute(
                    "SELECT COUNT(*) FROM email_chunk_vectors WHERE chunk_id=?",
                    (chunk_id,),
                ).fetchone()[0],
                0,
            )
            conn.close()

    def test_vector_candidates_uses_compact_vectors(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            conn = email_rag.ensure_db(Path(tmp) / "email_rag.db")
            now = email_rag.now_iso()
            for idx, vec in enumerate(([1.0, 0.0], [0.0, 1.0])):
                emb_json = json.dumps(vec)
                cur = conn.execute(
                    """
                    INSERT INTO email_chunks(
                      entry_id, chunk_index, subject, sender, received_time,
                      chunk_text, text_hash, embedding_json, embedding_norm, updated_at
                    ) VALUES(?, 0, ?, 'sender', '', ?, ?, ?, 1, ?)
                    """,
                    (f"e{idx}", f"s{idx}", f"body{idx}", f"h{idx}", emb_json, now),
                )
                email_rag._store_chunk_vector(conn, int(cur.lastrowid), list(vec), now)
            conn.commit()
            with mock.patch.object(
                email_rag, "_embed_query_local", return_value=[1.0, 0.0]
            ):
                hits = email_rag.vector_candidates(
                    conn,
                    "query",
                    1,
                    "",
                    "",
                    "",
                    1,
                    embed_backend="local",
                )
            self.assertEqual(len(hits), 1)
            self.assertEqual(hits[0]["entryId"], "e0")
            self.assertAlmostEqual(hits[0]["vector"], 1.0, places=5)
            conn.close()

    def test_embed_server_failure_is_killed_and_retried_one_shot(self) -> None:
        with (
            mock.patch.object(
                email_rag,
                "_embed_server_request_once",
                side_effect=RuntimeError("timed out"),
            ),
            mock.patch.object(email_rag, "_kill_embed_server") as kill_server,
            mock.patch.object(
                email_rag,
                "embed_texts_local",
                return_value=[[0.1, 0.2]],
            ) as one_shot,
        ):
            vectors = email_rag._embed_server_request(["query"])
        self.assertEqual(vectors, [[0.1, 0.2]])
        kill_server.assert_called_once_with()
        one_shot.assert_called_once()

    def test_local_query_embedding_defaults_to_one_shot(self) -> None:
        with (
            mock.patch.object(email_rag, "LOCAL_QUERY_EMBED_MODE", "oneshot"),
            mock.patch.object(
                email_rag,
                "embed_texts_local",
                return_value=[[0.3, 0.4]],
            ) as one_shot,
            mock.patch.object(email_rag, "_embed_server_request") as persistent,
        ):
            vector = email_rag._embed_query_local("query")
        self.assertEqual(vector, [0.3, 0.4])
        one_shot.assert_called_once()
        persistent.assert_not_called()

    def test_local_query_embedding_retries_a_fresh_process(self) -> None:
        with (
            mock.patch.object(email_rag, "LOCAL_QUERY_EMBED_MODE", "oneshot"),
            mock.patch.object(email_rag, "LOCAL_QUERY_EMBED_ATTEMPTS", 2),
            mock.patch.object(
                email_rag,
                "embed_texts_local",
                side_effect=[RuntimeError("CUDA wedged"), [[0.5, 0.6]]],
            ) as one_shot,
        ):
            vector = email_rag._embed_query_local("query")
        self.assertEqual(vector, [0.5, 0.6])
        self.assertEqual(one_shot.call_count, 2)

    def test_full_safe_defaults_cover_all_time(self) -> None:
        self.assertEqual(email_rag.DEFAULT_FULL_SAFE_HOURS_WINDOWS, [0])
        self.assertEqual(email_rag._parse_hours_windows("0"), [0])
        self.assertGreaterEqual(email_rag.DEFAULT_FULL_SAFE_COUNT_PER_STEP, 50000)
        args = email_rag.build_parser().parse_args(["sync-full-safe"])
        self.assertEqual(args.page_size, email_rag.DEFAULT_FULL_SAFE_PAGE_SIZE)

    def test_full_sync_requires_explicit_rebuild_flag(self) -> None:
        parser = email_rag.build_parser()
        default_args = parser.parse_args(["sync-full"])
        rebuild_args = parser.parse_args(["sync-full", "--rebuild"])
        self.assertFalse(default_args.rebuild)
        self.assertTrue(rebuild_args.rebuild)

    def test_full_safe_worker_command_preserves_relevant_flags(self) -> None:
        parser = email_rag.build_parser()
        args = parser.parse_args(
            [
                "--db",
                "/tmp/test-email-rag.db",
                "sync-full-safe",
                "--from-start",
                "--no-embed",
                "--no-gpu-probe",
                "--page-size",
                "123",
                "--attachments",
                "none",
                "--no-skip-existing",
            ]
        )
        command = email_rag._full_safe_worker_command(args)
        self.assertIn("--foreground", command)
        self.assertIn("--from-start", command)
        self.assertIn("--no-embed", command)
        self.assertIn("--no-gpu-probe", command)
        self.assertIn("--no-skip-existing", command)
        self.assertEqual(command[command.index("--page-size") + 1], "123")

    def test_outlook_page_timeout_scales_with_page_size(self) -> None:
        with mock.patch.object(email_rag, "_pipe_request", return_value=[]) as request:
            rows = email_rag.run_outlook_mail_read(
                200,
                "InboxAll",
                0,
                "all",
                4000,
                "none",
                Path("/tmp/not-used"),
            )
        self.assertEqual(rows, [])
        self.assertEqual(request.call_args.kwargs["timeout_sec"], 230.0)

    def test_page_checkpoint_resumes_from_last_committed_cursor(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db_path = Path(tmp) / "email_rag.db"
            parser = email_rag.build_parser()
            args = parser.parse_args(
                [
                    "--db",
                    str(db_path),
                    "sync-full-safe",
                    "--foreground",
                    "--no-embed",
                    "--attachments",
                    "none",
                    "--attachments-cleanup",
                    "none",
                    "--page-size",
                    "2",
                    "--count-per-step",
                    "4",
                ]
            )
            plan = {
                "count": 4,
                "folder": "InboxAll",
                "hours": 0,
                "mode": "all",
                "maxChars": 4000,
            }
            checkpoints: list[dict] = []
            first_page = [
                self._mail_row("new-1", "2026-07-26T18:10:00Z"),
                self._mail_row("new-2", "2026-07-26T18:09:00Z"),
            ]
            with (
                mock.patch.object(
                    email_rag,
                    "run_outlook_mail_read",
                    side_effect=[first_page, RuntimeError("simulated interruption")],
                ),
                mock.patch.object(email_rag.time, "sleep"),
            ):
                rc = email_rag._execute_sync(
                    args,
                    [plan],
                    profile="full-safe",
                    full_rebuild=False,
                    clear_embedding_cache=False,
                    continue_on_plan_error=False,
                    page_checkpoint=lambda state: checkpoints.append(dict(state)),
                )
            self.assertEqual(rc, 1)
            self.assertEqual(len(checkpoints), 1)
            resume_state = checkpoints[0]
            self.assertEqual(resume_state["processed"], 2)
            self.assertTrue(resume_state["before_iso"])

            older_page = [
                self._mail_row("old-1", "2026-07-26T18:08:00Z"),
                self._mail_row("old-2", "2026-07-26T18:07:00Z"),
            ]
            with mock.patch.object(
                email_rag,
                "run_outlook_mail_read",
                return_value=older_page,
            ) as outlook_read:
                rc = email_rag._execute_sync(
                    args,
                    [plan],
                    profile="full-safe",
                    full_rebuild=False,
                    clear_embedding_cache=False,
                    continue_on_plan_error=False,
                    resume_state=resume_state,
                )
            self.assertEqual(rc, 0)
            self.assertEqual(
                outlook_read.call_args.kwargs["before_iso"],
                resume_state["before_iso"],
            )
            conn = sqlite3.connect(db_path)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0], 4)
            conn.close()

    def test_progress_reconciles_orphaned_running_state(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db_path = Path(tmp) / "email_rag.db"
            conn = email_rag.ensure_db(db_path)
            email_rag.set_meta(conn, "progress_status", "running")
            email_rag.set_meta(conn, "progress_started_at", email_rag.now_iso())
            email_rag.set_meta(conn, "progress_updated_at", email_rag.now_iso())
            email_rag.set_meta(conn, "progress_pid", "999999999")
            conn.commit()
            conn.close()
            args = argparse.Namespace(db=db_path)
            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                rc = email_rag.cmd_progress(args)
            self.assertEqual(rc, 0)
            report = json.loads(stdout.getvalue())
            self.assertEqual(report["status"], "interrupted")
            self.assertEqual(report["rawStatus"], "interrupted")
            self.assertFalse(report["worker"]["alive"])


if __name__ == "__main__":
    unittest.main()
