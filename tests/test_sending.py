import asyncio
import json
import os
import subprocess
import unittest
from types import SimpleNamespace as NS
from unittest.mock import patch

from mcp.server.fastmcp import FastMCP
from outlook_mcp import sending
from outlook_mcp.server import register_sending


class SendingTests(unittest.TestCase):
    def setUp(self):
        self.env = patch.dict(os.environ, {"OUTLOOK_MCP_ENABLE_SEND": "1"})
        self.env.start()
        sending._pending.clear()

    def tearDown(self):
        self.env.stop()

    def prepare(self, **kwargs):
        payload = dict(sender_account="sender@example.invalid", to=["to@example.invalid"],
                       cc=["cc@example.invalid"], bcc=["bcc@example.invalid"], subject="Test", body="Full body\nSecond line")
        payload.update(kwargs)
        return sending.prepare(**payload)

    def test_opt_in_and_reject_header_injection(self):
        with patch.dict(os.environ, {"OUTLOOK_MCP_ENABLE_SEND": "0"}):
            with self.assertRaises(RuntimeError):
                self.prepare()
        for fields in ({"subject": "Hi\r\nBcc: x"}, {"to": ["Name <a@example.invalid>"]},
                       {"to": []}, {"to": ["x@example.invalid\nfoo"]}, {"body": "hidden\x00text"}):
            with self.assertRaises(ValueError):
                self.prepare(**fields)

    def test_prepare_never_calls_outlook_and_requires_subsequent_chat_approval(self):
        with patch.object(sending.subprocess, "run") as run:
            result = self.prepare()
        run.assert_not_called()
        self.assertFalse(result["sent"])
        self.assertIn("STOP", result["next_step"])

    @unittest.skipUnless(os.name == "nt", "Native Windows bridge")
    def test_preview_cannot_mutate_and_submission_uses_fixed_powershell_once(self):
        result = self.prepare()
        result["message"]["body"] = "tampered"
        with patch.object(sending.subprocess, "run", return_value=NS(returncode=0, stdout=b'{"sent":true}')) as run:
            sending.send(result["preparation_id"])
            actual = json.loads(run.call_args.kwargs["input"])
            self.assertEqual(actual["body"], "Full body\nSecond line")
            self.assertEqual(actual["bcc"], ["bcc@example.invalid"])
            command = run.call_args.args[0]
            self.assertTrue(command[0].lower().endswith("powershell.exe"))
            self.assertTrue(command[-1].endswith("send_prepared.ps1"))
            with self.assertRaises(ValueError):
                sending.send(result["preparation_id"])
            self.assertEqual(run.call_count, 1)

    @unittest.skipUnless(os.name == "nt", "Native Windows bridge")
    def test_expiry_and_timeout_are_consumed(self):
        result = self.prepare()
        with patch.object(sending.time, "monotonic", return_value=10**12):
            with self.assertRaises(ValueError):
                sending.send(result["preparation_id"])
        result = self.prepare()
        with patch.object(sending.subprocess, "run", side_effect=subprocess.TimeoutExpired("worker", 90)):
            with self.assertRaisesRegex(RuntimeError, "unknown"):
                sending.send(result["preparation_id"])
        with self.assertRaises(ValueError):
            sending.send(result["preparation_id"])

    def test_confirmation_test_never_spawns_outlook_or_window(self):
        result = self.prepare()
        with patch.object(sending.subprocess, "run") as run:
            output = sending.send(result["preparation_id"], confirmation_test=True)
        run.assert_not_called()
        self.assertEqual(output["status"], "chat_confirmation_test_passed")
        self.assertEqual((output["sent"], output["test"]), (False, True))
        with self.assertRaises(ValueError):
            sending.send(result["preparation_id"], confirmation_test=True)

    @unittest.skipUnless(os.name == "nt", "Native PowerShell mocked COM")
    def test_existing_skill_sender_and_legacy_modes_with_mock_com(self):
        subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                        "-File", str(sending.ROOT / "tests/sending.ps1")], check=True, timeout=30)

    def test_annotations_require_non_idempotent_write_and_chat_confirmation(self):
        server = FastMCP("test")
        register_sending(server)
        tools = {t.name: t for t in asyncio.run(server.list_tools())}
        self.assertTrue(tools["prepare_email"].annotations.readOnlyHint)
        self.assertFalse(tools["prepare_email"].annotations.idempotentHint)
        annotations = tools["send_prepared_email"].annotations
        self.assertFalse(annotations.readOnlyHint)
        self.assertFalse(annotations.idempotentHint)
        self.assertTrue(annotations.destructiveHint)
        self.assertIn("subsequent chat turn", tools["send_prepared_email"].description)
