"""Tunnel bootstrap keeps secrets off argv and sending strictly opt-in."""
import importlib.util
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("tunnel_setup", ROOT / "scripts/setup_tunnel.py")
setup = importlib.util.module_from_spec(spec)
spec.loader.exec_module(setup)


class TunnelSetupTests(unittest.TestCase):
    def invoke(self, args):
        with tempfile.TemporaryDirectory() as directory:
            key = Path(directory) / "key.dpapi"
            key.write_bytes(b"encrypted-test-data")
            crypto = SimpleNamespace(CryptUnprotectData=lambda *a: ("", b"synthetic-test-credential"))
            with patch.object(setup, "KEY", key), patch.dict(sys.modules, win32crypt=crypto), \
                    patch.object(sys, "argv", ["setup_tunnel.py", *args]), \
                    patch.object(setup.subprocess, "call", return_value=0) as call:
                with self.assertRaises(SystemExit) as result:
                    setup.main()
                self.assertEqual(result.exception.code, 0)
                return call.call_args

    def test_run_handshake_and_credential_environment_with_opt_in_send(self):
        for enabled in (False, True):
            with self.subTest(enabled=enabled):
                call = self.invoke(["run"] + (["--enable-send"] if enabled else []))
                command = call.args[0]
                self.assertIn("--mcp.stdio-send-initialized-notification", command)
                self.assertNotIn("synthetic-test-credential", " ".join(command))
                self.assertEqual(call.kwargs["env"]["CONTROL_PLANE_API_KEY"], "synthetic-test-credential")
                self.assertEqual(call.kwargs["env"]["OUTLOOK_MCP_ENABLE_SEND"], "1" if enabled else "0")

    def test_doctor_remains_readonly_without_run_only_flags(self):
        call = self.invoke(["doctor"])
        self.assertIn("--explain", call.args[0])
        self.assertNotIn("--mcp.stdio-send-initialized-notification", call.args[0])
        self.assertEqual(call.kwargs["env"]["OUTLOOK_MCP_ENABLE_SEND"], "0")
        self.assertEqual(call.kwargs["env"]["OUTLOOK_MCP_ENABLE_REFRESH"], "0")

    def test_refresh_is_independent_from_sending(self):
        call = self.invoke(["run", "--enable-refresh"])
        self.assertEqual(call.kwargs["env"]["OUTLOOK_MCP_ENABLE_REFRESH"], "1")
        self.assertEqual(call.kwargs["env"]["OUTLOOK_MCP_ENABLE_SEND"], "0")
