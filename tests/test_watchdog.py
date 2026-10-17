"""Recovery policy and loopback probe; never sends mail or touches Outlook."""
import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch, MagicMock

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("watchdog", ROOT / "scripts/tunnel_watchdog.py")
watchdog = importlib.util.module_from_spec(spec)
spec.loader.exec_module(watchdog)


class WatchdogTests(unittest.TestCase):
    def test_unready_grace_and_recovery(self):
        state = watchdog.Recovery(0)
        self.assertFalse(state.observe(179, False))
        self.assertTrue(state.observe(180, False))
        self.assertFalse(state.observe(181, True))
        self.assertFalse(state.observe(360, False))
        self.assertTrue(state.observe(361, False))

    def test_backoff_is_bounded_and_resets_on_ready(self):
        state = watchdog.Recovery(0)
        delays = [state.delay() for _ in range(15)]
        self.assertEqual(delays[:4], [5, 10, 20, 40])
        self.assertEqual(max(delays), 300)
        state.observe(5000, True)
        self.assertEqual(state.delay(), 5)

    def test_recovery_only_relaunches_transport_and_send_is_explicit(self):
        self.assertNotIn("--enable-send", watchdog.child_command(False))
        command = watchdog.child_command(True)
        self.assertEqual(command[-2:], ["run", "--enable-send"])
        self.assertEqual(Path(command[1]).name, "setup_tunnel.py")
        self.assertNotIn("send_prepared_email", command)

    def test_probe_rejects_external_redirect_credentials_and_missing_file(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "health.url"
            self.assertFalse(watchdog.ready(path))
            for url in ("https://127.0.0.1:123", "http://example.com:123", "http://127.0.0.1:123/evil",
                        "http://user@127.0.0.1:123", "http://127.0.0.1:123?secret=x"):
                path.write_text(url)
                self.assertFalse(watchdog.ready(path))
            path.write_text("http://127.0.0.1:123")
            with patch.object(watchdog, "HTTPConnection") as connection:
                response = connection.return_value.getresponse.return_value
                response.status = 302
                self.assertFalse(watchdog.ready(path))
                response.status = 200
                response.read.return_value = b"ready\n"
                self.assertTrue(watchdog.ready(path))
                response.read.return_value = b"not ready"
                self.assertFalse(watchdog.ready(path))
                response.read.side_effect = watchdog.HTTPException("Disconnected")
                self.assertFalse(watchdog.ready(path))

    def test_stop_only_targets_owned_child_tree(self):
        child = MagicMock(pid=1234)
        child.poll.return_value = None
        with patch.object(watchdog.subprocess, "run") as run:
            watchdog.terminate_tree(child)
            self.assertEqual(run.call_args.args[0][-4:], ["/PID", "1234", "/T", "/F"])
            child.wait.assert_called_once()
        child.poll.return_value = 1
        with patch.object(watchdog.subprocess, "run") as run:
            watchdog.terminate_tree(child)
            run.assert_not_called()
