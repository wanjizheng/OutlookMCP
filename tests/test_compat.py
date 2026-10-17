import base64
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class CompatTests(unittest.TestCase):
    @unittest.skipUnless(sys.platform == "win32", "Windows PowerShell regression")
    def test_real_powershell_paging(self):
        subprocess.run(
            ["powershell.exe", "-NoProfile", "-File", str(ROOT / "tests/paging.ps1")],
            check=True,
        )

    @unittest.skipUnless(sys.platform == "win32", "Windows PowerShell regression")
    def test_calendar_overlap_boundaries(self):
        subprocess.run(
            ["powershell.exe", "-NoProfile", "-File", str(ROOT / "tests/calendar.ps1")],
            check=True,
        )

    def test_python_cursor_keeps_identity_and_precision(self):
        # Load just the pure function to avoid importing Linux-only runtime code.
        import ast
        import datetime as dt

        tree = ast.parse(
            (ROOT / "compat/email-rag/email_rag.py").read_text(encoding="utf-8")
        )
        fn = next(
            n
            for n in tree.body
            if isinstance(n, ast.FunctionDef) and n.name == "_resolve_before_cursor"
        )
        ns = {
            "Any": object,
            "json": json,
            "parse_utc_timestamp": lambda t: dt.datetime.fromisoformat(
                t.replace("Z", "+00:00")
            ),
        }
        exec(compile(ast.Module(body=[fn], type_ignores=[]), "cursor", "exec"), ns)
        rows = [
            dict(EntryID=x, StoreID="s", ReceivedTime="2026-10-17T10:00:31.123456Z")
            for x in ("c", "a", "b")
        ]
        cursor = ns["_resolve_before_cursor"](rows, None)
        decoded = json.loads(base64.b64decode(cursor[3:]))
        self.assertEqual(decoded["k"], "s|a")
        self.assertIn(".123456", decoded["t"])
        with self.assertRaises(RuntimeError):
            ns["_resolve_before_cursor"](rows, cursor)

    @unittest.skipUnless(sys.platform == "win32", "Windows PowerShell regression")
    def test_calendar_evidence_never_adds_appointment(self):
        raw = (ROOT / "compat/outlook_reader/scripts/outlook-calendar").read_text(
            encoding="utf-8"
        )
        raw = raw.replace("\\$", "$")
        start = raw.index("function Get-ResponseStatusLabel")
        end = raw.index("$result = [System.Collections.ArrayList]::new()", start)
        code = (
            raw[start:end]
            + """
$script:called=$false
$script:fixtureMsg=[pscustomobject]@{MessageClass='IPM.Schedule.Meeting.Resp.Pos';ReceivedTime=[datetime]::Now}
$script:fixtureMsg|Add-Member ScriptMethod GetAssociatedAppointment {param($add) if($add){throw 'WRITE ATTEMPT'}; $script:called=$true; return $null}
$script:fixtureItems=[pscustomobject]@{Count=1}
$script:fixtureItems|Add-Member ScriptMethod Sort {param($field,$desc)}
$script:fixtureItems|Add-Member ScriptMethod Item {param($index) return $script:fixtureMsg}
$ns=[pscustomobject]@{}
$ns|Add-Member ScriptMethod GetDefaultFolder {param($id) return [pscustomobject]@{Items=$script:fixtureItems}}
$null=Get-MeetingResponseEvidence $ns ([datetime]::Today) ([datetime]::Today.AddDays(7))
if(-not $script:called){throw 'No evidence method was exercised'}
"""
        )
        with tempfile.TemporaryDirectory(dir=ROOT / "artifacts") as tmp:
            p = Path(tmp) / "calendar.ps1"
            p.write_text(code, encoding="utf-8-sig")
            subprocess.run(
                ["powershell.exe", "-NoProfile", "-File", str(p)], check=True
            )


if __name__ == "__main__":
    unittest.main()
