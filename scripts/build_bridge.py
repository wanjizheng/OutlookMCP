"""Generate native read-only bridge from the fixed OpenClaw source (no WSL at runtime)."""

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
s = ROOT / "compat/outlook_reader/scripts"
daemon = (s / "outlook_daemon.ps1").read_text(encoding="utf-8-sig")
a = daemon.index("function Get-OutlookSession")
b = daemon.index("# --- estimate handler ---")
library = daemon[a:b].replace(
    ". (Join-Path $PSScriptRoot 'mail_paging.ps1')",
    ". (Join-Path $PSScriptRoot '../compat/outlook_reader/scripts/mail_paging.ps1')",
)
(ROOT / "scripts/mail_library.ps1").write_text(library.rstrip()+"\n", encoding="utf-8-sig", newline="\n")
calendar = (s / "outlook-calendar").read_text(encoding="utf-8-sig")
a = calendar.index("\\$ns = \\$ol.GetNamespace")
b = calendar.index("\nEOF\n", a)
calendar = calendar[a:b].replace("\\$", "$")
(ROOT / "scripts/calendar_library.ps1").write_text(
    "function Read-OutlookCalendar {\nparam($ol,[int]$DaysVal,[int]$TopVal,[int]$StartOffsetVal)\n"
    + calendar
    + "\n}\n",
    encoding="utf-8-sig",
    newline="\n",
)

fixture = (ROOT / "tests/calendar.ps1").read_text(encoding="utf-8-sig")
fixture = re.sub(r"(?m)^\. \(Join-Path .*\r?\n", "", fixture, count=1)
(s / "test_calendar_read.ps1").write_text(
    (ROOT / "scripts/calendar_library.ps1").read_text(encoding="utf-8-sig")
    + "\n"
    + fixture,
    encoding="utf-8-sig",
    newline="\n",
)
