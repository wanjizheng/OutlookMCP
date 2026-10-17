"""Fixed-operation Windows PowerShell bridge; no arbitrary command or path tools."""

import base64
import datetime as dt
import json
import os
import subprocess
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
_lock = threading.Lock()


def next_cursor(rows):
    if not rows:
        return None

    def key(row):
        stamp = dt.datetime.fromisoformat(row["ReceivedTime"].replace("Z", "+00:00"))
        if stamp.tzinfo is None:
            stamp = stamp.astimezone()
        return stamp.astimezone(dt.timezone.utc), (
            row.get("StoreID", "") + "|" + row["EntryID"]
        ).lower()

    stamp, identity = min(map(key, rows))
    return (
        "v2:"
        + base64.b64encode(
            json.dumps(
                {"t": stamp.isoformat(), "k": identity}, separators=(",", ":")
            ).encode()
        ).decode()
    )


def invoke(operation, **args):
    if operation not in ("list", "calendar", "read", "status"):
        raise ValueError("Read-only operation required")
    if os.name != "nt":
        raise RuntimeError("Outlook bridge requires native Windows")
    exe = (
        Path(os.environ["SystemRoot"])
        / "System32/WindowsPowerShell/v1.0/powershell.exe"
    )
    with _lock:
        result = subprocess.run(
            [
                str(exe),
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(ROOT / "scripts/outlook_bridge.ps1"),
            ],
            input=json.dumps(
                dict(operation=operation, **args), ensure_ascii=False
            ).encode("utf-8"),
            capture_output=True,
            timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    try:
        out = json.loads(result.stdout.decode("utf-8-sig").strip())
    except (ValueError, UnicodeDecodeError) as exc:
        raise RuntimeError(
            "Outlook bridge returned invalid output; inspect local diagnostics"
        ) from exc
    if result.returncode or isinstance(out, dict) and "error" in out:
        raise RuntimeError(out.get("error", "Outlook bridge failed"))
    return out


def list_mail(folder="Inbox", count=20, hours=72, cursor="", max_chars=4000):
    if (
        not 1 <= count <= 500
        or not 0 <= hours <= 876000
        or not 1 <= max_chars <= 100000
    ):
        raise ValueError("Invalid mail limits")
    if folder.lower() not in (
        "inbox",
        "inboxall",
        "sent",
        "drafts",
        "deleted",
        "junk",
    ) and not folder.lower().startswith("inbox/"):
        raise ValueError("Unsupported folder")
    out = invoke(
        "list",
        folder=folder,
        count=count,
        hours=hours,
        mode="all",
        cursor=cursor,
        max_chars=max_chars,
    )
    rows = out if isinstance(out, list) else [out]
    return {
        "items": rows,
        "next_cursor": next_cursor(rows) if len(rows) == count else None,
        "source": "live_outlook",
    }


def calendar(days=7, top=100, offset=0):
    if not 1 <= days <= 3660 or not 1 <= top <= 500 or not -3660 <= offset <= 3660:
        raise ValueError("Invalid calendar range")
    out = invoke("calendar", days=days, top=top, offset=offset)
    return {
        "items": out if isinstance(out, list) else [out],
        "source": "live_outlook",
        "limit": top,
    }
