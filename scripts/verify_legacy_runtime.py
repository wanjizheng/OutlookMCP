"""Explicit OpenClaw deployment acceptance, run once in its existing WSL host."""

import json
import sys
from pathlib import Path

legacy = Path.home() / ".openclaw/workspace/skills/email-rag/scripts"
sys.path.insert(0, str(legacy))
import email_rag  # noqa: E402 - explicitly test the deployed source

kwargs = dict(
    count=2,
    folder="InboxAll",
    hours=0,
    mode="all",
    max_chars=1000,
    attachments_mode="none",
    attachments_dir=Path(
        "/mnt/e/wanjizheng/MyProject/OutlookMCP/artifacts/legacy-attachments"
    ),
)
try:
    first = email_rag.run_outlook_mail_read(**kwargs)
    cursor = email_rag._resolve_before_cursor(first, None)
    second = email_rag.run_outlook_mail_read(**kwargs, before_iso=cursor)
    ids = [r["EntryID"] for r in first + second]
    assert len(ids) == 4 and len(set(ids)) == 4
    assert all(
        r["ReceivedTime"].endswith("Z") and r.get("StoreID") for r in first + second
    )
    print(
        json.dumps(
            {
                "legacy_daemon_pages": 2,
                "unique_messages": len(ids),
                "utc_and_store_ids": True,
            }
        )
    )
finally:
    email_rag.shutdown_pipe()
