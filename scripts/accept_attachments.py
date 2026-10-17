"""Bounded live attachment acceptance in a separate local index; reports no content."""

import json
import os
import re
import subprocess
from pathlib import Path
from unittest.mock import patch

from outlook_mcp import attachments, bridge, rag

ROOT = Path(__file__).resolve().parents[1]
os.environ["OUTLOOK_MCP_DATA_DIR"] = str(ROOT / "artifacts/attachment-acceptance")
original = subprocess.run
calls = []


def native_only(args, *a, **kwargs):
    if any("wsl" in str(x).lower() or "/mnt/" in str(x).lower() for x in args):
        raise RuntimeError("Unexpected WSL dependency")
    calls.append(Path(str(args[0])).name)
    return original(args, *a, **kwargs)


with patch("subprocess.run", side_effect=native_only):
    rows = bridge.list_mail("InboxAll", 100, 720, max_chars=1000)["items"]
    candidates = [
        r
        for r in rows
        if any(
            Path(a.get("FileName", "")).suffix.lower() in attachments.SUPPORTED
            for a in r.get("Attachments", [])
        )
    ]
    accepted = None
    for row in candidates[:3]:
        records = attachments.collect(row, rag.db_path().parent)
        if any(r["status"] in attachments.FATAL for r in records):
            continue
        text = " ".join(r.get("text", "") for r in records if r["status"] == "ok")
        terms = re.findall(r"[A-Za-z]{6,}", text)
        if not terms:
            continue
        row["_AttachmentRecords"] = records
        rag.initialize()
        with rag.connect(False) as c:
            rag.ingest([row], c)
        embedded = rag.embed_pending(8)
        result = rag.search(terms[0], mode="hybrid")
        assert result["items"]
        assert result["mode"] == "hybrid" and result["fallback"] is None
        accepted = {
            "extracted": sum(r["status"] == "ok" for r in records),
            "retrieval_hits": len(result["items"]),
            "retrieval_mode": result["mode"],
            "embedded": embedded["embedded"],
        }
        break
    if accepted is None:
        raise RuntimeError(
            "No successfully extractable live attachment found in the bounded sample"
        )
    assert not list((rag.db_path().parent / "attachment-tmp").iterdir())
    print(
        json.dumps(
            dict(
                accepted,
                temporary_files_remaining=0,
                wsl_invocations=0,
                processes=sorted(set(calls)),
            )
        )
    )
