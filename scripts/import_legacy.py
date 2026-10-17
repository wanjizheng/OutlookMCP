"""One-time consistent SQLite backup. Run on the source OS, not on every MCP startup."""

import argparse
import json
import os
import sqlite3
from pathlib import Path


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--source", type=Path, required=True)
    p.add_argument("--destination", type=Path, required=True)
    a = p.parse_args()
    dest = a.destination.resolve()
    if dest.exists():
        p.error("Destination exists; refusing to overwrite an index")
    dest.parent.mkdir(parents=True, exist_ok=True)
    partial = dest.with_suffix(".importing")
    if partial.exists():
        p.error("Previous import temporary file exists; inspect before retrying")
    src = sqlite3.connect(
        a.source.resolve().as_uri() + "?mode=ro", uri=True, timeout=30
    )
    target = sqlite3.connect(partial)
    try:
        src.backup(target, pages=1024, sleep=0.1)
        if target.execute("PRAGMA quick_check").fetchone()[0] != "ok":
            raise RuntimeError("Imported database check failed")
        count = target.execute("SELECT count(*) FROM emails").fetchone()[0]
        target.execute("PRAGMA journal_mode=DELETE")
        target.close()
        src.close()
        os.replace(partial, dest)
        print(json.dumps({"ok": True, "emails": count, "bytes": dest.stat().st_size}))
    finally:
        target.close()
        src.close()


if __name__ == "__main__":
    main()
