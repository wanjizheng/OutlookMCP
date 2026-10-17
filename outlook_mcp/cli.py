"""Administrative commands are intentionally separate from the read-only MCP."""

import argparse
import json
import os
from contextlib import contextmanager

from . import bridge, rag


@contextmanager
def writer_lock():
    import msvcrt

    rag.db_path().parent.mkdir(parents=True, exist_ok=True)
    with (rag.db_path().parent / "sync.lock").open("a+b") as lock:
        lock.seek(0)
        if os.fstat(lock.fileno()).st_size == 0:
            lock.write(b"0")
            lock.flush()
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise RuntimeError(
                "Another native sync or embedding job is active"
            ) from exc
        try:
            yield
        finally:
            lock.seek(0)
            msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)


def sync(folders, hours, pages, resume=False):
    with writer_lock():
        return _sync(folders, hours, pages, resume)


def _sync(folders, hours, pages, resume=False):
    if resume and hours:
        raise ValueError(
            "Use --hours 0 with --resume for a stable full backfill window"
        )
    rag.initialize()
    counts = {}
    with rag.connect(False) as c:
        for folder in folders:
            checkpoint = "windows_cursor:" + folder.lower()
            saved = c.execute(
                "SELECT value FROM meta WHERE key=?", (checkpoint,)
            ).fetchone()
            cursor = saved[0] if resume and saved else ""
            total = 0
            for _ in range(pages):
                page = bridge.list_mail(folder, 100, hours, cursor, max_chars=100000)
                with c:
                    rag.ingest(page["items"], c)
                    if page["next_cursor"] and hours == 0:
                        c.execute(
                            "INSERT OR REPLACE INTO meta VALUES(?,?)",
                            (checkpoint, page["next_cursor"]),
                        )
                    elif hours == 0:
                        c.execute("DELETE FROM meta WHERE key=?", (checkpoint,))
                total += len(page["items"])
                if not page["next_cursor"]:
                    break
                if cursor == page["next_cursor"]:
                    raise RuntimeError("Paging did not advance")
                cursor = page["next_cursor"]
            counts[folder] = {"fetched": total, "has_more": bool(page["next_cursor"])}
    return counts


def main():
    p = argparse.ArgumentParser(description="Windows OutlookMCP local administration")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("status")
    sub.add_parser("doctor")
    e = sub.add_parser("embed-pending")
    e.add_argument("--limit", type=int, default=200)
    s = sub.add_parser("sync")
    s.add_argument("--folders", default="InboxAll,Sent")
    s.add_argument("--hours", type=int, default=72)
    s.add_argument("--pages", type=int, default=10)
    s.add_argument(
        "--resume",
        action="store_true",
        help="Resume an all-time backfill from its committed page",
    )
    q = sub.add_parser("search")
    q.add_argument("query")
    q.add_argument("--mode", default="hybrid", choices=["hybrid", "bm25", "vector"])
    a = p.parse_args()
    if a.cmd == "status":
        out = rag.status()
    elif a.cmd == "doctor":
        out = {
            "windows_native": os.name == "nt",
            "wsl_required": False,
            "outlook": bridge.invoke("status"),
            "rag": rag.status(),
            "embedding_model": (bridge.ROOT / "models/embedding.gguf").exists(),
        }
    elif a.cmd == "sync":
        if not 1 <= a.pages <= 10000 or not 0 <= a.hours <= 876000:
            p.error("Invalid sync limits")
        out = sync(a.folders.split(","), a.hours, a.pages, a.resume)
    elif a.cmd == "embed-pending":
        with writer_lock():
            out = rag.embed_pending(a.limit)
    else:
        out = rag.search(a.query, mode=a.mode)
    print(json.dumps(out, ensure_ascii=False))


if __name__ == "__main__":
    main()
