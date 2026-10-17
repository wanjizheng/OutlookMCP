"""Convert a consistent legacy snapshot to the smaller native retrieval schema."""

import argparse
import json
import os
from datetime import UTC, datetime
from pathlib import Path

from outlook_mcp import rag


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--source", type=Path, required=True)
    p.add_argument("--destination", type=Path, required=True)
    a = p.parse_args()
    dest = a.destination.resolve()
    if dest.exists():
        p.error("Destination exists; refusing to overwrite")
    partial = dest.with_suffix(".building")
    if partial.exists():
        p.error("Temporary compact database already exists")
    rag.initialize(partial)
    with rag.connect(False, partial) as c:
        c.execute(
            "ATTACH DATABASE ? AS legacy", (a.source.resolve().as_uri() + "?mode=ro",)
        )
        with c:
            c.execute(
                "INSERT INTO emails SELECT entry_id,stable_key,subject,sender,sender_email,received_time,body,updated_at FROM legacy.emails"
            )
            c.execute(
                "INSERT INTO email_chunks SELECT chunk_id,entry_id,chunk_index,subject,sender,received_time,chunk_text,text_hash,NULL,NULL,? FROM legacy.email_chunks",
                (datetime.now(UTC).isoformat(),),
            )
            c.execute(
                "INSERT INTO email_chunk_vectors SELECT * FROM legacy.email_chunk_vectors"
            )
            c.execute(
                "DELETE FROM email_chunk_vectors WHERE chunk_id NOT IN (SELECT chunk_id FROM email_chunks)"
            )
            c.execute("INSERT INTO meta SELECT * FROM legacy.meta")
            c.execute(
                "INSERT INTO email_chunks_fts SELECT entry_id,chunk_id,subject,sender,received_time,chunk_text FROM email_chunks"
            )
        c.execute("DETACH DATABASE legacy")
        if c.execute("PRAGMA quick_check").fetchone()[0] != "ok":
            raise RuntimeError("Compact index integrity check failed")
        report = {
            "emails": c.execute("SELECT count(*) FROM emails").fetchone()[0],
            "chunks": c.execute("SELECT count(*) FROM email_chunks").fetchone()[0],
            "vectors": c.execute("SELECT count(*) FROM email_chunk_vectors").fetchone()[
                0
            ],
        }
        c.execute("PRAGMA wal_checkpoint(TRUNCATE)")
    os.replace(partial, dest)
    report.update(
        integrity="ok", bytes=dest.stat().st_size, source_bytes=a.source.stat().st_size
    )
    print(json.dumps(report))


if __name__ == "__main__":
    main()
