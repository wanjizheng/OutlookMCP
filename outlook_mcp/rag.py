"""Local retrieval over the existing OpenClaw SQLite schema, without Linux runtime imports."""

import datetime as dt
import hashlib
import json
import os
import re
import sqlite3
import subprocess
from contextlib import contextmanager
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def db_path():
    return (
        Path(os.environ.get("OUTLOOK_MCP_DATA_DIR", str(ROOT / "data")))
        / "email_rag.db"
    )


@contextmanager
def connect(readonly=True, path=None):
    p = Path(path or db_path())
    if readonly:
        c = sqlite3.connect(p.resolve().as_uri() + "?mode=ro", uri=True, timeout=5)
        c.execute("PRAGMA query_only=ON")
    else:
        p.parent.mkdir(parents=True, exist_ok=True)
        c = sqlite3.connect(p.resolve().as_uri(), uri=True, timeout=30)
        c.execute("PRAGMA journal_mode=WAL")
    c.row_factory = sqlite3.Row
    try:
        with c:
            yield c
    finally:
        c.close()


def initialize(path=None):
    with connect(False, path) as c:
        c.executescript("""CREATE TABLE IF NOT EXISTS emails(entry_id TEXT PRIMARY KEY,stable_key TEXT,subject TEXT,sender TEXT,sender_email TEXT,received_time TEXT,body TEXT,updated_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY,value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS email_chunks(chunk_id INTEGER PRIMARY KEY AUTOINCREMENT,entry_id TEXT NOT NULL,chunk_index INTEGER NOT NULL,subject TEXT,sender TEXT,received_time TEXT,chunk_text TEXT NOT NULL,text_hash TEXT NOT NULL,embedding_json TEXT,embedding_norm REAL,updated_at TEXT NOT NULL,UNIQUE(entry_id,chunk_index));
CREATE TABLE IF NOT EXISTS email_chunk_vectors(chunk_id INTEGER PRIMARY KEY,dimensions INTEGER NOT NULL,embedding_f32 BLOB NOT NULL,updated_at TEXT NOT NULL);
CREATE VIRTUAL TABLE IF NOT EXISTS email_chunks_fts USING fts5(entry_id UNINDEXED,chunk_id UNINDEXED,subject,sender,received_time UNINDEXED,chunk_text,tokenize='unicode61');""")


def status():
    p = db_path()
    if not p.exists():
        return {
            "available": False,
            "reason": "Run the local import or sync command first",
            "wsl_required": False,
        }
    with connect() as c:
        return {
            "available": True,
            "emails": c.execute("SELECT count(*) FROM emails").fetchone()[0],
            "chunks": c.execute("SELECT count(*) FROM email_chunks").fetchone()[0],
            "last_sync": dict(
                c.execute(
                    "SELECT key,value FROM meta WHERE key IN ('last_sync_at','windows_last_sync_at')"
                )
            ),
            "retrieval": "local FTS5; hybrid when GGUF model is installed",
            "wsl_required": False,
        }


def embed(texts):
    model = ROOT / "models/embedding.gguf"
    if not model.exists():
        raise RuntimeError("Local embedding model is not installed")
    env = os.environ.copy()
    env["NODE_LLAMA_CPP_GPU"] = env.get("NODE_LLAMA_CPP_GPU", "false")
    proc = subprocess.run(
        [
            "node",
            str(ROOT / "scripts/embed_local.mjs"),
            "--one-shot",
            "--model-path",
            str(model),
        ],
        input=json.dumps({"texts": [t[:2000] for t in texts]}),
        text=True,
        encoding="utf-8",
        capture_output=True,
        timeout=45,
        env=env,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    for line in reversed(proc.stdout.splitlines()):
        try:
            data = json.loads(line)
            if "vectors" in data:
                return data["vectors"]
        except ValueError:
            pass
    raise RuntimeError("Local embedding failed")


def search(query, top=8, mode="hybrid"):
    if (
        not query.strip()
        or len(query) > 2000
        or not 1 <= top <= 30
        or mode not in ("bm25", "hybrid", "vector")
    ):
        raise ValueError("Invalid search request")
    tokens = re.findall(r"\w+", query, flags=re.UNICODE)
    if not tokens:
        return {"items": [], "mode": "bm25"}
    expression = " OR ".join('"' + t.replace('"', '""') + '"' for t in tokens)
    with connect() as c:
        lexical = [
            dict(r)
            for r in c.execute(
                "SELECT entry_id,chunk_id,subject,sender,received_time,chunk_text AS snippet,bm25(email_chunks_fts) AS score FROM email_chunks_fts WHERE email_chunks_fts MATCH ? ORDER BY score LIMIT ?",
                (expression, top * 4),
            )
        ]
        vector = []
        fallback = None
        if mode != "bm25":
            try:
                import numpy as np

                q = np.asarray(embed([query])[0], dtype=np.float32)
                q = q / max(float(np.linalg.norm(q)), 1e-12)
                best = []
                cur = c.execute(
                    "SELECT v.chunk_id,v.embedding_f32 FROM email_chunk_vectors v JOIN email_chunks c ON c.chunk_id=v.chunk_id WHERE v.dimensions=?",
                    (len(q),),
                )
                while batch := cur.fetchmany(512):
                    matrix = np.stack(
                        [np.frombuffer(r["embedding_f32"], dtype="<f4") for r in batch]
                    )
                    scores = (matrix @ q) / np.maximum(
                        np.linalg.norm(matrix, axis=1), 1e-12
                    )
                    best.extend(
                        (float(score), int(row["chunk_id"]))
                        for score, row in zip(scores, batch)
                    )
                    best = sorted(best, reverse=True)[: top * 4]
                for score, cid in best:
                    r = c.execute(
                        "SELECT entry_id,chunk_id,subject,sender,received_time,chunk_text AS snippet FROM email_chunks WHERE chunk_id=?",
                        (cid,),
                    ).fetchone()
                    if r:
                        vector.append(dict(r, score=score))
            except (RuntimeError, subprocess.TimeoutExpired, OSError) as exc:
                if mode == "vector":
                    raise RuntimeError("Vector search unavailable") from exc
                fallback = type(exc).__name__
        rankings = {}
        records = {}
        for group in (
            [lexical]
            if mode == "bm25"
            else [vector]
            if mode == "vector"
            else [lexical, vector]
        ):
            for rank, row in enumerate(group, 1):
                cid = row["chunk_id"]
                records[cid] = row
                rankings[cid] = rankings.get(cid, 0) + 1 / (60 + rank)
        results = []
        seen = set()
        for cid in sorted(rankings, key=rankings.get, reverse=True):
            row = records[cid]
            if row["entry_id"] in seen:
                continue
            seen.add(row["entry_id"])
            row["snippet"] = row["snippet"][:1800]
            row["source"] = "local_index"
            results.append(row)
            if len(results) >= top:
                break
        return {
            "items": results,
            "mode": mode if vector or mode == "bm25" else "bm25",
            "fallback": fallback,
            "freshness": status()["last_sync"],
        }


def read_indexed(entry_id, max_chars=12000):
    if not 1 <= max_chars <= 100000:
        raise ValueError("Invalid body limit")
    with connect() as c:
        r = c.execute(
            "SELECT entry_id,subject,sender,received_time,body FROM emails WHERE entry_id=?",
            (entry_id,),
        ).fetchone()
        if not r:
            raise ValueError("Message not found in local index")
        out = dict(r)
        body = out["body"] or ""
        out.update(
            body=body[:max_chars], truncated=len(body) > max_chars, source="local_index"
        )
        return out


def ingest(rows, conn):
    """Reuse legacy chunk size/overlap and schema; refresh changed bodies atomically."""
    now = dt.datetime.now(dt.timezone.utc).isoformat()
    for r in rows:
        eid = r["EntryID"]
        body = r.get("Body", "")
        subject = r.get("Subject", "")
        sender = r.get("Sender", "")
        stamp = r.get("ReceivedTime", "")
        old = conn.execute(
            "SELECT body,subject FROM emails WHERE entry_id=?", (eid,)
        ).fetchone()
        if old and old["body"] == body and old["subject"] == subject:
            continue
        conn.execute(
            "DELETE FROM email_chunk_vectors WHERE chunk_id IN (SELECT chunk_id FROM email_chunks WHERE entry_id=?)",
            (eid,),
        )
        conn.execute("DELETE FROM email_chunks_fts WHERE entry_id=?", (eid,))
        conn.execute("DELETE FROM email_chunks WHERE entry_id=?", (eid,))
        conn.execute(
            "INSERT OR REPLACE INTO emails VALUES(?,?,?,?,?,?,?,?)",
            (
                eid,
                None,
                subject,
                sender,
                r.get("SenderEmailAddress", ""),
                stamp,
                body,
                now,
            ),
        )
        content = subject + "\n" + body
        for index, start in enumerate(range(0, len(content), 750)):
            text = content[start : start + 900]
            cid = conn.execute(
                "INSERT INTO email_chunks(entry_id,chunk_index,subject,sender,received_time,chunk_text,text_hash,updated_at) VALUES(?,?,?,?,?,?,?,?)",
                (
                    eid,
                    index,
                    subject,
                    sender,
                    stamp,
                    text,
                    hashlib.sha256(text.encode()).hexdigest(),
                    now,
                ),
            ).lastrowid
            conn.execute(
                "INSERT INTO email_chunks_fts VALUES(?,?,?,?,?,?)",
                (eid, cid, subject, sender, stamp, text),
            )
    conn.execute(
        "INSERT OR REPLACE INTO meta VALUES(?,?)", ("windows_last_sync_at", now)
    )


def embed_pending(limit=200):
    """Fill native-sync vectors locally; never called implicitly by MCP tools."""
    if not 1 <= limit <= 100000:
        raise ValueError("Invalid embedding limit")
    import numpy as np

    done = 0
    with connect(False) as c:
        pending = c.execute(
            "SELECT c.chunk_id,c.chunk_text FROM email_chunks c LEFT JOIN email_chunk_vectors v ON v.chunk_id=c.chunk_id WHERE v.chunk_id IS NULL ORDER BY c.chunk_id LIMIT ?",
            (limit,),
        ).fetchall()
        for start in range(0, len(pending), 8):
            batch = pending[start : start + 8]
            vectors = embed([r["chunk_text"] for r in batch])
            if len(vectors) != len(batch):
                raise RuntimeError("Embedding response size mismatch")
            with c:
                for row, vector in zip(batch, vectors):
                    v = np.asarray(vector, dtype="<f4")
                    c.execute(
                        "INSERT OR REPLACE INTO email_chunk_vectors VALUES(?,?,?,?)",
                        (
                            row["chunk_id"],
                            len(v),
                            v.tobytes(),
                            dt.datetime.now(dt.UTC).isoformat(),
                        ),
                    )
            done += len(batch)
    return {"embedded": done, "requested_limit": limit}
