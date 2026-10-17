#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import fcntl
import heapq
import hashlib
import json
import math
import os
import re
import select
import signal
import shutil
import socket
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Callable

try:
    import numpy as np
except ImportError:  # pragma: no cover - exercised only on minimal installs
    np = None

DEFAULT_DB = Path.home() / ".openclaw" / "workspace" / "data" / "email_rag" / "email_rag.db"
OUTLOOK_READ_CMD = str(Path.home() / "bin" / "outlook-mail-read")
OFFICE_DOC_IO_CMD = str(Path.home() / "bin" / "office-doc-io")

# --- Outlook daemon (hidden PowerShell host) ---
# Use the hidden daemon instead of spawning a PowerShell window per Outlook
# call. See outlook_reader/scripts/{start_daemon.vbs, outlook_daemon.ps1,
# outlook_pipe_client.ps1}.
_OUTLOOK_READER_DIR = Path.home() / ".openclaw" / "workspace" / "skills" / "outlook_reader" / "scripts"
OUTLOOK_PIPE_CLIENT_PS1 = _OUTLOOK_READER_DIR / "outlook_pipe_client.ps1"
OUTLOOK_DAEMON_PS1 = _OUTLOOK_READER_DIR / "outlook_daemon.ps1"
START_DAEMON_VBS = _OUTLOOK_READER_DIR / "start_daemon.vbs"
OUTLOOK_PIPE_NAME = "OpenClawOutlookDaemon"
CSCRIPT_EXE = "/mnt/c/Windows/System32/cscript.exe"

# On Windows, pass this to subprocess.run/Popen creationflags= to suppress
# the flash of a console window for the child process. 0 on non-Windows.
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

DEFAULT_EMBED_BASE_URL = "http://127.0.0.1:8000/v1"
DEFAULT_EMBED_MODEL = "text-embedding-3-large"
DEFAULT_EMBED_API_KEY = "local-gateway-placeholder"
DEFAULT_EMBED_BATCH = 1
DEFAULT_EMBED_TIMEOUT = 120
DEFAULT_EMBED_BACKEND = "local"
LOCAL_EMBED_SCRIPT = Path(__file__).parent / "embed_local.mjs"
LOCAL_EMBED_TIMEOUT_OVERHEAD = 60   # model load overhead (seconds)
LOCAL_EMBED_TIMEOUT_PER_TEXT = 3    # upper-bound per text (seconds)
LOCAL_EMBED_BATCH_TIMEOUT_OVERHEAD = float(
    os.environ.get("EMAIL_RAG_BATCH_TIMEOUT_OVERHEAD", "15")
)
LOCAL_EMBED_BATCH_TIMEOUT_PER_TEXT = float(
    os.environ.get("EMAIL_RAG_BATCH_TIMEOUT_PER_TEXT", "0.5")
)

DEFAULT_DAILY_COUNT = 50
DEFAULT_DAILY_FOLDER = "InboxAll"
DEFAULT_DAILY_HOURS = 24
DEFAULT_DAILY_MODE = "all"
DEFAULT_DAILY_MAX_CHARS = 4000
DEFAULT_DAILY_MAX_WINDOWS = 30

DEFAULT_FULL_COUNT_PER_FOLDER = 50000
DEFAULT_FULL_FOLDERS = ["InboxAll", "Sent", "Drafts"]
DEFAULT_FULL_MODE = "all"
DEFAULT_FULL_MAX_CHARS = 4000
DEFAULT_FULL_SAFE_FOLDERS = ["InboxAll", "Sent"]
# 0 means no time filter. A "full" default must not silently stop at five
# years and omit older mail.
DEFAULT_FULL_SAFE_HOURS_WINDOWS = [0]
DEFAULT_FULL_SAFE_COUNT_PER_STEP = 50000
DEFAULT_PAGED_READ_SIZE = 200
DEFAULT_FULL_SAFE_PAGE_SIZE = 100
MAX_OUTLOOK_READ_COUNT = 500
PAGE_SLEEP_SEC = 1  # seconds to sleep between paged Outlook reads
FULL_SAFE_UNIT_PREFIX = "email-rag-full-safe"
FULL_SAFE_DATA_DIR = Path.home() / ".openclaw" / "workspace" / "data" / "email_rag"
FULL_SAFE_LOG = FULL_SAFE_DATA_DIR / "full-safe.log"
FULL_SAFE_LOCK = FULL_SAFE_DATA_DIR / "full-safe.lock"
FULL_SAFE_PAGE_KEYS = (
    "full_safe.page.signature",
    "full_safe.page.step_index",
    "full_safe.page.before_iso",
    "full_safe.page.processed",
    "full_safe.page.fetched",
    "full_safe.page.synced",
    "full_safe.page.skipped_existing",
    "full_safe.page.pages",
    "full_safe.page.updated_at",
)

DEFAULT_ATTACHMENT_MODE = "export"
DEFAULT_ATTACHMENT_EXPORT_DIR = Path.home() / ".openclaw" / "workspace" / "data" / "email_rag" / "attachments"
DEFAULT_ATTACHMENT_TEST_DIR = Path.home() / ".openclaw" / "workspace" / "data" / "email_rag" / "attachments_test"
DEFAULT_ATTACHMENT_MAX_PER_EMAIL = 6
DEFAULT_ATTACHMENT_MAX_CHARS = 12000
DEFAULT_ATTACHMENT_CLEANUP = "all"
DEFAULT_ESTIMATE_FOLDERS = "InboxAll,Sent"


class SyncInterrupted(KeyboardInterrupt):
    """Raised when a long-running sync receives SIGTERM/SIGINT."""


def make_open_command(entry_id: str) -> str:
    escaped = entry_id.replace('"', '\\"')
    return f'{Path.home() / "bin" / "outlook-mail-open"} --entry-id "{escaped}"'


# ── GPU contention probe (WSL2 → Windows) ─────────────────────────────────
#
# The local embed backend uses CUDA via node-llama-cpp. When a Windows-side
# process (game, 3D app, LM Studio, etc.) is using the same GPU, embed calls
# get queued/OOM and slow down 2-10×. WSL2's nvidia-smi cannot see Windows
# PIDs, so we have to probe Windows-side GPU engine utilization directly via
# PowerShell `Get-Counter`. Returns the list of contending PIDs so we can
# name-and-shame them before the user invests an hour in a slow sync.

_GPU_PROBE_THRESHOLD_PCT = 25.0  # only flag processes above this 3D-engine util
_GPU_PROBE_TIMEOUT_SEC = 15      # hard cap so probe never blocks sync


def _run_ps_quick(ps_script: str, max_lines: int = 32, timeout_sec: float = _GPU_PROBE_TIMEOUT_SEC) -> str:
    """Run a short PowerShell script and return stdout, killed as soon as we have enough lines.

    subprocess.run() with timeout doesn't work for `Get-Counter`: the cmdlet emits
    its CSV quickly but the powershell.exe process often lingers (loading perf
    counters, etc.), and `communicate()` waits for exit → false TimeoutExpired.
    So we stream stdout via Popen + readline, kill the process once we have
    `max_lines` data lines (or 1 line for our use case here), and never wait
    longer than `timeout_sec`.
    """
    try:
        proc = subprocess.Popen(
            ["powershell.exe", "-NoProfile", "-Command", ps_script],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            errors="replace",  # MainWindowTitle may contain non-UTF-8 (e.g. GBK 中文)
            bufsize=1,
            creationflags=CREATE_NO_WINDOW,  # never flash a CMD window
        )
    except (FileNotFoundError, OSError):
        return ""

    deadline = time.monotonic() + timeout_sec
    collected: list[str] = []
    header: str | None = None
    data_lines = 0
    try:
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            if data_lines >= max_lines:
                break
            try:
                ready, _, _ = select.select([proc.stdout], [], [], remaining)
            except (OSError, ValueError):
                break
            if not ready:
                break
            line = proc.stdout.readline()
            if not line:
                # EOF or closed pipe
                break
            stripped = line.rstrip()
            if header is None:
                header = stripped
                continue
            collected.append(stripped)
            data_lines += 1
    finally:
        try:
            proc.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=2)
        except subprocess.TimeoutExpired:
            try:
                proc.kill()
            except Exception:
                pass

    if header is None:
        return ""
    return header + "\n" + "\n".join(collected)


def _probe_gpu_contention(threshold_pct: float = _GPU_PROBE_THRESHOLD_PCT) -> list[dict[str, Any]]:
    """Probe Windows GPU engine utilization. Returns list of {pid, util, name}.

    Returns [] on any failure (not WSL2, no powershell.exe, timeout, etc.) so
    it is always safe to call. Only PIDs whose 3D engine util >= threshold_pct
    are returned.
    """
    if sys.platform != "linux":
        return []

    ps_script = (
        "Get-Counter '\\GPU Engine(*)\\Utilization Percentage' -SampleInterval 1 -MaxSamples 1 "
        "-ErrorAction SilentlyContinue | Select-Object -ExpandProperty CounterSamples "
        "| Where-Object { $_.CookedValue -ge " + f"{threshold_pct}" + " } "
        "| Select-Object InstanceName, CookedValue "
        "| Sort-Object CookedValue -Descending "
        "| ConvertTo-Csv -NoTypeInformation"
    )
    stdout = _run_ps_quick(ps_script, max_lines=8, timeout_sec=_GPU_PROBE_TIMEOUT_SEC)
    if not stdout.strip():
        return []

    contenders: list[dict[str, Any]] = []
    for line in stdout.splitlines()[1:]:  # skip CSV header
        line = line.strip()
        if not line or '","' not in line:
            continue
        try:
            # CSV row: "pid_<pid>_luid_<luid>_phys_<phys>_eng_<eng>_engtype_<type>","<value>"
            name_part, value_part = line.rsplit('","', 1)
            name_part = name_part.lstrip('"')
            value_part = value_part.rstrip().rstrip('"')
            util = float(value_part)
            if "_luid_" not in name_part:
                continue
            pid_str = name_part.split("_luid_")[0].replace("pid_", "")
            pid = int(pid_str)
            contenders.append({"pid": pid, "util": util})
        except (ValueError, IndexError):
            continue

    return contenders


def _resolve_pid_names(pids: list[int]) -> dict[int, str]:
    """Map PIDs to 'ProcessName | MainWindowTitle' via Windows Get-Process.

    Returns {pid: '<name>'}; empty dict on any failure.
    """
    if not pids:
        return {}
    pid_list = ",".join(str(p) for p in pids)
    # Force UTF-8 output so non-ASCII MainWindowTitle (e.g. 中文) isn't GBK-mojibake on read.
    ps_script = (
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
        f"Get-Process | Where-Object {{ $_.Id -in @({pid_list}) }} "
        "| Select-Object Id, ProcessName, MainWindowTitle "
        "| ForEach-Object { \"{0}|{1}|{2}\" -f $_.Id, $_.ProcessName, $_.MainWindowTitle }"
    )
    stdout = _run_ps_quick(ps_script, max_lines=len(pids) + 4, timeout_sec=_GPU_PROBE_TIMEOUT_SEC)
    if not stdout.strip():
        return {}

    out: dict[int, str] = {}
    for line in stdout.splitlines():
        line = line.strip()
        if "|" not in line:
            continue
        parts = line.split("|", 2)
        if len(parts) < 2:
            continue
        try:
            pid = int(parts[0])
        except ValueError:
            continue
        name = parts[1]
        title = parts[2] if len(parts) > 2 else ""
        out[pid] = f"{name}" + (f" — {title}" if title else "")
    return out


def _warn_gpu_contention(contenders: list[dict[str, Any]]) -> None:
    """Print a loud, actionable warning when GPU contention is detected.

    No-op if contenders is empty.
    """
    if not contenders:
        return

    names = _resolve_pid_names([c["pid"] for c in contenders])

    print("", flush=True)
    print("!" * 72, flush=True)
    print("!  GPU CONTENTION DETECTED — local embedding will be SLOW (2-10x)", flush=True)
    print("!" * 72, flush=True)
    for c in contenders[:5]:
        pid = c["pid"]
        util = c["util"]
        label = names.get(pid, "(unknown)")
        print(f"!    PID {pid:>6}  {util:>5.1f}% 3D  {label}", flush=True)
    print("!", flush=True)
    print("!  The local embed backend (CUDA) shares this GPU with the process(es)", flush=True)
    print("!  above. They were not visible to WSL2's nvidia-smi — that's why they", flush=True)
    print("!  were missed. To fix:", flush=True)
    print("!    1. Close or suspend the GPU consumer (game, 3D app, LM Studio, …)", flush=True)
    print("!    2. Or rerun with --no-embed (BM25-only sync; loses semantic search)", flush=True)
    print("!    3. Or wait until GPU is free, then rerun", flush=True)
    print("!  Disable this probe with --no-gpu-probe.", flush=True)
    print("!" * 72, flush=True)
    print("", flush=True)


def ensure_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        db_path.parent.chmod(0o700)
    except OSError:
        pass
    conn = sqlite3.connect(db_path)
    try:
        db_path.chmod(0o600)
    except OSError:
        pass
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    for runtime_path in (db_path, Path(f"{db_path}-wal"), Path(f"{db_path}-shm")):
        if runtime_path.exists():
            try:
                runtime_path.chmod(0o600)
            except OSError:
                pass

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS emails (
            entry_id TEXT PRIMARY KEY,
            stable_key TEXT,
            subject TEXT,
            sender TEXT,
            sender_email TEXT,
            received_time TEXT,
            body TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )
    # Lightweight migration for existing DBs created before stable_key.
    email_cols = [r[1] for r in conn.execute("PRAGMA table_info(emails)").fetchall()]
    if "stable_key" not in email_cols:
        conn.execute("ALTER TABLE emails ADD COLUMN stable_key TEXT")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS embedding_cache (
            text_hash TEXT PRIMARY KEY,
            embedding_json TEXT NOT NULL,
            embedding_norm REAL NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS email_chunks (
            chunk_id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id TEXT NOT NULL,
            chunk_index INTEGER NOT NULL,
            subject TEXT,
            sender TEXT,
            received_time TEXT,
            chunk_text TEXT NOT NULL,
            text_hash TEXT NOT NULL,
            embedding_json TEXT,
            embedding_norm REAL,
            updated_at TEXT NOT NULL,
            UNIQUE(entry_id, chunk_index)
        )
        """
    )
    # Keep a compact float32 copy of every embedding for vector search.
    # Parsing the JSON representation for all chunks on every query used to
    # dominate search latency (10-30s for ~37k chunks).  This side table keeps
    # the existing JSON schema/API compatible while making cosine search a
    # bounded binary read plus a vectorized matrix multiply.
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS email_chunk_vectors (
            chunk_id INTEGER PRIMARY KEY,
            dimensions INTEGER NOT NULL,
            embedding_f32 BLOB NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TRIGGER IF NOT EXISTS email_chunks_delete_vector
        AFTER DELETE ON email_chunks
        BEGIN
            DELETE FROM email_chunk_vectors WHERE chunk_id = OLD.chunk_id;
        END
        """
    )
    conn.execute(
        """
        CREATE VIRTUAL TABLE IF NOT EXISTS email_chunks_fts
        USING fts5(
            entry_id UNINDEXED,
            chunk_id UNINDEXED,
            subject,
            sender,
            received_time UNINDEXED,
            chunk_text,
            tokenize='unicode61'
        )
        """
    )

    cols = [r[1] for r in conn.execute("PRAGMA table_info(email_chunks_fts)").fetchall()]
    if cols and "chunk_id" not in cols:
        conn.execute("DROP TABLE email_chunks_fts")
        conn.execute(
            """
            CREATE VIRTUAL TABLE email_chunks_fts
            USING fts5(
                entry_id UNINDEXED,
                chunk_id UNINDEXED,
                subject,
                sender,
                received_time UNINDEXED,
                chunk_text,
                tokenize='unicode61'
            )
            """
        )

    conn.execute("CREATE INDEX IF NOT EXISTS idx_email_chunks_entry ON email_chunks(entry_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_email_chunks_hash ON email_chunks(text_hash)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_emails_stable_key ON emails(stable_key)")
    return conn


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def parse_utc_timestamp(raw: str | None) -> dt.datetime | None:
    if not raw:
        return None
    val = raw.strip()
    if not val:
        return None
    if val.endswith("Z"):
        val = val[:-1] + "+00:00"
    try:
        parsed = dt.datetime.fromisoformat(val)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=dt.timezone.utc)
    return parsed.astimezone(dt.timezone.utc)


def set_meta(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute(
        "INSERT INTO meta(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value),
    )


def get_meta(conn: sqlite3.Connection, key: str) -> str | None:
    row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row[0] if row else None


def _safe_int(raw: str | None, default: int = 0) -> int:
    try:
        return int(raw or default)
    except (TypeError, ValueError):
        return default


def _email_rag_sync_pid_alive(pid: int) -> bool:
    """Return True only for a live email-rag sync worker, not a reused PID."""
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
        cmdline = Path(f"/proc/{pid}/cmdline").read_bytes().replace(b"\0", b" ")
    except (FileNotFoundError, ProcessLookupError, PermissionError, OSError):
        return False
    return b"email_rag.py" in cmdline and any(
        marker in cmdline
        for marker in (b"sync-full-safe", b"sync-full", b"sync-default", b" sync ")
    )


def _systemd_user_unit_active(unit: str) -> bool:
    try:
        result = subprocess.run(
            ["systemctl", "--user", "is-active", "--quiet", unit],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=5,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        return False
    return result.returncode == 0


def _clear_full_safe_page_checkpoint(conn: sqlite3.Connection) -> None:
    conn.executemany("DELETE FROM meta WHERE key=?", [(key,) for key in FULL_SAFE_PAGE_KEYS])


def _load_full_safe_page_checkpoint(
    conn: sqlite3.Connection,
    *,
    signature: str,
    step_index: int,
) -> dict[str, Any] | None:
    if get_meta(conn, "full_safe.page.signature") != signature:
        return None
    if _safe_int(get_meta(conn, "full_safe.page.step_index"), -1) != step_index:
        return None
    before_iso = get_meta(conn, "full_safe.page.before_iso") or None
    processed = max(0, _safe_int(get_meta(conn, "full_safe.page.processed")))
    if processed <= 0 or not before_iso or not before_iso.startswith("v2:"):
        return None
    return {
        "before_iso": before_iso,
        "processed": processed,
        "fetched": max(0, _safe_int(get_meta(conn, "full_safe.page.fetched"))),
        "synced": max(0, _safe_int(get_meta(conn, "full_safe.page.synced"))),
        "skipped_existing": max(
            0, _safe_int(get_meta(conn, "full_safe.page.skipped_existing"))
        ),
        "pages": max(0, _safe_int(get_meta(conn, "full_safe.page.pages"))),
    }


def _write_full_safe_page_checkpoint(
    conn: sqlite3.Connection,
    *,
    signature: str,
    step_index: int,
    state: dict[str, Any],
) -> None:
    set_meta(conn, "full_safe.page.signature", signature)
    set_meta(conn, "full_safe.page.step_index", str(step_index))
    set_meta(conn, "full_safe.page.before_iso", str(state.get("before_iso") or ""))
    set_meta(conn, "full_safe.page.processed", str(max(0, int(state.get("processed", 0)))))
    set_meta(conn, "full_safe.page.fetched", str(max(0, int(state.get("fetched", 0)))))
    set_meta(conn, "full_safe.page.synced", str(max(0, int(state.get("synced", 0)))))
    set_meta(
        conn,
        "full_safe.page.skipped_existing",
        str(max(0, int(state.get("skipped_existing", 0)))),
    )
    set_meta(conn, "full_safe.page.pages", str(max(0, int(state.get("pages", 0)))))
    set_meta(conn, "full_safe.page.updated_at", now_iso())
    set_meta(conn, "full_safe.updated_at", now_iso())
    conn.commit()


def _mark_sync_interrupted(db_path: Path, reason: str) -> None:
    conn = ensure_db(db_path)
    now = now_iso()
    set_meta(conn, "progress_status", "interrupted")
    set_meta(conn, "progress_error", reason)
    set_meta(conn, "progress_pid", "")
    set_meta(conn, "progress_current_folder", "")
    set_meta(conn, "progress_finished_at", now)
    set_meta(conn, "progress_updated_at", now)
    set_meta(conn, "full_safe.status", "interrupted")
    set_meta(conn, "full_safe.updated_at", now)
    conn.commit()
    conn.close()


def _acquire_full_safe_lock() -> int | None:
    FULL_SAFE_DATA_DIR.mkdir(parents=True, exist_ok=True)
    try:
        FULL_SAFE_DATA_DIR.chmod(0o700)
    except OSError:
        pass
    fd = os.open(FULL_SAFE_LOCK, os.O_RDWR | os.O_CREAT, 0o600)
    try:
        fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        os.close(fd)
        return None
    os.ftruncate(fd, 0)
    os.write(fd, f"{os.getpid()}\n".encode("ascii"))
    os.fsync(fd)
    return fd


def _release_full_safe_lock(fd: int) -> None:
    try:
        fcntl.flock(fd, fcntl.LOCK_UN)
    finally:
        os.close(fd)


def _execute_sync_locked(
    args: argparse.Namespace,
    plans: list[dict[str, Any]],
    **kwargs: Any,
) -> int:
    """Serialize ordinary syncs against the durable full-safe worker."""
    lock_fd = _acquire_full_safe_lock()
    if lock_fd is None:
        print(
            json.dumps(
                {
                    "ok": False,
                    "busy": True,
                    "error": "Another email RAG sync is already running.",
                    "hint": "Use `email-rag progress` and retry after it finishes.",
                },
                ensure_ascii=False,
            )
        )
        return 3
    try:
        return _execute_sync(args, plans, **kwargs)
    finally:
        _release_full_safe_lock(lock_fd)


def chunk_text(text: str, chunk_size: int = 900, overlap: int = 150) -> list[str]:
    text = re.sub(r"\s+", " ", (text or "")).strip()
    if not text:
        return []
    out: list[str] = []
    start = 0
    n = len(text)
    while start < n:
        end = min(start + chunk_size, n)
        out.append(text[start:end])
        if end >= n:
            break
        start = max(end - overlap, start + 1)
    return out


def _parse_outlook_payload(raw: str) -> list[dict[str, Any]]:
    if not raw:
        raise RuntimeError("outlook-mail-read returned no output")
    # PowerShell/bash may prepend noise (e.g. "read failed 9: Bad file descriptor").
    # Strip everything before the first JSON delimiter.
    json_start = -1
    for i, ch in enumerate(raw):
        if ch in ("{", "["):
            json_start = i
            break
    if json_start > 0:
        raw = raw[json_start:]
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Cannot parse outlook-mail-read output: {exc}. output={raw[:500]}") from exc

    if isinstance(payload, dict):
        if payload.get("error"):
            raise RuntimeError(str(payload.get("error")))
        if "EntryID" in payload:
            return [payload]
        raise RuntimeError(f"Unexpected payload object shape: {str(payload)[:300]}")

    if not isinstance(payload, list):
        raise RuntimeError(f"Unexpected payload type: {type(payload).__name__}")
    return payload


# --- Outlook daemon transport -----------------------------------------------
# Routes all Outlook COM calls through a single hidden PowerShell daemon
# (outlook_daemon.ps1) over the named pipe `OpenClawOutlookDaemon`. The daemon
# is launched via start_daemon.vbs which uses WindowStyle=0 (SW_HIDE), so no
# console window flashes even though every call from Python spawns a
# powershell.exe (the pipe client).
#
# Why this matters: previously each Outlook read spawned a visible PowerShell
# window from `~/bin/outlook-mail-read`. Paged sync hits Outlook dozens of
# times per run, so the windows stacked up.

_pipe_proc: subprocess.Popen | None = None
_pipe_lock: threading.Lock | None = None
_pipe_request_lock = threading.Lock()


def _get_pipe_lock():
    global _pipe_lock
    if _pipe_lock is None:
        _pipe_lock = threading.Lock()
    return _pipe_lock


def _pipe_alive() -> bool:
    return _pipe_proc is not None and _pipe_proc.poll() is None


def _daemon_ping(timeout_sec: float = 4.0) -> bool:
    r"""Best-effort check that the Outlook daemon is up.

    The daemon's `ping` handler is intentionally NOT used here — it triggers a
    `Get-OutlookSession` call that takes ~30s when the Outlook COM attach
    fails, which would always exceed our timeout. Instead we just check that
    the named pipe is bound. The daemon binds the pipe immediately after
    startup (before the first request is processed), so presence of
    `\\.\pipe\OpenClawOutlookDaemon` is a strong liveness signal.
    """
    if not OUTLOOK_PIPE_CLIENT_PS1.exists():
        return False
    try:
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-Command", "[bool]([System.IO.Directory]::GetFiles('\\\\.\\pipe\\') "
             "| Where-Object { $_ -like '*OpenClawOutlookDaemon' })"],
            capture_output=True,
            text=True,
            timeout=timeout_sec,
            creationflags=CREATE_NO_WINDOW,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return False
    return proc.stdout.strip().lower() == "true"


def _is_elevated() -> bool:
    """True if the current WSL→Windows PowerShell is running as Administrator.

    Elevated PowerShell can't talk to non-elevated Outlook COM (Windows blocks
    the cross-integrity-level call with CO_E_SERVER_EXEC_FAILURE after a ~30s
    timeout). The bash `outlook-mail-read` wrapper works around this by
    routing through `schtasks /RL LIMITED` to drop privileges; we do the same
    when launching the daemon. See outlook_reader/scripts/_common.sh for the
    historical rationale.
    """
    try:
        out = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
             "[bool]([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)"],
            capture_output=True, text=True, timeout=10,
            creationflags=CREATE_NO_WINDOW,
        ).stdout.strip().lower()
        return out == "true"
    except Exception:
        return False


SCHTASKS_EXE = "/mnt/c/Windows/System32/schtasks.exe"
DAEMON_STDERR_LOG_WIN = r"C:\Users\wanjizheng\AppData\Local\Temp\outlook-daemon-stderr.log"
DAEMON_STDERR_LOG_WSL = Path("/mnt/c/Users/wanjizheng/AppData/Local/Temp/outlook-daemon-stderr.log")
DAEMON_PID_FILE_WIN = r"C:\Users\wanjizheng\AppData\Local\Temp\outlook-daemon.pid"


def _kill_stale_outlook_state(*, kill_procs: bool = True, rotate_log: bool = True) -> None:
    """Kill leftover daemon / pipe-client processes and clean up schtasks.

    Without this, a crash or Ctrl-C in a previous email-rag run leaves
    behind:

      * the long-lived `outlook_daemon.ps1` (holds the named pipe),
      * any idle `outlook_pipe_client.ps1` (harmless on its own, but
        confusing in `Get-Process`),
      * `OclOutlookDaemon_*` schtasks whose BATs reference a path that
        no longer exists,
      * `ocl_daemon_*.bat` files in %TEMP% from old schtasks launches,
      * the previous `outlook-daemon-stderr.log` if its handle is held
        open by a partially-killed PowerShell process (new daemon's
        `Out-File -Append` then fails with "file is being used by
        another process").

    Args:
      kill_procs: stop stale powershell.exe processes whose CommandLine
        references outlook_daemon.ps1 / outlook_pipe_client.ps1. Only safe
        when the daemon is NOT currently alive (otherwise we'd kill the
        live daemon the user is talking to).
      rotate_log: delete outlook-daemon-stderr.log and outlook-daemon.pid.
        Only safe when the daemon is dead (otherwise the live daemon's
        append would fail mid-write).

    Always-on (regardless of flags):
      * unregister stale OclOutlookDaemon_* schtasks (cheap, harmless)
      * delete ocl_daemon_*.bat files older than 5 minutes

    We tolerate all errors here — best-effort cleanup that should never
    block startup.
    """
    proc_kill_step = ""
    if kill_procs:
        proc_kill_step = (
            "Get-CimInstance Win32_Process -Filter \"Name='powershell.exe'\" | "
            "Where-Object { $_.CommandLine -like '*outlook_daemon.ps1*' -or "
            "$_.CommandLine -like '*outlook_pipe_client.ps1*' } | "
            "ForEach-Object { Stop-Process -Id $_.ProcessId -Force; "
            "[Console]::Error.WriteLine(\"killed outlook proc pid=$($_.ProcessId)\") }; "
        )

    log_rotate_step = ""
    if rotate_log:
        log_rotate_step = (
            "if (Test-Path -LiteralPath 'C:\\Users\\wanjizheng\\AppData\\Local\\Temp\\outlook-daemon-stderr.log') { "
            "  try { Remove-Item -LiteralPath 'C:\\Users\\wanjizheng\\AppData\\Local\\Temp\\outlook-daemon-stderr.log' -Force "
            "  } catch { [Console]::Error.WriteLine(\"log rotation failed: $($_.Exception.Message)\") } "
            "}; "
            "if (Test-Path -LiteralPath 'C:\\Users\\wanjizheng\\AppData\\Local\\Temp\\outlook-daemon.pid') { "
            "  Remove-Item -LiteralPath 'C:\\Users\\wanjizheng\\AppData\\Local\\Temp\\outlook-daemon.pid' -Force "
            "  -ErrorAction SilentlyContinue "
            "}; "
        )

    ps_script = (
        "$ErrorActionPreference = 'SilentlyContinue'; "
        + proc_kill_step
        + "Get-ScheduledTask -ErrorAction SilentlyContinue | "
        "Where-Object { $_.TaskName -like 'OclOutlookDaemon_*' } | "
        "ForEach-Object { Unregister-ScheduledTask -TaskName $_.TaskName "
        "-Confirm:$false; "
        "[Console]::Error.WriteLine(\"deleted schtask=$($_.TaskName)\") }; "
        "Get-ChildItem 'C:\\Users\\wanjizheng\\AppData\\Local\\Temp\\ocl_daemon_*.bat' "
        "-ErrorAction SilentlyContinue | "
        "Where-Object { $_.LastWriteTime -lt (Get-Date).AddMinutes(-5) } | "
        "ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue; "
        "[Console]::Error.WriteLine(\"deleted bat=$($_.Name)\") }; "
        + log_rotate_step
        + "'cleanup_ok'"
    )
    try:
        subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", ps_script],
            capture_output=True, timeout=15,
            creationflags=CREATE_NO_WINDOW,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        # Best-effort; don't let cleanup failures stop us from starting.
        pass


def _launch_daemon_via_schtasks() -> None:
    """Launch outlook_daemon.ps1 via `schtasks /RL LIMITED` — no BAT wrapper.

    Used when the current shell is elevated. /RL LIMITED runs the task as
    the user's normal (non-elevated) token, which can attach to non-elevated
    Outlook COM.

    Why no BAT file?
      Earlier versions wrapped the powershell invocation in a tiny .bat
      file in %TEMP% and passed that to schtasks. But schtasks runs .bat
      TRs by spawning `cmd.exe /c <bat>` — which **always flashes a CMD
      console window**, even though the powershell inside is `-WindowStyle
      Hidden`. There is no way to suppress this for a .bat TR; the
      intermediate cmd.exe is non-negotiable.

      Passing the powershell invocation directly as the TR sidesteps this:
      schtasks calls `CreateProcess` on powershell.exe directly, with no
      cmd.exe intermediate, so no console window flashes. Bonus: one less
      file in %TEMP% to clean up.
    """
    daemon_win = subprocess.run(
        ["wslpath", "-w", str(OUTLOOK_DAEMON_PS1)],
        capture_output=True, text=True, timeout=5,
        creationflags=CREATE_NO_WINDOW,
    ).stdout.strip()

    uniq = f"{os.getpid()}_{int(time.time() * 1000)}"
    task_name = f"OclOutlookDaemon_{uniq}"
    # Build TR: powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "<daemon_win>".
    # The .ps1 path is quoted so spaces in the Windows path are preserved.
    tr = (
        f'powershell.exe -NoProfile -ExecutionPolicy Bypass '
        f'-WindowStyle Hidden -File "{daemon_win}" '
        f'-IdleTimeoutSec 86400'
    )

    try:
        subprocess.run(
            [SCHTASKS_EXE, "/Create", "/F", "/TN", task_name, "/TR", tr,
             "/SC", "ONCE", "/ST", "00:00", "/RL", "LIMITED"],
            capture_output=True, timeout=10,
            creationflags=CREATE_NO_WINDOW,
            check=True,
        )
    except (subprocess.CalledProcessError, FileNotFoundError) as exc:
        raise RuntimeError(f"schtasks /Create failed: {exc}") from exc
    try:
        subprocess.run(
            [SCHTASKS_EXE, "/Run", "/TN", task_name],
            capture_output=True, timeout=10,
            creationflags=CREATE_NO_WINDOW,
            check=True,
        )
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(f"schtasks /Run failed: {exc}") from exc


def _windows_script_path(path: Path) -> str:
    """PowerShell -File needs a Windows path even for non-elevated WSL launches."""
    if os.name == "nt":
        return str(path)
    result = subprocess.run(["wslpath", "-w", str(path)], capture_output=True,
                            text=True, timeout=5, check=True)
    translated = result.stdout.strip()
    if not translated:
        raise RuntimeError("Unable to translate Outlook bridge path")
    return translated


def _ensure_daemon() -> None:
    """Start the hidden Outlook daemon if it isn't already running.

    PowerShell startup details:
      * `-WindowStyle Hidden` keeps the daemon console off-screen.
      * When the current shell is elevated, we launch via
        `schtasks /RL LIMITED` to drop to non-elevated; otherwise an elevated
        PowerShell can't attach to non-elevated Outlook COM (returns
        CO_E_SERVER_EXEC_FAILURE after ~30s).

    start_daemon.vbs is intentionally NOT used: it embeds shell-style
    `>` / `2>` redirects, which `WshShell.Run` passes as positional args to
    powershell.exe (not interpreted as redirects) and the daemon silently
    fails to start.

    Always-on cleanup:
      * If daemon is ALIVE: do a SAFE partial cleanup — unregister any
        stale OclOutlookDaemon_* schtasks and delete stale ocl_daemon_*.bat
        files from previous runs. Do NOT kill the live daemon or rotate
        its log (kill_procs=False, rotate_log=False).
      * If daemon is DEAD: do FULL cleanup before relaunch — kill any
        leftover powershell procs, delete schtasks, delete BATs, rotate
        the log. This avoids the previous-run-leaves-held-handle failure
        where the new daemon's Out-File -Append silently fails.
    """
    if _daemon_ping():
        # Daemon is up and serving requests. Light cleanup only — do NOT
        # touch the live daemon process or its log file.
        _kill_stale_outlook_state(kill_procs=False, rotate_log=False)
        return
    # Full cleanup before launching new daemon — otherwise a crashed prior
    # run leaves behind a daemon that holds the pipe (can't bind a new one)
    # and a log file with a held handle (new daemon's Out-File -Append then
    # fails with "file is being used by another process", silently dropping
    # every subsequent log line).
    _kill_stale_outlook_state()
    if not OUTLOOK_DAEMON_PS1.exists():
        raise RuntimeError(f"Outlook daemon script missing: {OUTLOOK_DAEMON_PS1}")
    try:
        if _is_elevated():
            _launch_daemon_via_schtasks()
        else:
            subprocess.Popen(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                 "-WindowStyle", "Hidden", "-File", _windows_script_path(OUTLOOK_DAEMON_PS1)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW,
            )
    except FileNotFoundError as exc:
        raise RuntimeError(f"powershell.exe not found on PATH: {exc}") from exc
    # Wait for the daemon to come up. PID file is written early, but the
    # pipe server isn't bound until the main loop reaches WaitForConnection.
    # schtasks + PowerShell cold start can take 5-10s; allow 30s.
    deadline = time.time() + 30.0
    while time.time() < deadline:
        if _daemon_ping():
            return
        time.sleep(0.4)
    raise RuntimeError(
        "Outlook daemon did not come up. "
        "Check %TEMP%\\outlook-daemon-stderr.log for details."
    )


def _start_pipe_proc() -> subprocess.Popen:
    if not OUTLOOK_PIPE_CLIENT_PS1.exists():
        raise RuntimeError(f"Outlook pipe client missing: {OUTLOOK_PIPE_CLIENT_PS1}")
    return subprocess.Popen(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-WindowStyle", "Hidden", "-File", _windows_script_path(OUTLOOK_PIPE_CLIENT_PS1)],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=CREATE_NO_WINDOW,
    )


def _ensure_pipe_proc() -> subprocess.Popen:
    """Return a live, persistent outlook_pipe_client.ps1 Popen.

    Auto-starts the daemon if needed. Re-spawns the pipe client if its
    previous instance died (e.g. after a daemon restart).
    """
    global _pipe_proc
    with _get_pipe_lock():
        if _pipe_alive():
            return _pipe_proc
        # If the previous proc exited, drain anything left in it.
        if _pipe_proc is not None:
            try:
                _pipe_proc.kill()
            except Exception:
                pass
            _pipe_proc = None
        _ensure_daemon()
        _pipe_proc = _start_pipe_proc()
        return _pipe_proc


def _pipe_request(req: dict[str, Any], timeout_sec: float = 60.0) -> dict[str, Any]:
    """Send one JSON-line request to the pipe and read one JSON-line response.

    Handles pipe-client death mid-request by killing, restarting once, and
    retrying. Two restarts max — beyond that the caller gets a clear error.
    """
    if timeout_sec <= 0:
        raise ValueError("timeout_sec must be positive")
    line = json.dumps(req, ensure_ascii=False) + "\n"
    last_err: Exception | None = None
    # A persistent JSON-line pipe has one shared stdin/stdout pair. Serialize
    # the complete exchange so concurrent callers cannot consume each
    # other's responses.
    with _pipe_request_lock:
        for _attempt in range(2):
            proc = _ensure_pipe_proc()
            try:
                assert proc.stdin is not None and proc.stdout is not None
                proc.stdin.write(line.encode("utf-8"))
                proc.stdin.flush()
            except (BrokenPipeError, OSError) as exc:
                last_err = exc
                _force_kill_pipe()
                continue

            try:
                ready, _, _ = select.select([proc.stdout], [], [], timeout_sec)
            except (OSError, ValueError) as exc:
                last_err = exc
                _force_kill_pipe()
                continue
            if not ready:
                last_err = TimeoutError(
                    f"Outlook pipe request timed out after {timeout_sec:.1f}s"
                )
                _force_kill_pipe()
                continue
            try:
                resp_line = proc.stdout.readline()
            except OSError as exc:
                last_err = exc
                _force_kill_pipe()
                continue
            if not resp_line:
                last_err = RuntimeError("pipe client closed stdout before responding")
                _force_kill_pipe()
                continue
            text = resp_line.decode("utf-8", errors="replace").strip()
            try:
                return json.loads(text)
            except json.JSONDecodeError as exc:
                _force_kill_pipe()
                raise RuntimeError(
                    f"Outlook daemon returned non-JSON: {text[:300]}"
                ) from exc
    raise RuntimeError(f"Outlook pipe request failed after retry: {last_err}")


def _force_kill_pipe() -> None:
    global _pipe_proc
    with _get_pipe_lock():
        if _pipe_proc is not None:
            try:
                _pipe_proc.kill()
            except Exception:
                pass
            _pipe_proc = None


def shutdown_pipe() -> None:
    """Best-effort cleanup; safe to call multiple times.

    IMPORTANT: do NOT send `{"cmd":"shutdown"}` to the pipe client here.
    The pipe client forwards that to the long-lived outlook_daemon.ps1,
    which would exit on every Python invocation. With the daemon having a
    24h IdleTimeout, killing it on every Python exit would force a 5-7s
    daemon cold start (and a brief CMD-window flash from the schtasks
    relaunch) on every email-rag command. Instead we just reap the short-
    lived pipe client process; the daemon stays alive across Python runs.
    """
    global _pipe_proc
    with _get_pipe_lock():
        if _pipe_proc is None:
            return
        if _pipe_proc.poll() is None:
            # Pipe client still alive. Closing stdin makes the
            # outlook_pipe_client.ps1 ReadLine() loop return null and exit
            # gracefully; if it doesn't exit in 3s, kill it as fallback.
            try:
                if _pipe_proc.stdin is not None:
                    _pipe_proc.stdin.close()
            except Exception:
                pass
            try:
                _pipe_proc.wait(timeout=3)
            except Exception:
                try:
                    _pipe_proc.kill()
                except Exception:
                    pass
        _pipe_proc = None


import atexit
atexit.register(shutdown_pipe)

# NOTE: there is deliberately NO atexit handler for the persistent embed server
# spawned by _ensure_embed_server(). The server runs in its own session
# (start_new_session=True), survives Python exit, and self-exits after
# EMBED_SERVER_IDLE_TIMEOUT_MS of inactivity. The mirror to this is
# shutdown_pipe() above, which similarly does not kill the long-lived outlook
# daemon. See [[email-rag-skip-existing-design]] in MEMORY for context.


def run_outlook_mail_read(
    count: int,
    folder: str,
    hours: int,
    mode: str,
    max_chars: int,
    attachments_mode: str,
    attachments_dir: Path,
    before_iso: str | None = None,
) -> list[dict[str, Any]]:
    effective_count = max(1, min(int(count), MAX_OUTLOOK_READ_COUNT))

    # Translate attachments_dir to a Windows path for the daemon, since the
    # daemon runs as a real Windows process and SaveAsFile needs a Win32 path.
    attachments_dir_win = ""
    if attachments_mode == "export":
        attachments_dir.mkdir(parents=True, exist_ok=True)
        try:
            attachments_dir.chmod(0o700)
        except OSError:
            pass
        try:
            attachments_dir_win = subprocess.run(
                ["wslpath", "-w", str(attachments_dir.resolve())],
                capture_output=True, text=True, timeout=5,
                creationflags=CREATE_NO_WINDOW,
            ).stdout.strip()
        except Exception:
            attachments_dir_win = ""

    req: dict[str, Any] = {
        "cmd": "list",
        "count": effective_count,
        "folder": folder,
        "hours": hours,
        "mode": mode,
        "max_chars": max_chars,
        "attachments": attachments_mode,
    }
    if attachments_mode == "export":
        req["attachments_dir"] = str(attachments_dir)
        req["attachments_dir_win"] = attachments_dir_win
    if before_iso:
        req["before_iso"] = before_iso

    # Attachment export performs COM SaveAs calls inside the Outlook daemon and
    # legitimately takes longer for a larger page. Keep a hard ceiling so a
    # wedged COM call still fails, while avoiding false 120s timeouts on healthy
    # pages. Liveness is tracked separately by PID/systemd unit.
    per_item_budget = 3.0 if attachments_mode == "export" else 1.0
    read_timeout_sec = min(600.0, max(120.0, 30.0 + (count * per_item_budget)))
    resp = _pipe_request(req, timeout_sec=read_timeout_sec)
    if attachments_mode == "export":
        _secure_private_tree(attachments_dir)
    if isinstance(resp, dict):
        if resp.get("error"):
            raise RuntimeError(str(resp["error"]))
        # PowerShell's ConvertTo-Json unwraps single-element collections into a
        # single object instead of a one-element array. Wrap it back so the
        # caller always gets a list.
        if "EntryID" in resp:
            return [resp]
        raise RuntimeError(
            f"Outlook daemon returned unexpected dict: {str(resp)[:300]}"
        )
    if not isinstance(resp, list):
        raise RuntimeError(
            f"Outlook daemon returned unexpected payload type: {type(resp).__name__}: "
            f"{str(resp)[:300]}"
        )
    return resp


def _secure_private_tree(root: Path) -> None:
    """Best-effort 0700/0600 permissions for private exported email data."""
    if not root.exists() or root.is_symlink():
        return
    try:
        root.chmod(0o700)
    except OSError:
        pass
    for path in root.rglob("*"):
        if path.is_symlink():
            continue
        try:
            path.chmod(0o700 if path.is_dir() else 0o600)
        except OSError:
            pass


def _resolve_before_cursor(rows: list[dict[str, Any]], prev_before_iso: str | None) -> str | None:
    import base64
    keyed = [(parse_utc_timestamp(str(r.get("ReceivedTime") or "")),
              (str(r.get("StoreID") or "") + "|" + str(r.get("EntryID") or "")).lower()) for r in rows]
    keyed = [(t, k) for t, k in keyed if t is not None]
    if not keyed:
        return None
    stamp, key = min(keyed)
    payload = json.dumps({"t": stamp.isoformat(), "k": key}, separators=(",", ":"))
    cursor = "v2:" + base64.b64encode(payload.encode()).decode()
    if cursor == prev_before_iso:
        raise RuntimeError("Mail paging cursor did not advance; refusing to skip a boundary")
    return cursor


def estimate_outlook_folder_counts(folders: list[str]) -> dict[str, Any]:
    folder_list = [f.strip() for f in folders if f and f.strip()]
    if not folder_list:
        folder_list = ["InboxAll", "Sent"]

    resp = _pipe_request({"cmd": "estimate", "folders": folder_list}, timeout_sec=120.0)
    if not isinstance(resp, dict):
        raise RuntimeError(f"Outlook daemon returned unexpected estimate type: {type(resp).__name__}")
    if not resp.get("ok", False):
        raise RuntimeError(str(resp.get("error", "Outlook estimate failed")))
    return resp


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def vec_norm(vec: list[float]) -> float:
    s = 0.0
    for x in vec:
        s += x * x
    return math.sqrt(s)


def dot(a: list[float], b: list[float]) -> float:
    return sum(x * y for x, y in zip(a, b))


def cosine(a: list[float], b: list[float], norm_a: float | None = None, norm_b: float | None = None) -> float:
    na = norm_a if norm_a is not None else vec_norm(a)
    nb = norm_b if norm_b is not None else vec_norm(b)
    if na == 0 or nb == 0:
        return 0.0
    return dot(a, b) / (na * nb)


def embed_texts(base_url: str, api_key: str, model: str, texts: list[str], timeout_sec: int) -> list[list[float]]:
    if not texts:
        return []

    url = base_url.rstrip("/") + "/embeddings"
    payload = {"model": model, "input": texts}
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), method="POST")
    req.add_header("Content-Type", "application/json")
    if api_key:
        req.add_header("Authorization", f"Bearer {api_key}")

    try:
        with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
            body = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace") if exc.fp else ""
        raise RuntimeError(f"Embedding HTTP {exc.code}: {body[:300]}") from exc
    except Exception as exc:
        raise RuntimeError(f"Embedding request failed: {exc}") from exc

    try:
        data = json.loads(body)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Invalid embedding response JSON: {body[:300]}") from exc

    rows = data.get("data")
    if not isinstance(rows, list) or not rows:
        raise RuntimeError(f"Invalid embedding response shape: {body[:300]}")

    rows = sorted(rows, key=lambda r: int(r.get("index", 0)))
    out: list[list[float]] = []
    for r in rows:
        emb = r.get("embedding")
        if not isinstance(emb, list) or not emb:
            raise RuntimeError("Embedding response missing vector")
        out.append([float(x) for x in emb])
    return out


LOCAL_EMBED_MAX_CHARS = 2000   # truncate each text before sending to local model (~512 tokens)
LOCAL_EMBED_BATCH_SIZE = 16    # smaller failure domain when the shared WSL2 GPU is contended

# Persistent embed server (used by search query path; amortises ~26s model load
# across many searches within a session and across Python invocations).
# TCP because (a) it survives Python exit on Linux/WSL2 when the child is
# spawned with start_new_session=True (POSIX setsid → own session, no SIGHUP
# when parent shell dies), (b) Python doesn't need a separate "pipe client"
# subprocess the way outlook_daemon does — socket.create_connection() does
# the job inline.
EMBED_SERVER_HOST = "127.0.0.1"
EMBED_SERVER_PORT = int(os.environ.get("EMAIL_RAG_EMBED_PORT", "8765"))
EMBED_SERVER_IDLE_TIMEOUT_MS = int(os.environ.get("EMAIL_RAG_EMBED_IDLE_MS", str(30 * 60 * 1000)))
EMBED_SERVER_LOG = Path.home() / ".openclaw" / "workspace" / "data" / "email_rag" / "embed-server.log"
EMBED_SERVER_PIDFILE = Path.home() / ".openclaw" / "email-rag-embed-server.pid"
EMBED_SERVER_LOCKDIR = Path.home() / ".openclaw" / "email-rag-embed-server.lock"
EMBED_SERVER_START_TIMEOUT = 90.0
EMBED_SERVER_REQUEST_TIMEOUT = float(os.environ.get("EMAIL_RAG_EMBED_REQUEST_TIMEOUT", "10"))
EMBED_ONESHOT_FALLBACK_TIMEOUT = float(os.environ.get("EMAIL_RAG_EMBED_FALLBACK_TIMEOUT", "20"))
EMBED_SERVER_PROBE_TIMEOUT = 1.5
LOCAL_QUERY_EMBED_MODE = os.environ.get("EMAIL_RAG_QUERY_EMBED_MODE", "oneshot").strip().lower()
if LOCAL_QUERY_EMBED_MODE not in {"oneshot", "persistent"}:
    LOCAL_QUERY_EMBED_MODE = "oneshot"
LOCAL_QUERY_EMBED_TIMEOUT = float(os.environ.get("EMAIL_RAG_QUERY_EMBED_TIMEOUT", "6"))
LOCAL_QUERY_EMBED_ATTEMPTS = max(
    1, int(os.environ.get("EMAIL_RAG_QUERY_EMBED_ATTEMPTS", "2"))
)
_embed_server_proc: "subprocess.Popen | None" = None


def _embed_server_alive(timeout_sec: float = EMBED_SERVER_PROBE_TIMEOUT) -> bool:
    """True only when the expected embed server answers its ping protocol."""
    try:
        with socket.create_connection(
            (EMBED_SERVER_HOST, EMBED_SERVER_PORT), timeout=timeout_sec
        ) as sock:
            sock.settimeout(timeout_sec)
            sock.sendall(b'{"cmd":"ping"}\n')
            response = b""
            while b"\n" not in response and len(response) < 4096:
                chunk = sock.recv(4096)
                if not chunk:
                    break
                response += chunk
            if b"\n" not in response:
                return False
            payload = json.loads(response.split(b"\n", 1)[0].decode("utf-8"))
            return payload == {"ok": True, "service": "email-rag-embed"}
    except (ConnectionRefusedError, OSError, socket.timeout, json.JSONDecodeError):
        return False


def _is_expected_embed_server_pid(pid: int) -> bool:
    """Avoid signalling an unrelated process after PID reuse/stale pidfiles."""
    if pid <= 0:
        return False
    try:
        cmdline = Path(f"/proc/{pid}/cmdline").read_bytes().replace(b"\0", b" ")
    except OSError:
        return False
    return b"embed_local.mjs" in cmdline and b"--server" in cmdline


def _acquire_embed_starter_lock(timeout: float = EMBED_SERVER_START_TIMEOUT) -> bool:
    """Atomic mkdir-based starter lock. True iff we are the starter.

    Two concurrent `email-rag search` calls would otherwise both try to spawn
    `node`, the loser getting EADDRINUSE and a confusing restart. With this
    lock, only one Python process spawns; the other just waits for the port.
    """
    deadline = time.monotonic() + timeout
    EMBED_SERVER_LOCKDIR.parent.mkdir(parents=True, exist_ok=True)
    while time.monotonic() < deadline:
        try:
            EMBED_SERVER_LOCKDIR.mkdir()
            return True
        except FileExistsError:
            # Sanity check: if the lock holder's PID file points at a dead
            # process, steal the lock. (os.kill(pid, 0) doesn't work on
            # Windows; we only run this on Linux/WSL2.)
            pidfile = EMBED_SERVER_PIDFILE
            if pidfile.exists():
                try:
                    pid = int(pidfile.read_text().strip())
                    if not _is_expected_embed_server_pid(pid):
                        raise ProcessLookupError
                    os.kill(pid, 0)
                except (ValueError, ProcessLookupError, PermissionError, OSError):
                    try:
                        EMBED_SERVER_LOCKDIR.rmdir()
                    except FileNotFoundError:
                        pass
                    continue
            time.sleep(0.1)
    return False


def _release_embed_starter_lock() -> None:
    try:
        EMBED_SERVER_LOCKDIR.rmdir()
    except FileNotFoundError:
        pass


def _spawn_embed_server(embed_script: Path) -> "subprocess.Popen":
    """Spawn the persistent embed server, detached from this Python process.

    On Linux/WSL2 we use start_new_session=True (POSIX setsid) so the child
    becomes its own session leader — it survives both normal Python exit AND
    kill -9 of the parent shell, and won't get SIGHUP. We keep
    CREATE_NO_WINDOW for any stray Windows-binary fallback.

    Stdin goes to DEVNULL; stderr/stdout go to EMBED_SERVER_LOG via the parent
    opening it in append mode and inheriting the FD into the child. The parent
    closes its own FD immediately after Popen so we don't leak.
    """
    EMBED_SERVER_LOG.parent.mkdir(parents=True, exist_ok=True)
    log_fd = open(EMBED_SERVER_LOG, "ab")
    try:
        EMBED_SERVER_LOG.chmod(0o600)
    except OSError:
        pass
    try:
        proc = subprocess.Popen(
            [
                "node", str(embed_script),
                "--server", str(EMBED_SERVER_PORT),
                "--idle-timeout-ms", str(EMBED_SERVER_IDLE_TIMEOUT_MS),
            ],
            stdin=subprocess.DEVNULL,
            stdout=log_fd,
            stderr=subprocess.STDOUT,
            start_new_session=True,
            creationflags=CREATE_NO_WINDOW,
            close_fds=True,
        )
    finally:
        log_fd.close()
    return proc


def _ensure_embed_server(embed_script: Path = LOCAL_EMBED_SCRIPT) -> None:
    """Idempotent: probe → if down, become starter or wait → spawn → poll.

    First call pays the ~26s model load cost; every subsequent call within
    the idle window finds the server already alive.
    """
    global _embed_server_proc
    if _embed_server_alive():
        return
    if not embed_script.exists():
        raise RuntimeError(f"Embed server script missing: {embed_script}")
    if not _acquire_embed_starter_lock():
        # Another Python process is starting it. Just wait for the port.
        deadline = time.monotonic() + EMBED_SERVER_START_TIMEOUT
        while time.monotonic() < deadline:
            if _embed_server_alive():
                return
            time.sleep(0.2)
        raise RuntimeError(
            f"Embed server did not become ready in {EMBED_SERVER_START_TIMEOUT}s. "
            f"Check {EMBED_SERVER_LOG} for details."
        )
    try:
        _embed_server_proc = _spawn_embed_server(embed_script)
        deadline = time.monotonic() + EMBED_SERVER_START_TIMEOUT
        while time.monotonic() < deadline:
            if _embed_server_alive():
                return
            if _embed_server_proc.poll() is not None:
                # Server died before binding — usually EADDRINUSE (someone
                # else won the race) or model-load failure.
                try:
                    log_tail = EMBED_SERVER_LOG.read_text(errors="replace")[-500:]
                except Exception:
                    log_tail = "(could not read log)"
                raise RuntimeError(
                    f"Embed server exited rc={_embed_server_proc.returncode} before binding. "
                    f"Recent log: {log_tail}"
                )
            time.sleep(0.2)
        # Timed out waiting for the bind. Kill the child and surface an error.
        try:
            _embed_server_proc.kill()
        except Exception:
            pass
        raise RuntimeError(
            f"Embed server did not bind port {EMBED_SERVER_PORT} within "
            f"{EMBED_SERVER_START_TIMEOUT}s. Check {EMBED_SERVER_LOG}."
        )
    finally:
        _release_embed_starter_lock()


def _embed_server_request_once(
    texts: list[str],
    timeout_sec: float = EMBED_SERVER_REQUEST_TIMEOUT,
) -> list[list[float]]:
    """Send one JSON-line request to the persistent server."""
    if not texts:
        return []
    _ensure_embed_server()
    safe_texts = [t[:LOCAL_EMBED_MAX_CHARS] for t in texts]
    payload = (json.dumps({"texts": safe_texts}, ensure_ascii=False) + "\n").encode("utf-8")
    deadline = time.monotonic() + timeout_sec
    try:
        with socket.create_connection(
            (EMBED_SERVER_HOST, EMBED_SERVER_PORT), timeout=EMBED_SERVER_PROBE_TIMEOUT
        ) as sock:
            sock.settimeout(max(0.5, deadline - time.monotonic()))
            sock.sendall(payload)
            buf = b""
            while True:
                chunk = sock.recv(65536)
                if not chunk:
                    raise RuntimeError("Embed server closed connection before responding")
                buf += chunk
                nl = buf.find(b"\n")
                if nl != -1:
                    line = buf[:nl].decode("utf-8", errors="replace").strip()
                    data = json.loads(line)
                    if "error" in data:
                        raise RuntimeError(f"Embed server error: {data['error']}")
                    vectors = data.get("vectors")
                    if not isinstance(vectors, list):
                        raise RuntimeError(
                            f"Embed server returned unexpected payload: {str(data)[:200]}"
                        )
                    if len(vectors) != len(texts):
                        raise RuntimeError(
                            f"Embed server count mismatch: expected {len(texts)}, "
                            f"got {len(vectors)}"
                        )
                    return [[float(x) for x in v] for v in vectors]
    except (ConnectionRefusedError, ConnectionResetError, BrokenPipeError,
            socket.timeout) as exc:
        raise RuntimeError(
            f"Embed server unreachable ({type(exc).__name__}: {exc}). "
            f"It may have crashed or hung on CUDA contention."
        ) from exc


def _embed_server_request(
    texts: list[str],
    timeout_sec: float = EMBED_SERVER_REQUEST_TIMEOUT,
) -> list[list[float]]:
    """Use the warm server, automatically recovering from a stuck CUDA request.

    A WSL2 CUDA worker can remain alive while a native call is wedged behind a
    Windows GPU consumer.  Previously every search waited 60 seconds and then
    told the user to kill the process manually.  Fail fast, remove the wedged
    server, and retry once through the reliable one-shot path instead.
    """
    if not texts:
        return []
    try:
        return _embed_server_request_once(texts, timeout_sec=timeout_sec)
    except RuntimeError as server_exc:
        _kill_embed_server()
        try:
            return embed_texts_local(
                texts,
                LOCAL_EMBED_SCRIPT,
                timeout_sec=max(1, int(EMBED_ONESHOT_FALLBACK_TIMEOUT)),
            )
        except RuntimeError as fallback_exc:
            raise RuntimeError(
                f"Persistent embed server failed ({server_exc}); "
                f"automatic one-shot fallback also failed ({fallback_exc})"
            ) from fallback_exc


def _kill_embed_server() -> bool:
    """Kill the persistent embed server. Returns True if we sent a kill signal.

    On WSL2/Linux the node process is a Linux process (not visible to
    powershell.exe on the Windows side), so we kill by PID file via os.kill.
    """
    global _embed_server_proc
    killed = False
    if _embed_server_proc is not None and _embed_server_proc.poll() is None:
        try:
            _embed_server_proc.kill()
            killed = True
        except Exception:
            pass
        try:
            _embed_server_proc.wait(timeout=2)
        except Exception:
            pass
    _embed_server_proc = None
    # Kill via PID file (handles servers started by previous Python invocations).
    if EMBED_SERVER_PIDFILE.exists():
        try:
            pid = int(EMBED_SERVER_PIDFILE.read_text().strip())
            if _is_expected_embed_server_pid(pid):
                try:
                    os.kill(pid, 15)  # SIGTERM
                    killed = True
                    # Wait briefly for graceful exit
                    for _ in range(20):
                        try:
                            os.kill(pid, 0)
                        except ProcessLookupError:
                            break
                        time.sleep(0.1)
                    else:
                        # Still alive after 2s — force kill
                        try:
                            os.kill(pid, 9)
                        except ProcessLookupError:
                            pass
                except ProcessLookupError:
                    pass  # already dead
                except PermissionError:
                    pass
        except (ValueError, OSError):
            pass
    EMBED_SERVER_PIDFILE.unlink(missing_ok=True)
    return killed


def embed_texts_local(texts: list[str], script_path: Path, timeout_sec: int | None = None) -> list[list[float]]:
    """Embed texts using local GGUF model via node-llama-cpp (embed_local.mjs, one-shot mode).

    For large batches use embed_texts_local_persistent() instead.
    timeout_sec defaults to LOCAL_EMBED_TIMEOUT_OVERHEAD + len(texts) * LOCAL_EMBED_TIMEOUT_PER_TEXT.
    """
    if not texts:
        return []
    if timeout_sec is None:
        timeout_sec = LOCAL_EMBED_TIMEOUT_OVERHEAD + len(texts) * LOCAL_EMBED_TIMEOUT_PER_TEXT

    # Truncate individual texts to avoid V8 string length errors / tokenizer overflow
    safe_texts = [t[:LOCAL_EMBED_MAX_CHARS] for t in texts]
    req_json = json.dumps({"texts": safe_texts})
    try:
        result = subprocess.run(
            ["node", str(script_path), "--one-shot"],
            input=req_json.encode(),
            capture_output=True,
            timeout=timeout_sec,
            creationflags=CREATE_NO_WINDOW,
        )
    except FileNotFoundError:
        raise RuntimeError("node binary not found; install Node.js to use --embed-backend local")
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"Local embed timed out after {timeout_sec}s ({len(texts)} texts). "
            "Consider using --embed-backend api or increasing LOCAL_EMBED_TIMEOUT_PER_TEXT."
        )

    stdout = result.stdout.decode("utf-8", errors="replace").strip()
    stderr = result.stderr.decode("utf-8", errors="replace").strip()

    if not stdout:
        raise RuntimeError(f"Local embed produced no output. stderr: {stderr[:500]}")

    # Scan from the end: the actual {vectors}/{error} response is the last JSON line.
    # Skip CMake/tsc noise that node-llama-cpp leaks onto stdout on first call.
    last_json_error = None
    for line in reversed(stdout.splitlines()):
        line = line.strip()
        if not line:
            continue
        try:
            data = json.loads(line)
        except json.JSONDecodeError as e:
            last_json_error = (e, line)
            continue  # try next line further up
        if "error" in data:
            raise RuntimeError(f"Local embed error: {data['error']}")
        vectors = data.get("vectors")
        if not isinstance(vectors, list):
            # not the response we're looking for; keep scanning
            continue
        return [[float(x) for x in v] for v in vectors]

    if last_json_error is not None:
        e, line = last_json_error
        raise RuntimeError(f"Local embed: no valid response line in stdout (last bad line: {line[:300]})") from e
    raise RuntimeError(f"Local embed: no valid response line in stdout: {stdout[:300]}")


def embed_texts_local_persistent(
    texts: list[str],
    script_path: Path,
    batch_size: int = LOCAL_EMBED_BATCH_SIZE,
    progress_cb: Any = None,
) -> list[list[float]]:
    """Embed large text lists using a single persistent embed_local.mjs process (server mode).

    Spawns the node process once (model loads once), then sends texts in batches via stdin
    and reads results from stdout. Much faster than one-shot for large corpora.
    progress_cb(done, total) is called after each batch if provided.
    """
    if not texts:
        return []

    # node-llama-cpp can emit one warning per embedding batch. Keeping stderr
    # in an unread PIPE eventually fills the OS pipe buffer and deadlocks a
    # large sync. A temporary file is non-blocking for the child and still lets
    # us include diagnostics if the process fails.
    stderr_file = tempfile.TemporaryFile()
    try:
        proc = subprocess.Popen(
            ["node", str(script_path)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=stderr_file,
            creationflags=CREATE_NO_WINDOW,
        )
    except FileNotFoundError:
        stderr_file.close()
        raise RuntimeError("node binary not found; install Node.js to use --embed-backend local")

    def _stderr_tail(limit: int = 500) -> str:
        try:
            stderr_file.flush()
            stderr_file.seek(0, os.SEEK_END)
            size = stderr_file.tell()
            stderr_file.seek(max(0, size - limit))
            return stderr_file.read().decode("utf-8", errors="replace")
        except Exception:
            return ""

    all_vectors: list[list[float]] = []
    completed = False
    try:
        for i in range(0, len(texts), batch_size):
            batch = texts[i : i + batch_size]
            safe_batch = [t[:LOCAL_EMBED_MAX_CHARS] for t in batch]
            req_line = json.dumps({"texts": safe_batch}) + "\n"
            try:
                proc.stdin.write(req_line.encode())
                proc.stdin.flush()
            except BrokenPipeError:
                stderr_out = _stderr_tail()
                raise RuntimeError(f"embed_local.mjs process died. stderr: {stderr_out[:500]}")

            # Read the response line (blocking).
            # node-llama-cpp's first call (or after CUDA prebuilt mismatch) leaks
            # CMake/tsc build output onto stdout ("Not searching for unused variables...",
            # "-- The C compiler identification is GNU 13.3.0", etc.). Skip any
            # non-JSON preamble lines until we find a real {vectors}/{error} response.
            data = None
            skip_lines = 0
            MAX_SKIP_LINES = 50_000  # generous cap; cmake output is ~100 lines
            response_deadline = time.monotonic() + (
                LOCAL_EMBED_BATCH_TIMEOUT_OVERHEAD
                + len(batch) * LOCAL_EMBED_BATCH_TIMEOUT_PER_TEXT
            )
            while True:
                remaining_timeout = response_deadline - time.monotonic()
                if remaining_timeout <= 0:
                    raise RuntimeError(
                        "embed_local.mjs timed out waiting for an embedding response"
                    )
                try:
                    ready, _, _ = select.select(
                        [proc.stdout], [], [], remaining_timeout
                    )
                except (OSError, ValueError) as exc:
                    raise RuntimeError(
                        f"embed_local.mjs stdout wait failed: {exc}"
                    ) from exc
                if not ready:
                    raise RuntimeError(
                        "embed_local.mjs timed out waiting for an embedding response"
                    )
                resp_line = proc.stdout.readline()
                if not resp_line:
                    stderr_out = _stderr_tail() if proc.poll() is not None else ""
                    raise RuntimeError(
                        f"embed_local.mjs closed stdout before JSON response "
                        f"(after {skip_lines} non-JSON lines). stderr: {stderr_out[:500]}"
                    )
                decoded = resp_line.decode("utf-8", errors="replace").strip()
                if not decoded:
                    skip_lines += 1
                    if skip_lines > MAX_SKIP_LINES:
                        raise RuntimeError("embed_local.mjs: too many empty stdout lines")
                    continue
                try:
                    data = json.loads(decoded)
                    break
                except json.JSONDecodeError:
                    # Non-JSON noise from cmake/tsc child process — skip and try next line.
                    skip_lines += 1
                    if skip_lines > MAX_SKIP_LINES:
                        raise RuntimeError(
                            f"embed_local.mjs: too many non-JSON stdout lines; last: {decoded[:300]}"
                        )
                    continue

            if "error" in data:
                raise RuntimeError(f"Local embed error: {data['error']}")

            vectors = data.get("vectors")
            if not isinstance(vectors, list) or len(vectors) != len(batch):
                raise RuntimeError(
                    f"Local embed response count mismatch: expected {len(batch)}, got {len(vectors) if isinstance(vectors, list) else 'none'}"
                )
            all_vectors.extend([[float(x) for x in v] for v in vectors])

            if progress_cb is not None:
                progress_cb(len(all_vectors), len(texts))

        completed = True
    finally:
        try:
            proc.stdin.close()
        except Exception:
            pass
        try:
            if not completed and proc.poll() is None:
                # A native CUDA call can ignore graceful shutdown while wedged.
                # Kill immediately after our bounded response timeout.
                proc.kill()
                proc.wait(timeout=5)
            else:
                try:
                    proc.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait(timeout=5)
        finally:
            stderr_file.close()

    return all_vectors


def _resolve_embed_enabled(args: argparse.Namespace) -> bool:
    if getattr(args, "no_embed", False):
        return False
    return True


def _positive_int(raw: str) -> int:
    try:
        value = int(raw)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("expected an integer") from exc
    if value <= 0:
        raise argparse.ArgumentTypeError("expected a positive integer")
    return value


def _nonnegative_int(raw: str) -> int:
    try:
        value = int(raw)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("expected an integer") from exc
    if value < 0:
        raise argparse.ArgumentTypeError("expected a non-negative integer")
    return value


def _add_embedding_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--no-embed", action="store_true", help="Disable embedding generation")
    parser.add_argument(
        "--gpu-probe",
        dest="gpu_probe",
        action="store_true",
        default=True,
        help="(default) Probe Windows GPU before sync; warn if a non-WSL2 process is using it.",
    )
    parser.add_argument(
        "--no-gpu-probe",
        dest="gpu_probe",
        action="store_false",
        help="Skip the GPU contention probe at sync start.",
    )
    parser.add_argument("--embed-base-url", default=DEFAULT_EMBED_BASE_URL)
    parser.add_argument("--embed-model", default=DEFAULT_EMBED_MODEL)
    parser.add_argument("--embed-api-key", default=DEFAULT_EMBED_API_KEY)
    parser.add_argument("--embed-batch", type=_positive_int, default=DEFAULT_EMBED_BATCH)
    parser.add_argument("--embed-timeout", type=_positive_int, default=DEFAULT_EMBED_TIMEOUT)
    parser.add_argument(
        "--embed-backend",
        choices=["local", "api"],
        default=DEFAULT_EMBED_BACKEND,
        help=(
            "Embedding backend: 'local' uses GGUF via node-llama-cpp (default, GPU auto-detected, "
            "CPU fallback), 'api' uses the OpenAI-compatible HTTP API."
        ),
    )


def _add_paging_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--page-size",
        type=_positive_int,
        default=DEFAULT_PAGED_READ_SIZE,
        help="Read Outlook messages in smaller pages (default: 200, hard max: 500).",
    )
    parser.add_argument(
        "--disable-paged-read",
        action="store_true",
        help="Disable page-by-page reads and fetch each plan in one large request.",
    )


def _add_attachment_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--attachments",
        choices=["none", "meta", "export"],
        default=DEFAULT_ATTACHMENT_MODE,
        help="Attachment processing in sync.",
    )
    parser.add_argument("--attachments-dir", type=Path, default=DEFAULT_ATTACHMENT_EXPORT_DIR)
    parser.add_argument("--max-attachments-per-email", type=_nonnegative_int, default=DEFAULT_ATTACHMENT_MAX_PER_EMAIL)
    parser.add_argument("--attachment-max-chars", type=_nonnegative_int, default=DEFAULT_ATTACHMENT_MAX_CHARS)
    parser.add_argument(
        "--attachments-cleanup",
        choices=["none", "all"],
        default=DEFAULT_ATTACHMENT_CLEANUP,
        help="Cleanup exported attachment files after sync. 'all' deletes all files/dirs under attachments-dir.",
    )


def _add_skip_existing_args(parser: argparse.ArgumentParser) -> None:
    """Add --skip-existing / --no-skip-existing flags. Default ON.

    When enabled, after each Outlook page is fetched, rows whose entry_id
    (or canonical stable_key) already exists in the local emails table are
    dropped before _ingest_rows — so they don't get re-chunked, re-embedded,
    or re-upserted. The setting is global per sync run; it cannot be turned
    off per-plan.

    Use --no-skip-existing to force the old behavior (re-process every row,
    including ones already in DB). This is useful when an email's body has
    changed and you want to re-embed it.
    """
    parser.add_argument(
        "--skip-existing",
        dest="skip_existing",
        action="store_true",
        default=True,
        help="Skip Outlook rows whose entry_id is already in the DB (default: True — avoids duplicating RAG work).",
    )
    parser.add_argument(
        "--no-skip-existing",
        dest="skip_existing",
        action="store_false",
        help="Force re-ingest of every Outlook row, including those already in DB (re-chunk + re-embed everything).",
    )


def _compute_row_stable_key(row: dict[str, Any]) -> str:
    """Reproduce the same stable_key the _ingest_rows path uses, so we can
    pre-filter rows before they hit the DB.
    """
    def _norm(s: str) -> str:
        return re.sub(r"\s+", " ", (s or "").strip().lower())

    sender = str(row.get("Sender") or "")
    sender_email = str(row.get("SenderEmailAddress") or "")
    subject = str(row.get("Subject") or "")
    received_time = str(row.get("ReceivedTime") or "")
    body = str(row.get("Body") or "")
    stable_basis = "\n".join(
        [
            _norm(sender_email) or _norm(sender),
            _norm(subject),
            _norm(received_time),
            _norm(body)[:2000],
        ]
    )
    return sha256_text(stable_basis)


def _filter_existing_rows(
    conn: sqlite3.Connection,
    rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    """Drop rows whose entry_id or stable_key already exists in the emails table.

    Returns (filtered_rows, skipped_count). Filtered rows are the ones to
    actually ingest. Returns the original list unchanged if `rows` is empty
    or if there's nothing in the DB to compare against.

    Match strategy:
      * entry_id first (fast exact match, primary dedup key)
      * stable_key as fallback (handles Outlook EntryID drift between
        sessions and the synthetic:<sha> fallback for items with no EntryID)
    """
    if not rows:
        return rows, 0

    # Build lookup sets of all known entry_ids and stable_keys in one query
    # each. Loading the full tables is fine — emails table is text-only and
    # bounded by the user's mailbox size; for a 50k-mailbox this is ~5MB.
    try:
        existing_ids = {
            str(r[0]) for r in conn.execute("SELECT entry_id FROM emails").fetchall() if r[0]
        }
        existing_keys = {
            str(r[0]) for r in conn.execute("SELECT stable_key FROM emails WHERE stable_key IS NOT NULL").fetchall() if r[0]
        }
    except sqlite3.OperationalError:
        # Table doesn't exist yet (first sync) — nothing to dedup against
        return rows, 0

    if not existing_ids and not existing_keys:
        return rows, 0

    filtered: list[dict[str, Any]] = []
    skipped = 0
    for row in rows:
        entry_id = str(row.get("EntryID") or "").strip()
        if entry_id and entry_id in existing_ids:
            skipped += 1
            continue
        stable_key = _compute_row_stable_key(row)
        if stable_key in existing_keys:
            skipped += 1
            continue
        filtered.append(row)

    return filtered, skipped


def _clear_index(conn: sqlite3.Connection, drop_embedding_cache: bool) -> None:
    conn.execute("DELETE FROM email_chunks_fts")
    conn.execute("DELETE FROM email_chunks")
    conn.execute("DELETE FROM emails")
    if drop_embedding_cache:
        conn.execute("DELETE FROM embedding_cache")


def _extract_attachment_text(path: Path) -> str:
    if not path.exists():
        return ""
    try:
        proc = subprocess.run(
            [OFFICE_DOC_IO_CMD, "read", "--input", str(path), "--format", "text"],
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=CREATE_NO_WINDOW,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return ""
    if proc.returncode != 0:
        return ""
    raw = (proc.stdout or "").strip()
    if not raw:
        return ""
    try:
        payload = json.loads(raw)
        if isinstance(payload, dict):
            return str(payload.get("text") or "")
    except json.JSONDecodeError:
        return ""
    return ""


def _path_size_bytes(path: Path) -> int:
    if not path.exists():
        return 0
    if path.is_file():
        try:
            return path.stat().st_size
        except OSError:
            return 0
    total = 0
    for p in path.rglob("*"):
        if p.is_file():
            try:
                total += p.stat().st_size
            except OSError:
                pass
    return total


def _cleanup_exported_attachments(export_dir: Path) -> dict[str, Any]:
    report: dict[str, Any] = {
        "mode": "all",
        "dir": str(export_dir),
        "deletedFiles": 0,
        "deletedDirs": 0,
        "deletedBytes": 0,
    }
    allowed_root = DEFAULT_ATTACHMENT_EXPORT_DIR.parent.resolve()
    target = export_dir.expanduser().resolve()
    try:
        target.relative_to(allowed_root)
    except ValueError:
        report["error"] = (
            f"Refusing cleanup outside email-rag data root: {allowed_root}"
        )
        return report
    export_dir = target
    report["dir"] = str(export_dir)
    if not export_dir.exists():
        return report

    for child in list(export_dir.iterdir()):
        report["deletedBytes"] += _path_size_bytes(child)
        if child.is_dir():
            try:
                # Count directories and files before delete for reporting.
                report["deletedDirs"] += 1
                report["deletedDirs"] += sum(1 for p in child.rglob("*") if p.is_dir())
                report["deletedFiles"] += sum(1 for p in child.rglob("*") if p.is_file())
            except OSError:
                pass
            shutil.rmtree(child, ignore_errors=True)
        else:
            try:
                if child.is_file():
                    report["deletedFiles"] += 1
            except OSError:
                pass
            try:
                os.remove(child)
            except OSError:
                pass
    return report


def _ingest_rows(
    conn: sqlite3.Connection,
    rows: list[dict[str, Any]],
    *,
    embed_enabled: bool,
    embed_base_url: str,
    embed_api_key: str,
    embed_model: str,
    embed_batch: int,
    embed_timeout: int,
    embed_backend: str = DEFAULT_EMBED_BACKEND,
    embed_script: Path = LOCAL_EMBED_SCRIPT,
    stats: dict[str, Any],
    attachment_max_per_email: int,
    attachment_max_chars: int,
) -> str | None:
    pending_embed: list[tuple[int, str, int, str, str]] = []
    now = now_iso()

    for item in rows:
        entry_id = str(item.get("EntryID") or "").strip()

        subject = str(item.get("Subject") or "")
        sender = str(item.get("Sender") or "")
        sender_email = str(item.get("SenderEmailAddress") or "")
        received_time = str(item.get("ReceivedTime") or "")
        body = str(item.get("Body") or "")
        raw_body = body

        def _norm(s: str) -> str:
            return re.sub(r"\s+", " ", (s or "").strip().lower())

        # Stable dedup key to avoid duplicate ingestion when Outlook EntryID changes.
        stable_basis = "\n".join(
            [
                _norm(sender_email) or _norm(sender),
                _norm(subject),
                _norm(received_time),
                _norm(raw_body)[:2000],
            ]
        )
        stable_key = sha256_text(stable_basis)
        if not entry_id:
            # Some Outlook item shapes do not expose EntryID reliably.
            # Fall back to a deterministic synthetic key so ingestion does not drop them.
            entry_id = f"synthetic:{stable_key}"
        existing = conn.execute("SELECT entry_id FROM emails WHERE stable_key = ? LIMIT 1", (stable_key,)).fetchone()
        if existing and str(existing[0]) and str(existing[0]) != entry_id:
            # Reuse canonical entry_id so chunks don't duplicate under a new key.
            entry_id = str(existing[0])
            stats["dedupMerged"] += 1

        attachments = item.get("Attachments")
        attachment_text_blocks: list[str] = []
        if isinstance(attachments, list) and attachments:
            for att in attachments[: max(0, attachment_max_per_email)]:
                if not isinstance(att, dict):
                    continue
                p = str(att.get("ExportedPath") or "").strip()
                if not p:
                    continue
                ap = Path(p)
                if ap.suffix.lower() not in {".pdf", ".docx", ".pptx", ".xlsx", ".csv", ".txt", ".md"}:
                    continue
                txt = _extract_attachment_text(ap)
                if txt:
                    name = str(att.get("SafeFileName") or att.get("FileName") or "attachment")
                    attachment_text_blocks.append(f"[Attachment: {name}]\n{txt[: max(0, attachment_max_chars)]}")
                    stats["attachmentsTextUsed"] += 1
                else:
                    stats["attachmentsFailed"] += 1

        if attachment_text_blocks:
            body = body + "\n\n" + "\n\n".join(attachment_text_blocks)

        conn.execute(
            """
            INSERT INTO emails(entry_id, stable_key, subject, sender, sender_email, received_time, body, updated_at)
            VALUES(?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(entry_id) DO UPDATE SET
              stable_key=excluded.stable_key,
              subject=excluded.subject,
              sender=excluded.sender,
              sender_email=excluded.sender_email,
              received_time=excluded.received_time,
              body=excluded.body,
              updated_at=excluded.updated_at
            """,
            (entry_id, stable_key, subject, sender, sender_email, received_time, body, now),
        )

        conn.execute("DELETE FROM email_chunks_fts WHERE entry_id = ?", (entry_id,))
        conn.execute("DELETE FROM email_chunks WHERE entry_id = ?", (entry_id,))

        chunks = chunk_text(body)
        if not chunks and subject:
            chunks = [subject]

        for idx, c in enumerate(chunks):
            text_hash = sha256_text(c)
            emb_json = None
            emb_norm = None

            if embed_enabled:
                cached = conn.execute(
                    "SELECT embedding_json, embedding_norm FROM embedding_cache WHERE text_hash = ?",
                    (text_hash,),
                ).fetchone()
                if cached:
                    emb_json = cached[0]
                    emb_norm = float(cached[1])
                    stats["embeddedCacheHits"] += 1

            cur = conn.execute(
                """
                INSERT INTO email_chunks(
                  entry_id, chunk_index, subject, sender, received_time,
                  chunk_text, text_hash, embedding_json, embedding_norm, updated_at
                ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (entry_id, idx, subject, sender, received_time, c, text_hash, emb_json, emb_norm, now),
            )
            chunk_id = int(cur.lastrowid)
            if emb_json is not None:
                _store_chunk_vector_json(conn, chunk_id, emb_json, now)

            conn.execute(
                """
                INSERT INTO email_chunks_fts(entry_id, chunk_id, subject, sender, received_time, chunk_text)
                VALUES(?, ?, ?, ?, ?, ?)
                """,
                (entry_id, chunk_id, subject, sender, received_time, c),
            )

            if embed_enabled and emb_json is None:
                pending_embed.append((chunk_id, entry_id, idx, text_hash, c))

        stats["syncedEmails"] += 1
        stats["syncedChunks"] += len(chunks)

    if embed_enabled and pending_embed:
        all_texts = [x[4] for x in pending_embed]
        try:
            if embed_backend == "local":
                all_vectors = embed_texts_local_persistent(all_texts, embed_script)
            else:
                all_vectors = []
                for i in range(0, len(all_texts), max(1, embed_batch)):
                    batch_texts = all_texts[i : i + max(1, embed_batch)]
                    all_vectors.extend(
                        embed_texts(
                            base_url=embed_base_url,
                            api_key=embed_api_key,
                            model=embed_model,
                            texts=batch_texts,
                            timeout_sec=embed_timeout,
                        )
                    )
            if len(all_vectors) != len(pending_embed):
                raise RuntimeError(
                    f"Embedding count mismatch: expected {len(pending_embed)}, got {len(all_vectors)}"
                )
        except RuntimeError as exc:
            return str(exc)

        for (chunk_id, entry_id, chunk_index, text_hash, _text), vec in zip(pending_embed, all_vectors):
            nrm = vec_norm(vec)
            emb_json = json.dumps(vec, separators=(",", ":"))
            conn.execute(
                """
                UPDATE email_chunks
                SET embedding_json = ?, embedding_norm = ?, updated_at = ?
                WHERE entry_id = ? AND chunk_index = ?
                """,
                (emb_json, nrm, now, entry_id, chunk_index),
            )
            _store_chunk_vector(conn, chunk_id, vec, now)
            conn.execute(
                """
                INSERT INTO embedding_cache(text_hash, embedding_json, embedding_norm, updated_at)
                VALUES(?, ?, ?, ?)
                ON CONFLICT(text_hash) DO UPDATE SET
                  embedding_json=excluded.embedding_json,
                  embedding_norm=excluded.embedding_norm,
                  updated_at=excluded.updated_at
                """,
                (text_hash, emb_json, nrm, now),
            )
            stats["embeddedNew"] += 1

    return None


def _execute_sync(
    args: argparse.Namespace,
    plans: list[dict[str, Any]],
    *,
    profile: str,
    full_rebuild: bool,
    clear_embedding_cache: bool,
    continue_on_plan_error: bool,
    resume_state: dict[str, Any] | None = None,
    page_checkpoint: Callable[[dict[str, Any]], None] | None = None,
) -> int:
    conn = ensure_db(args.db)
    embed_enabled = _resolve_embed_enabled(args)
    embed_backend = getattr(args, "embed_backend", DEFAULT_EMBED_BACKEND)
    paged_read = not bool(getattr(args, "disable_paged_read", False))

    # GPU contention probe: only meaningful when we're actually going to use the
    # local GPU backend. Skip quietly for --no-embed or --embed-backend api.
    if embed_enabled and embed_backend == "local" and getattr(args, "gpu_probe", True):
        try:
            contenders = _probe_gpu_contention()
            if contenders:
                _warn_gpu_contention(contenders)
        except Exception as _probe_exc:  # never let the probe block a sync
            print(f"[gpu-probe] skipped due to error: {_probe_exc}", flush=True)

    page_size = max(1, min(int(getattr(args, "page_size", DEFAULT_PAGED_READ_SIZE)), MAX_OUTLOOK_READ_COUNT))

    resume_state = resume_state or {}
    resume_processed = max(0, int(resume_state.get("processed", 0)))
    resume_fetched = max(0, int(resume_state.get("fetched", 0)))
    resume_synced = max(0, int(resume_state.get("synced", 0)))
    resume_skipped = max(0, int(resume_state.get("skipped_existing", 0)))
    resume_pages = max(0, int(resume_state.get("pages", 0)))
    resume_before_iso = str(resume_state.get("before_iso") or "").strip() or None

    sync_started = now_iso()
    total_target = int(sum(max(0, int(p["count"])) for p in plans))
    set_meta(conn, "progress_status", "running")
    set_meta(conn, "progress_profile", profile)
    set_meta(conn, "progress_started_at", sync_started)
    set_meta(conn, "progress_updated_at", sync_started)
    set_meta(conn, "progress_finished_at", "")
    set_meta(conn, "progress_error", "")
    set_meta(conn, "progress_pid", str(os.getpid()))
    set_meta(conn, "progress_total_target", str(total_target))
    set_meta(conn, "progress_done_target", str(resume_processed))
    set_meta(conn, "progress_done_fetched", str(resume_fetched))
    set_meta(conn, "progress_done_synced", str(resume_synced))
    set_meta(conn, "progress_skipped_existing", str(resume_skipped))
    set_meta(conn, "progress_current_folder", "")
    conn.commit()

    # Do not destroy a usable index before Outlook has returned even one
    # successful page. A connectivity/COM failure at the start of a full
    # rebuild must leave the previous index intact.
    rebuild_pending = full_rebuild

    stats: dict[str, Any] = {
        "syncedEmails": 0,
        "syncedChunks": 0,
        "embeddedNew": 0,
        "embeddedCacheHits": 0,
        "dedupMerged": 0,
        "skippedExisting": resume_skipped,
        "attachmentsTextUsed": 0,
        "attachmentsFailed": 0,
    }

    embedding_error = None
    plan_errors: list[dict[str, Any]] = []
    plan_reports: list[dict[str, Any]] = []
    done_target = resume_processed
    done_fetched = resume_fetched
    done_synced = resume_synced

    for plan_index, plan in enumerate(plans):
        set_meta(conn, "progress_current_folder", str(plan["folder"]))
        set_meta(conn, "progress_updated_at", now_iso())
        conn.commit()
        use_resume = plan_index == 0 and resume_processed > 0 and resume_before_iso
        remaining = max(0, int(plan["count"]) - (resume_processed if use_resume else 0))
        plan_fetched = resume_fetched if use_resume else 0
        plan_pages = resume_pages if use_resume else 0
        before_iso: str | None = resume_before_iso if use_resume else None
        plan_failed = False

        while remaining > 0:
            ask_count = min(page_size, remaining) if paged_read else remaining
            try:
                rows = run_outlook_mail_read(
                    ask_count,
                    plan["folder"],
                    plan["hours"],
                    plan["mode"],
                    plan["maxChars"],
                    args.attachments,
                    args.attachments_dir,
                    before_iso=before_iso,
                )
            except RuntimeError as exc:
                err_msg = str(exc)
                set_meta(conn, "progress_done_target", str(done_target))
                set_meta(conn, "progress_updated_at", now_iso())
                conn.commit()
                plan_errors.append({"folder": plan["folder"], "error": err_msg})
                plan_reports.append(
                    {
                        "folder": plan["folder"],
                        "count": plan["count"],
                        "hours": plan["hours"],
                        "mode": plan["mode"],
                        "maxChars": plan["maxChars"],
                        "fetched": plan_fetched,
                        "pages": plan_pages,
                        "error": err_msg,
                    }
                )
                if continue_on_plan_error:
                    plan_failed = True
                    break
                if embedding_error is None:
                    embedding_error = err_msg
                plan_failed = True
                break

            if rebuild_pending:
                _clear_index(conn, drop_embedding_cache=clear_embedding_cache)
                rebuild_pending = False

            fetched_now = len(rows)
            plan_pages += 1
            plan_fetched += fetched_now
            done_target += fetched_now
            done_fetched += fetched_now
            synced_before = stats["syncedEmails"]

            # Skip-existing filter: drop rows whose entry_id or stable_key is
            # already in the DB. Done BETWEEN fetch and ingest so we don't pay
            # for re-chunk, re-embed, or re-upsert on rows we already have.
            # Disabled by --no-skip-existing (e.g. when the user wants to
            # re-process an email whose body has changed).
            #
            # CRITICAL: compute the paging cursor (before_iso) FIRST from the
            # unfiltered rows, then filter. If we filter first and all rows
            # are skipped, _resolve_before_cursor would be called on an empty
            # list and the cursor would never advance — infinite loop on the
            # same page. We solve this by capturing the cursor from the
            # original fetched rows now, then doing the filter.
            if paged_read and rows:
                _next_cursor = _resolve_before_cursor(rows, before_iso)
            else:
                _next_cursor = None

            skip_existing = bool(getattr(args, "skip_existing", True))
            if skip_existing and rows:
                rows, skipped_this_page = _filter_existing_rows(conn, rows)
                stats["skippedExisting"] = stats.get("skippedExisting", 0) + skipped_this_page
                if skipped_this_page:
                    print(
                        f"[skip-existing] {skipped_this_page}/{fetched_now} rows already in DB; "
                        f"{len(rows)} to ingest",
                        flush=True,
                    )

            embedding_error = _ingest_rows(
                conn,
                rows,
                embed_enabled=embed_enabled,
                embed_base_url=args.embed_base_url,
                embed_api_key=args.embed_api_key,
                embed_model=args.embed_model,
                embed_batch=args.embed_batch,
                embed_timeout=args.embed_timeout,
                embed_backend=getattr(args, "embed_backend", DEFAULT_EMBED_BACKEND),
                embed_script=LOCAL_EMBED_SCRIPT,
                stats=stats,
                attachment_max_per_email=args.max_attachments_per_email,
                attachment_max_chars=args.attachment_max_chars,
            )
            done_synced += max(0, int(stats["syncedEmails"]) - int(synced_before))
            set_meta(conn, "progress_done_target", str(done_target))
            set_meta(conn, "progress_done_fetched", str(done_fetched))
            set_meta(conn, "progress_done_synced", str(done_synced))
            set_meta(
                conn,
                "progress_skipped_existing",
                str(max(0, int(stats.get("skippedExisting", 0)))),
            )
            set_meta(conn, "progress_updated_at", now_iso())
            conn.commit()
            if embedding_error:
                plan_failed = True
                break

            remaining = max(0, remaining - fetched_now)
            if page_checkpoint is not None and paged_read and fetched_now > 0:
                page_checkpoint(
                    {
                        "before_iso": _next_cursor,
                        "processed": done_target,
                        "fetched": done_fetched,
                        "synced": done_synced,
                        "skipped_existing": max(
                            0, int(stats.get("skippedExisting", 0))
                        ),
                        "pages": plan_pages,
                    }
                )
            if not paged_read:
                break
            if fetched_now < ask_count:
                break
            # Use the cursor we computed BEFORE the skip-existing filter above
            # (_next_cursor). Re-computing on the filtered `rows` would be
            # wrong because the filter may have removed every row on this
            # page; in that case we still need to move on.
            before_iso = _next_cursor
            if not before_iso:
                break
            # Brief pause between pages to prevent Outlook COM stalls
            time.sleep(PAGE_SLEEP_SEC)

        if not plan_failed:
            plan_reports.append(
                {
                    "folder": plan["folder"],
                    "count": plan["count"],
                    "hours": plan["hours"],
                    "mode": plan["mode"],
                    "maxChars": plan["maxChars"],
                    "fetched": plan_fetched,
                    "pages": plan_pages,
                    "pageSize": page_size if paged_read else int(plan["count"]),
                }
            )

        if embedding_error and not continue_on_plan_error:
            break

    now = now_iso()
    sync_ok = embedding_error is None and len(plan_errors) == 0
    if sync_ok:
        conn.execute(
            "INSERT INTO meta(key, value) VALUES('last_sync_at', ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (now,),
        )
        conn.execute(
            "INSERT INTO meta(key, value) VALUES('last_sync_profile', ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (profile,),
        )
    set_meta(conn, "progress_status", "ok" if sync_ok else "error")
    set_meta(conn, "progress_error", embedding_error or "")
    set_meta(conn, "progress_pid", "")
    set_meta(conn, "progress_current_folder", "")
    set_meta(conn, "progress_finished_at", now_iso())
    set_meta(conn, "progress_updated_at", now_iso())
    conn.commit()

    cleanup_report: dict[str, Any] | None = None
    if args.attachments == "export" and getattr(args, "attachments_cleanup", "none") == "all":
        cleanup_report = _cleanup_exported_attachments(args.attachments_dir)

    out = {
        "ok": sync_ok,
        "profile": profile,
        "fullRebuild": full_rebuild,
        "syncedEmails": stats["syncedEmails"],
        "syncedChunks": stats["syncedChunks"],
        "dedupMerged": stats["dedupMerged"],
        "skippedExisting": stats["skippedExisting"],
        "plans": plan_reports,
        "planErrors": plan_errors,
        "embedding": {
            "enabled": embed_enabled,
            "new": stats["embeddedNew"],
            "cacheHits": stats["embeddedCacheHits"],
            "pending": max(0, stats["syncedChunks"] - (stats["embeddedNew"] + stats["embeddedCacheHits"])),
            "error": embedding_error,
            "model": args.embed_model,
            "baseUrl": args.embed_base_url,
        },
        "attachments": {
            "mode": args.attachments,
            "dir": str(args.attachments_dir),
            "textUsed": stats["attachmentsTextUsed"],
            "extractFailed": stats["attachmentsFailed"],
            "cleanup": cleanup_report
            if cleanup_report is not None
            else {
                "mode": getattr(args, "attachments_cleanup", "none"),
                "dir": str(args.attachments_dir),
            },
        },
    }
    print(json.dumps(out, ensure_ascii=False))
    rc = 0 if (embedding_error is None and len(plan_errors) == 0) else 1
    conn.close()
    return rc


def cmd_sync(args: argparse.Namespace) -> int:
    """Custom sync with explicit range; defaults are "newest only via anchor".

    The default args (count=50, folder=InboxAll, hours=24) are tight by
    design — the user said "default should be newest only, not all". The
    anchor-based widening (same as cmd_sync_default) ensures the window
    automatically grows to catch up if the DB is behind.
    """
    window_hours = max(1, int(args.hours))
    window_count = 1

    conn = ensure_db(args.db)
    now_utc = dt.datetime.now(dt.timezone.utc)

    # Anchor on the newest email already in the DB; fall back to last_sync_at
    # meta so empty DBs still get a sensible first-run window.
    anchor: "dt.datetime | None" = None
    max_row = conn.execute("SELECT MAX(received_time) FROM emails").fetchone()
    if max_row and max_row[0]:
        anchor = parse_utc_timestamp(max_row[0])
    if anchor is None:
        meta_row = conn.execute("SELECT value FROM meta WHERE key='last_sync_at'").fetchone()
        anchor = parse_utc_timestamp(meta_row[0] if meta_row else None)

    if anchor is not None:
        lag_hours = max(0, int(math.ceil((now_utc - anchor).total_seconds() / 3600))) + 2
        if lag_hours > window_hours:
            window_count = int(math.ceil(lag_hours / window_hours))

    window_count = max(1, min(window_count, DEFAULT_DAILY_MAX_WINDOWS))
    effective_hours = window_hours * window_count
    effective_count = max(1, int(args.count) * window_count)

    plans = [
        {
            "count": effective_count,
            "folder": args.folder,
            "hours": effective_hours,
            "mode": args.mode,
            "maxChars": args.max_chars,
        }
    ]
    conn.close()
    return _execute_sync_locked(
        args,
        plans,
        profile="custom",
        full_rebuild=False,
        clear_embedding_cache=False,
        continue_on_plan_error=False,
    )


def cmd_sync_default(args: argparse.Namespace) -> int:
    window_hours = max(1, int(args.hours))
    window_count = 1

    conn = ensure_db(args.db)
    now_utc = dt.datetime.now(dt.timezone.utc)

    # Anchor on the newest email already in the DB — this is the most reliable
    # indicator of what has actually been indexed.  Unlike last_sync_at (which
    # records when the sync script ran, not what it actually captured), the
    # MAX(received_time) in emails guarantees no gap between the last indexed
    # message and the current window.  If the DB is empty we fall back to
    # last_sync_at so that first-run behaviour is unchanged.
    anchor: "dt.datetime | None" = None
    max_row = conn.execute("SELECT MAX(received_time) FROM emails").fetchone()
    if max_row and max_row[0]:
        anchor = parse_utc_timestamp(max_row[0])
    if anchor is None:
        meta_row = conn.execute("SELECT value FROM meta WHERE key='last_sync_at'").fetchone()
        anchor = parse_utc_timestamp(meta_row[0] if meta_row else None)

    if anchor is not None:
        # Add a small overlap buffer so boundary emails are never missed.
        lag_hours = max(0, int(math.ceil((now_utc - anchor).total_seconds() / 3600))) + 2
        if lag_hours > window_hours:
            window_count = int(math.ceil(lag_hours / window_hours))

    window_count = max(1, min(window_count, DEFAULT_DAILY_MAX_WINDOWS))
    effective_hours = window_hours * window_count
    effective_count = max(1, int(args.count) * window_count)

    plans = [
        {
            "count": effective_count,
            "folder": args.folder,
            "hours": effective_hours,
            "mode": args.mode,
            "maxChars": args.max_chars,
        }
    ]
    conn.close()
    return _execute_sync_locked(
        args,
        plans,
        profile="default",
        full_rebuild=False,
        clear_embedding_cache=False,
        continue_on_plan_error=False,
    )


def _parse_folders(raw: str) -> list[str]:
    parts = [x.strip() for x in raw.split(",") if x.strip()]
    if not parts:
        return DEFAULT_FULL_FOLDERS
    return parts


def _parse_hours_windows(raw: str) -> list[int]:
    parts = [x.strip() for x in raw.split(",") if x.strip()]
    out: list[int] = []
    for p in parts:
        try:
            v = int(p)
        except ValueError:
            continue
        if v >= 0:
            out.append(v)
    if not out:
        return list(DEFAULT_FULL_SAFE_HOURS_WINDOWS)
    return out


def _full_safe_signature(folders: list[str], windows: list[int], count_per_step: int, mode: str, max_chars: int) -> str:
    payload = {
        "folders": folders,
        "windows": windows,
        "countPerStep": count_per_step,
        "mode": mode,
        "maxChars": max_chars,
    }
    return sha256_text(json.dumps(payload, sort_keys=True, separators=(",", ":")))


def _build_full_safe_plans(
    folders: list[str],
    windows: list[int],
    count_per_step: int,
    mode: str,
    max_chars: int,
) -> list[dict[str, Any]]:
    plans: list[dict[str, Any]] = []
    for folder in folders:
        for h in windows:
            plans.append(
                {
                    "count": count_per_step,
                    "folder": folder,
                    "hours": h,
                    "mode": mode,
                    "maxChars": max_chars,
                }
            )
    return plans


def cmd_sync_full(args: argparse.Namespace) -> int:
    folders = _parse_folders(args.folders)
    plans = [
        {
            "count": args.count_per_folder,
            "folder": folder,
            "hours": 0,
            "mode": args.mode,
            "maxChars": args.max_chars,
        }
        for folder in folders
    ]
    rebuild = bool(getattr(args, "rebuild", False)) and not bool(
        getattr(args, "no_rebuild", False)
    )
    return _execute_sync_locked(
        args,
        plans,
        profile="full",
        full_rebuild=rebuild,
        clear_embedding_cache=(not args.keep_embedding_cache) if rebuild else False,
        continue_on_plan_error=True,
    )


def _full_safe_worker_command(args: argparse.Namespace) -> list[str]:
    cmd = [
        sys.executable,
        str(Path(__file__).resolve()),
        "--db",
        str(args.db),
        "sync-full-safe",
        "--foreground",
        "--count-per-step",
        str(args.count_per_step),
        "--folders",
        str(args.folders),
        "--hours-windows",
        str(args.hours_windows),
        "--mode",
        str(args.mode),
        "--max-chars",
        str(args.max_chars),
        "--page-size",
        str(args.page_size),
        "--embed-base-url",
        str(args.embed_base_url),
        "--embed-model",
        str(args.embed_model),
        "--embed-api-key",
        str(args.embed_api_key),
        "--embed-batch",
        str(args.embed_batch),
        "--embed-timeout",
        str(args.embed_timeout),
        "--embed-backend",
        str(args.embed_backend),
        "--attachments",
        str(args.attachments),
        "--attachments-dir",
        str(args.attachments_dir),
        "--max-attachments-per-email",
        str(args.max_attachments_per_email),
        "--attachment-max-chars",
        str(args.attachment_max_chars),
        "--attachments-cleanup",
        str(args.attachments_cleanup),
    ]
    if args.from_start:
        cmd.append("--from-start")
    if args.no_embed:
        cmd.append("--no-embed")
    if not args.gpu_probe:
        cmd.append("--no-gpu-probe")
    if args.disable_paged_read:
        cmd.append("--disable-paged-read")
    if not args.skip_existing:
        cmd.append("--no-skip-existing")
    return cmd


def _launch_full_safe_service(args: argparse.Namespace) -> int:
    conn = ensure_db(args.db)
    existing_pid = _safe_int(get_meta(conn, "progress_pid"))
    existing_unit = get_meta(conn, "progress_unit") or ""
    if _email_rag_sync_pid_alive(existing_pid) or (
        existing_unit and _systemd_user_unit_active(existing_unit)
    ):
        print(
            json.dumps(
                {
                    "ok": True,
                    "started": False,
                    "alreadyRunning": True,
                    "pid": existing_pid or None,
                    "unit": existing_unit or None,
                    "log": get_meta(conn, "progress_log") or str(FULL_SAFE_LOG),
                    "hint": "Use `email-rag progress` to monitor it.",
                },
                ensure_ascii=False,
            )
        )
        conn.close()
        return 0

    FULL_SAFE_DATA_DIR.mkdir(parents=True, exist_ok=True)
    try:
        FULL_SAFE_DATA_DIR.chmod(0o700)
    except OSError:
        pass
    FULL_SAFE_LOG.touch(mode=0o600, exist_ok=True)
    try:
        FULL_SAFE_LOG.chmod(0o600)
    except OSError:
        pass

    unit = (
        f"{FULL_SAFE_UNIT_PREFIX}-"
        f"{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}-{os.getpid()}.service"
    )
    worker_cmd = _full_safe_worker_command(args)
    launch_cmd = [
        "systemd-run",
        "--user",
        f"--unit={unit}",
        "--collect",
        "--no-block",
        "--property=Type=exec",
        "--property=KillMode=control-group",
        "--property=TimeoutStartSec=infinity",
        "--property=TimeoutStopSec=30s",
        f"--property=StandardOutput=append:{FULL_SAFE_LOG}",
        f"--property=StandardError=append:{FULL_SAFE_LOG}",
        "--setenv=PYTHONUNBUFFERED=1",
        f"--setenv=PATH={os.environ.get('PATH', os.defpath)}",
        f"--setenv=HOME={Path.home()}",
        *worker_cmd,
    ]
    now = now_iso()
    set_meta(conn, "progress_status", "starting")
    set_meta(conn, "progress_profile", "full-safe")
    set_meta(conn, "progress_started_at", now)
    set_meta(conn, "progress_updated_at", now)
    set_meta(conn, "progress_finished_at", "")
    set_meta(conn, "progress_error", "")
    set_meta(conn, "progress_pid", "")
    set_meta(conn, "progress_unit", unit)
    set_meta(conn, "progress_log", str(FULL_SAFE_LOG))
    conn.commit()
    try:
        result = subprocess.run(
            launch_cmd,
            check=False,
            capture_output=True,
            text=True,
            timeout=15,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError) as exc:
        set_meta(conn, "progress_status", "error")
        set_meta(conn, "progress_error", str(exc))
        set_meta(conn, "progress_finished_at", now_iso())
        set_meta(conn, "progress_updated_at", now_iso())
        conn.commit()
        print(
            json.dumps(
                {
                    "ok": False,
                    "error": f"Unable to submit durable systemd user service: {exc}",
                    "hint": "Check `systemctl --user is-system-running`; do not run a full sync through a time-limited agent exec.",
                },
                ensure_ascii=False,
            )
        )
        conn.close()
        return 1
    if result.returncode != 0:
        error_text = (result.stderr or result.stdout).strip() or (
            f"systemd-run exited {result.returncode}"
        )
        set_meta(conn, "progress_status", "error")
        set_meta(conn, "progress_error", error_text)
        set_meta(conn, "progress_finished_at", now_iso())
        set_meta(conn, "progress_updated_at", now_iso())
        conn.commit()
        print(
            json.dumps(
                {
                    "ok": False,
                    "error": error_text,
                },
                ensure_ascii=False,
            )
        )
        conn.close()
        return 1

    conn.close()

    active = False
    for _ in range(20):
        if _systemd_user_unit_active(unit):
            active = True
            break
        time.sleep(0.1)
    print(
        json.dumps(
            {
                "ok": True,
                "started": True,
                "durable": True,
                "unit": unit,
                "active": active,
                "log": str(FULL_SAFE_LOG),
                "hint": "The job survives agent/exec timeouts. Use `email-rag progress` to monitor it.",
            },
            ensure_ascii=False,
        )
    )
    return 0


def cmd_sync_full_safe(args: argparse.Namespace) -> int:
    if not bool(getattr(args, "foreground", False)) and not bool(
        getattr(args, "dry_run", False)
    ):
        return _launch_full_safe_service(args)

    lock_fd = _acquire_full_safe_lock()
    if lock_fd is None:
        print(
            json.dumps(
                {
                    "ok": False,
                    "error": "Another sync-full-safe worker already holds the lock.",
                    "hint": "Use `email-rag progress`; do not start a second writer.",
                },
                ensure_ascii=False,
            )
        )
        return 3

    previous_handlers: dict[int, Any] = {}

    def _interrupt(signum: int, _frame: Any) -> None:
        raise SyncInterrupted(f"sync-full-safe interrupted by signal {signum}")

    for signum in (signal.SIGTERM, signal.SIGINT):
        previous_handlers[signum] = signal.getsignal(signum)
        signal.signal(signum, _interrupt)
    try:
        return _cmd_sync_full_safe_foreground(args)
    except SyncInterrupted as exc:
        _mark_sync_interrupted(args.db, str(exc))
        print(
            json.dumps(
                {
                    "ok": False,
                    "interrupted": True,
                    "error": str(exc),
                    "hint": "Run `email-rag sync-full-safe` again; it will resume from the last committed page.",
                },
                ensure_ascii=False,
            ),
            flush=True,
        )
        return 130
    finally:
        for signum, handler in previous_handlers.items():
            signal.signal(signum, handler)
        _release_full_safe_lock(lock_fd)


def _cmd_sync_full_safe_foreground(args: argparse.Namespace) -> int:
    conn = ensure_db(args.db)
    folders = _parse_folders(args.folders)
    windows = _parse_hours_windows(args.hours_windows)
    plans = _build_full_safe_plans(
        folders=folders,
        windows=windows,
        count_per_step=max(1, int(args.count_per_step)),
        mode=args.mode,
        max_chars=args.max_chars,
    )

    # Optimization: sort so the broadest hours-window runs first within each folder.
    # The default windows [24, 72, 168, ...] are cumulative ("last N hours"), so each
    # smaller window is fully contained in the next bigger one. After the biggest
    # window finishes, every smaller window for the same folder is a no-op (its
    # emails are already in DB). With 2 folders × 8 windows, this turns 16 active
    # plans into 2 active plans + 14 free skips — saving the bulk of the 1s/page
    # sleep + Outlook re-read cost (the actual bottleneck per sync logs).
    def coverage_rank(hours: int) -> float:
        return math.inf if hours == 0 else float(hours)

    plans.sort(key=lambda p: (p["folder"], -coverage_rank(int(p["hours"]))))

    total = len(plans)
    if total == 0:
        print(json.dumps({"ok": False, "error": "No plans generated for sync-full-safe."}, ensure_ascii=False))
        conn.close()
        return 1

    sig = _full_safe_signature(
        folders=folders,
        windows=windows,
        count_per_step=max(1, int(args.count_per_step)),
        mode=args.mode,
        max_chars=args.max_chars,
    )
    sig_key = "full_safe.signature"
    idx_key = "full_safe.index"
    total_key = "full_safe.total"
    status_key = "full_safe.status"
    updated_key = "full_safe.updated_at"
    # Plan order identifier — bump this string whenever the plan execution order
    # changes. Old checkpoints with a mismatching value are auto-discarded (start
    # over from step 0) so users never resume into the wrong plan.
    order_key = "full_safe.plan_order"
    current_plan_order = "global-stable-keyset-v4"

    prev_sig = get_meta(conn, sig_key)
    prev_order = get_meta(conn, order_key)
    start_idx = 0
    if not args.from_start and prev_sig == sig and prev_order == current_plan_order:
        start_idx = int(get_meta(conn, idx_key) or "0")
        if start_idx < 0:
            start_idx = 0
        if start_idx > total:
            start_idx = total
    elif not args.from_start and prev_sig == sig and prev_order and prev_order != current_plan_order:
        # Old checkpoint with old order — auto-restart to avoid resuming into the
        # wrong plan index. Print a hint so the user understands the implicit reset.
        print(
            f"[plan-order] checkpoint plan_order='{prev_order}' doesn't match current "
            f"'{current_plan_order}'; restarting from step 0 (safe auto-reset).",
            flush=True,
        )

    # Dry-run: print which plans will run vs be skipped, then exit. Does NOT
    # touch checkpoint state, so it's safe to run repeatedly for preview.
    if getattr(args, "dry_run", False):
        rc = _full_safe_dry_run(plans, start_idx)
        conn.close()
        return rc

    if args.from_start or prev_sig != sig or prev_order != current_plan_order:
        _clear_full_safe_page_checkpoint(conn)
    set_meta(conn, sig_key, sig)
    set_meta(conn, order_key, current_plan_order)
    set_meta(conn, total_key, str(total))
    set_meta(conn, idx_key, str(start_idx))
    set_meta(conn, status_key, "running")
    set_meta(conn, updated_key, now_iso())
    conn.commit()

    done = start_idx
    failed_step: dict[str, Any] | None = None

    # Skip-optimization state: for each folder, the max hours already covered.
    # If covered >= current plan hours, the current plan is redundant.
    covered_max_hours: dict[str, float] = {}
    # Re-derive from completed plans so resume behaves correctly.
    for j in range(start_idx):
        s = plans[j]
        covered_max_hours[s["folder"]] = max(
            covered_max_hours.get(s["folder"], 0), coverage_rank(int(s["hours"]))
        )

    for i in range(start_idx, total):
        step = plans[i]
        folder = step["folder"]
        hours = int(step["hours"])
        hours_coverage = coverage_rank(hours)

        if covered_max_hours.get(folder, 0) >= hours_coverage:
            # Fully covered by an already-completed bigger window — skip without
            # touching Outlook or embed. The checkpoint still advances so resume
            # picks up correctly.
            done = i + 1
            set_meta(conn, idx_key, str(done))
            set_meta(conn, status_key, "running")
            set_meta(conn, updated_key, now_iso())
            conn.commit()
            print(
                f"[skip] step {i+1}/{total} {folder}/{hours}h "
                f"covered by {covered_max_hours[folder]}h plan",
                flush=True,
            )
            continue

        resume_state = None
        if not args.from_start:
            resume_state = _load_full_safe_page_checkpoint(
                conn,
                signature=sig,
                step_index=i,
            )
        if resume_state:
            print(
                f"[resume-page] step {i+1}/{total} {folder}/{hours}h "
                f"processed={resume_state['processed']} pages={resume_state['pages']} "
                f"before={resume_state['before_iso']}",
                flush=True,
            )

        def _save_page(state: dict[str, Any], *, _step_index: int = i) -> None:
            _write_full_safe_page_checkpoint(
                conn,
                signature=sig,
                step_index=_step_index,
                state=state,
            )

        rc = _execute_sync(
            args,
            [step],
            profile="full-safe",
            # A resumable "safe" sync must never destroy the usable index.
            # --from-start resets only the checkpoint; it does not clear data.
            full_rebuild=False,
            clear_embedding_cache=False,
            continue_on_plan_error=False,
            resume_state=resume_state,
            page_checkpoint=_save_page,
        )
        if rc != 0:
            failed_step = {
                "index": i,
                "folder": step["folder"],
                "hours": step["hours"],
                "count": step["count"],
            }
            break

        # Mark this folder as covered up to `hours`. Smaller subsequent plans
        # for the same folder will hit the skip branch above.
        covered_max_hours[folder] = max(
            covered_max_hours.get(folder, 0), hours_coverage
        )

        done = i + 1
        _clear_full_safe_page_checkpoint(conn)
        set_meta(conn, idx_key, str(done))
        set_meta(conn, status_key, "running")
        set_meta(conn, updated_key, now_iso())
        conn.commit()

    if failed_step is not None:
        set_meta(conn, status_key, "error")
        set_meta(conn, updated_key, now_iso())
        conn.commit()
        print(
            json.dumps(
                {
                    "ok": False,
                    "profile": "full-safe",
                    "resumeIndex": done,
                    "totalSteps": total,
                    "failedStep": failed_step,
                    "hint": "Run sync-full-safe again to resume from resumeIndex.",
                },
                ensure_ascii=False,
            )
        )
        conn.close()
        return 1

    set_meta(conn, status_key, "complete")
    set_meta(conn, idx_key, str(total))
    set_meta(conn, updated_key, now_iso())
    set_meta(conn, "progress_status", "ok")
    set_meta(conn, "progress_pid", "")
    set_meta(conn, "progress_current_folder", "")
    set_meta(conn, "progress_finished_at", now_iso())
    set_meta(conn, "progress_updated_at", now_iso())
    _clear_full_safe_page_checkpoint(conn)
    conn.commit()
    print(
        json.dumps(
            {
                "ok": True,
                "profile": "full-safe",
                "resumeIndex": total,
                "totalSteps": total,
                "completed": True,
            },
            ensure_ascii=False,
        )
    )
    conn.close()
    return 0


def _full_safe_dry_run(plans: list[dict[str, Any]], start_idx: int) -> int:
    """Print which plans will run vs be skipped under broadest-first order.

    Does NOT modify checkpoint state. Safe to run repeatedly.
    """
    total = len(plans)
    def coverage_rank(hours: int) -> float:
        return math.inf if hours == 0 else float(hours)

    covered_max_hours: dict[str, float] = {}
    run_count = 0
    skip_count = 0
    already_done = start_idx

    print(f"sync-full-safe plan preview  (order: broadest-first, total={total}, already_done={already_done})", flush=True)
    print(f"{'idx':>4}  {'status':<7}  {'folder':<12}  {'hours':>8}  {'count':>6}", flush=True)
    print("-" * 48, flush=True)
    for i, step in enumerate(plans):
        folder = step["folder"]
        hours = int(step["hours"])
        if i < start_idx:
            status = "DONE"
            covered_max_hours[folder] = max(
                covered_max_hours.get(folder, 0), coverage_rank(hours)
            )
        elif covered_max_hours.get(folder, 0) >= coverage_rank(hours):
            status = "SKIP"
            skip_count += 1
        else:
            status = "RUN"
            run_count += 1
            covered_max_hours[folder] = max(
                covered_max_hours.get(folder, 0), coverage_rank(hours)
            )
        print(f"{i:>4}  {status:<7}  {folder:<12}  {hours:>8}  {step['count']:>6}", flush=True)
    print("-" * 48, flush=True)
    print(f"Summary: {run_count} will RUN, {skip_count} will SKIP, {already_done} already DONE.", flush=True)
    return 0


def cmd_cleanup(args: argparse.Namespace) -> int:
    dirs = args.dirs or [DEFAULT_ATTACHMENT_EXPORT_DIR, DEFAULT_ATTACHMENT_TEST_DIR]
    reports: list[dict[str, Any]] = []
    for d in dirs:
        reports.append(_cleanup_exported_attachments(Path(d)))
    ok = not any("error" in report for report in reports)
    print(json.dumps({"ok": ok, "reports": reports}, ensure_ascii=False))
    return 0 if ok else 2


def cmd_cleanup_daemon(args: argparse.Namespace) -> int:
    """Clean up stale Outlook daemon state.

    Default mode (no flag) is SAFE — only deletes stale schtasks and old
    ocl_daemon_*.bat files. The live daemon is left alone (so any in-flight
    pipe requests continue to work; the next email-rag sync won't even need
    to relaunch it).

    Use --kill-daemon if the daemon is hung, the pipe is wedged, or you want
    a totally clean slate. The next sync will relaunch the daemon.
    """
    kill_daemon = bool(getattr(args, "kill_daemon", False))
    _kill_stale_outlook_state(kill_procs=kill_daemon, rotate_log=kill_daemon)
    print(
        json.dumps(
            {
                "ok": True,
                "killedLiveDaemon": kill_daemon,
                "rotatedLog": kill_daemon,
                "nextSyncWillRelaunchDaemon": kill_daemon,
            },
            ensure_ascii=False,
        )
    )
    return 0


def cmd_kill_embed_server(args: argparse.Namespace) -> int:
    """Stop the persistent local embed server.

    Idempotent: a no-op if the server isn't running. The next search will
    respawn it (paying the ~26s model load cost again).
    """
    was_alive = _embed_server_alive()
    killed = False
    if was_alive:
        # Polite path: try the shutdown JSON command first.
        try:
            with socket.create_connection(
                (EMBED_SERVER_HOST, EMBED_SERVER_PORT), timeout=2.0
            ) as sock:
                sock.sendall((json.dumps({"cmd": "shutdown"}) + "\n").encode("utf-8"))
        except Exception:
            pass
    # Force path also handles an older email-rag server that does not yet
    # implement the ping protocol. The PID identity guard prevents killing an
    # unrelated process if the pidfile is stale.
    killed = _kill_embed_server()
    # Wait up to 5s for the port to release.
    for _ in range(50):
        if not _embed_server_alive():
            break
        time.sleep(0.1)
    print(
        json.dumps(
            {
                "ok": True,
                "wasAlive": was_alive,
                "killed": killed,
                "stillAlive": _embed_server_alive(),
            },
            ensure_ascii=False,
        )
    )
    return 0


def cmd_estimate(args: argparse.Namespace) -> int:
    conn = ensure_db(args.db)
    folders = _parse_folders(args.folders)
    estimate = estimate_outlook_folder_counts(folders)
    rag_emails = conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0]
    out = {
        "ok": True,
        "folders": estimate.get("folders", {}),
        "outlookTotal": int(estimate.get("total", 0)),
        "ragIndexedEmails": int(rag_emails),
        "gapEstimate": max(0, int(estimate.get("total", 0)) - int(rag_emails)),
        "generatedAt": estimate.get("generatedAt"),
    }
    print(json.dumps(out, ensure_ascii=False))
    conn.close()
    return 0


def cmd_progress(args: argparse.Namespace) -> int:
    conn = ensure_db(args.db)
    email_count = conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0]
    chunk_count = conn.execute("SELECT COUNT(*) FROM email_chunks").fetchone()[0]

    status = get_meta(conn, "progress_status") or "idle"
    profile = get_meta(conn, "progress_profile")
    started_at = get_meta(conn, "progress_started_at")
    updated_at = get_meta(conn, "progress_updated_at")
    finished_at = get_meta(conn, "progress_finished_at")
    current_folder = get_meta(conn, "progress_current_folder")
    total_target = int(get_meta(conn, "progress_total_target") or "0")
    done_target = int(get_meta(conn, "progress_done_target") or "0")
    done_fetched = int(get_meta(conn, "progress_done_fetched") or "0")
    done_synced = int(get_meta(conn, "progress_done_synced") or "0")
    skipped_existing = int(get_meta(conn, "progress_skipped_existing") or "0")
    progress_error = get_meta(conn, "progress_error")
    progress_pid = _safe_int(get_meta(conn, "progress_pid"))
    progress_unit = get_meta(conn, "progress_unit") or ""
    progress_log = get_meta(conn, "progress_log") or ""
    safe_idx = int(get_meta(conn, "full_safe.index") or "0")
    safe_total = int(get_meta(conn, "full_safe.total") or "0")
    safe_status = get_meta(conn, "full_safe.status")
    page_step = _safe_int(get_meta(conn, "full_safe.page.step_index"), -1)
    page_cursor = get_meta(conn, "full_safe.page.before_iso")
    page_processed = _safe_int(get_meta(conn, "full_safe.page.processed"))
    page_pages = _safe_int(get_meta(conn, "full_safe.page.pages"))
    started_dt = parse_utc_timestamp(started_at)
    updated_dt = parse_utc_timestamp(updated_at)
    finished_dt = parse_utc_timestamp(finished_at)
    now_dt = dt.datetime.now(dt.timezone.utc)
    elapsed_end_dt = (
        now_dt
        if status in {"starting", "running"}
        else (finished_dt or updated_dt or now_dt)
    )
    elapsed_sec = (
        int((elapsed_end_dt - started_dt).total_seconds()) if started_dt else 0
    )
    stale_sec = int((now_dt - updated_dt).total_seconds()) if updated_dt else 0

    pct = 0.0
    if total_target > 0:
        pct = min(100.0, (done_target / total_target) * 100.0)
    pid_alive = _email_rag_sync_pid_alive(progress_pid)
    unit_active = bool(progress_unit) and _systemd_user_unit_active(progress_unit)
    process_alive = pid_alive or unit_active
    effective_status = status
    if status == "running" and not process_alive:
        effective_status = "interrupted"
        status = "interrupted"
        safe_status = "interrupted"
        reconciled_at = now_iso()
        finished_at = reconciled_at
        updated_at = reconciled_at
        current_folder = ""
        progress_pid = 0
        stale_sec = 0
        reconciled_dt = parse_utc_timestamp(reconciled_at) or now_dt
        elapsed_sec = (
            int((reconciled_dt - started_dt).total_seconds()) if started_dt else 0
        )
        set_meta(conn, "progress_status", "interrupted")
        set_meta(conn, "progress_error", "Worker process is no longer alive.")
        set_meta(conn, "progress_pid", "")
        set_meta(conn, "progress_current_folder", "")
        set_meta(conn, "progress_finished_at", finished_at)
        set_meta(conn, "progress_updated_at", updated_at)
        set_meta(conn, "full_safe.status", "interrupted")
        set_meta(conn, "full_safe.updated_at", updated_at)
        conn.commit()
        progress_error = "Worker process is no longer alive."
    elif status == "starting" and not process_alive and stale_sec >= 30:
        effective_status = "error"
        progress_error = progress_error or "Durable worker did not become active."
    elif status == "running" and stale_sec >= 180:
        # A live process may legitimately spend several minutes inside one
        # Outlook COM or embedding call. Report the heartbeat lag without
        # falsely declaring the live worker stale.
        effective_status = "running"
    if status in {"ok", "error", "interrupted"} and current_folder:
        current_folder = ""
        set_meta(conn, "progress_current_folder", "")
        conn.commit()

    out = {
        "ok": True,
        "status": effective_status,
        "rawStatus": status,
        "profile": profile,
        "startedAt": started_at,
        "updatedAt": updated_at,
        "finishedAt": finished_at,
        "currentFolder": current_folder,
        "targetEmailsPlanned": total_target,
        "plannedProcessed": done_target,
        "fetchedSoFar": done_fetched,
        "syncedSoFar": done_synced,
        "skippedExisting": skipped_existing,
        "progressPctByPlan": round(pct, 2),
        "elapsedSec": elapsed_sec,
        "staleSec": stale_sec,
        "worker": {
            "pid": progress_pid or None,
            "pidAlive": pid_alive,
            "unit": progress_unit or None,
            "unitActive": unit_active,
            "alive": process_alive,
            "log": progress_log or None,
        },
        "error": progress_error or None,
        "safeCheckpoint": {
            "status": safe_status,
            "stepIndex": safe_idx,
            "stepTotal": safe_total,
            "pageStepIndex": page_step if page_step >= 0 else None,
            "pageCursor": page_cursor or None,
            "pageProcessed": page_processed,
            "pageCount": page_pages,
        },
        "ragIndexedEmails": int(email_count),
        "ragChunks": int(chunk_count),
    }
    print(json.dumps(out, ensure_ascii=False))
    conn.close()
    return 0


def _fts_query(q: str) -> str:
    tokens = [t for t in re.split(r"\s+", q.strip()) if t]
    if not tokens:
        return '""'
    escaped = [f'"{t.replace(chr(34), "")}"*' for t in tokens[:14]]
    return " OR ".join(escaped)


def bm25_candidates(conn: sqlite3.Connection, query: str, limit: int) -> list[dict[str, Any]]:
    fts_q = _fts_query(query)
    try:
        cur = conn.execute(
            """
            SELECT
              chunk_id,
              entry_id,
              subject,
              sender,
              received_time,
              snippet(email_chunks_fts, 5, '[', ']', ' ... ', 22) AS snippet,
              bm25(email_chunks_fts) AS score
            FROM email_chunks_fts
            WHERE email_chunks_fts MATCH ?
            ORDER BY score
            LIMIT ?
            """,
            (fts_q, limit),
        )
    except sqlite3.OperationalError:
        return []

    out: list[dict[str, Any]] = []
    for row in cur.fetchall():
        out.append(
            {
                "chunkId": int(row[0]),
                "entryId": row[1],
                "subject": row[2],
                "sender": row[3],
                "receivedTime": row[4],
                "snippet": row[5],
                "bm25": float(row[6]) if row[6] is not None else 999.0,
            }
        )
    return out


def _vector_to_f32_blob(vec: list[float]) -> bytes:
    if np is not None:
        return np.asarray(vec, dtype="<f4").tobytes()
    import array

    values = array.array("f", (float(x) for x in vec))
    if sys.byteorder != "little":
        values.byteswap()
    return values.tobytes()


def _store_chunk_vector(
    conn: sqlite3.Connection,
    chunk_id: int,
    vec: list[float],
    updated_at: str,
) -> None:
    conn.execute(
        """
        INSERT INTO email_chunk_vectors(chunk_id, dimensions, embedding_f32, updated_at)
        VALUES(?, ?, ?, ?)
        ON CONFLICT(chunk_id) DO UPDATE SET
          dimensions=excluded.dimensions,
          embedding_f32=excluded.embedding_f32,
          updated_at=excluded.updated_at
        """,
        (chunk_id, len(vec), _vector_to_f32_blob(vec), updated_at),
    )


def _store_chunk_vector_json(
    conn: sqlite3.Connection,
    chunk_id: int,
    embedding_json: str,
    updated_at: str,
) -> bool:
    try:
        vec = [float(x) for x in json.loads(embedding_json)]
    except (TypeError, ValueError, json.JSONDecodeError):
        return False
    _store_chunk_vector(conn, chunk_id, vec, updated_at)
    return True


def _backfill_chunk_vectors(conn: sqlite3.Connection, batch_size: int = 500) -> int:
    """Populate compact vectors once for legacy rows that only have JSON."""
    total = 0
    cur = conn.execute(
        """
        SELECT c.chunk_id, c.embedding_json, c.updated_at
        FROM email_chunks AS c
        LEFT JOIN email_chunk_vectors AS v ON v.chunk_id = c.chunk_id
        WHERE c.embedding_json IS NOT NULL AND v.chunk_id IS NULL
        """
    )
    while True:
        rows = cur.fetchmany(batch_size)
        if not rows:
            break
        for chunk_id, embedding_json, updated_at in rows:
            if _store_chunk_vector_json(
                conn, int(chunk_id), embedding_json, updated_at or now_iso()
            ):
                total += 1
        conn.commit()
    return total


def _embed_query_local(query: str) -> list[float]:
    """Embed one search query using the reliable default local strategy.

    On this WSL2 host, a newly spawned persistent CUDA server can wedge on its
    first native request even though the equivalent one-shot process completes
    in a few seconds.  Prefer the deterministic one-shot path for interactive
    search.  The warm persistent path remains available as an explicit opt-in
    via EMAIL_RAG_QUERY_EMBED_MODE=persistent.
    """
    if LOCAL_QUERY_EMBED_MODE == "persistent":
        return _embed_server_request([query])[0]
    errors: list[str] = []
    for _attempt in range(LOCAL_QUERY_EMBED_ATTEMPTS):
        try:
            return embed_texts_local(
                [query],
                LOCAL_EMBED_SCRIPT,
                timeout_sec=max(1, int(LOCAL_QUERY_EMBED_TIMEOUT)),
            )[0]
        except RuntimeError as exc:
            errors.append(str(exc))
    raise RuntimeError(
        f"Local query embedding failed after {LOCAL_QUERY_EMBED_ATTEMPTS} "
        f"bounded CUDA attempts: {'; '.join(errors)}"
    )


def vector_candidates(
    conn: sqlite3.Connection,
    query: str,
    top_n: int,
    embed_base_url: str,
    embed_api_key: str,
    embed_model: str,
    embed_timeout: int,
    embed_backend: str = DEFAULT_EMBED_BACKEND,
    embed_script: Path = LOCAL_EMBED_SCRIPT,
) -> list[dict[str, Any]]:
    if embed_backend == "local":
        qvec = _embed_query_local(query)
    else:
        qvec = embed_texts(embed_base_url, embed_api_key, embed_model, [query], embed_timeout)[0]
    qn = vec_norm(qvec)
    if qn == 0 or not qvec:
        return []

    _backfill_chunk_vectors(conn)
    cur = conn.execute(
        """
        SELECT c.chunk_id, c.entry_id, c.subject, c.sender, c.received_time,
               c.chunk_text, v.embedding_f32, v.dimensions, c.embedding_norm
        FROM email_chunks AS c
        JOIN email_chunk_vectors AS v ON v.chunk_id = c.chunk_id
        WHERE c.embedding_json IS NOT NULL
        """
    )

    # Keep only the global top-N while reading in batches.  This bounds memory
    # even when the mailbox grows to hundreds of thousands of chunks.
    best: list[tuple[float, int, tuple[Any, ...]]] = []
    query_dim = len(qvec)
    qarr = np.asarray(qvec, dtype=np.float32) if np is not None else None

    while True:
        rows = cur.fetchmany(1024)
        if not rows:
            break
        valid = [
            row
            for row in rows
            if row[6] and int(row[7]) == query_dim
            and len(row[6]) == query_dim * 4
        ]
        if not valid:
            continue

        if np is not None:
            matrix = np.empty((len(valid), query_dim), dtype=np.float32)
            for idx, row in enumerate(valid):
                matrix[idx] = np.frombuffer(row[6], dtype="<f4", count=query_dim)
            norms = np.asarray(
                [float(row[8]) if row[8] else 0.0 for row in valid],
                dtype=np.float32,
            )
            dots = matrix @ qarr
            denom = norms * np.float32(qn)
            scores = np.divide(
                dots,
                denom,
                out=np.zeros_like(dots),
                where=denom != 0,
            )
            score_values = scores.tolist()
        else:
            import array

            score_values = []
            for row in valid:
                dvec = array.array("f")
                dvec.frombytes(row[6])
                if sys.byteorder != "little":
                    dvec.byteswap()
                dn = float(row[8]) if row[8] else vec_norm(dvec)
                score_values.append(cosine(qvec, dvec, qn, dn))

        for row, sim in zip(valid, score_values):
            item = (float(sim), int(row[0]), row)
            if len(best) < top_n:
                heapq.heappush(best, item)
            elif sim > best[0][0]:
                heapq.heapreplace(best, item)

    out: list[dict[str, Any]] = []
    for sim, chunk_id, row in sorted(best, reverse=True):
        out.append(
            {
                "chunkId": chunk_id,
                "entryId": row[1],
                "subject": row[2],
                "sender": row[3],
                "receivedTime": row[4],
                "snippet": (row[5] or "")[:260],
                "vector": sim,
            }
        )
    return out


def aggregate_hits(
    bm25_hits: list[dict[str, Any]],
    vec_hits: list[dict[str, Any]],
    top_k: int,
    text_weight: float,
    vector_weight: float,
    mode: str,
) -> list[dict[str, Any]]:
    by_entry: dict[str, dict[str, Any]] = {}

    def bm25_to_rel(score: float) -> float:
        return 1.0 / (1.0 + max(0.0, score))

    for h in bm25_hits:
        eid = h["entryId"]
        rec = by_entry.setdefault(
            eid,
            {
                "entryId": eid,
                "subject": h.get("subject"),
                "sender": h.get("sender"),
                "receivedTime": h.get("receivedTime"),
                "snippet": h.get("snippet"),
                "bm25": None,
                "vector": None,
            },
        )
        cur = rec.get("bm25")
        if cur is None or h["bm25"] < cur:
            rec["bm25"] = h["bm25"]
            rec["snippet"] = h.get("snippet") or rec.get("snippet")

    for h in vec_hits:
        eid = h["entryId"]
        rec = by_entry.setdefault(
            eid,
            {
                "entryId": eid,
                "subject": h.get("subject"),
                "sender": h.get("sender"),
                "receivedTime": h.get("receivedTime"),
                "snippet": h.get("snippet"),
                "bm25": None,
                "vector": None,
            },
        )
        cur = rec.get("vector")
        if cur is None or h["vector"] > cur:
            rec["vector"] = h["vector"]
            if not rec.get("snippet"):
                rec["snippet"] = h.get("snippet")

    ranked: list[dict[str, Any]] = []
    for rec in by_entry.values():
        bm25_rel = bm25_to_rel(rec["bm25"]) if rec["bm25"] is not None else 0.0
        vec_rel = max(0.0, float(rec["vector"])) if rec["vector"] is not None else 0.0

        if mode == "bm25":
            final = bm25_rel
        elif mode == "vector":
            final = vec_rel
        else:
            final = text_weight * bm25_rel + vector_weight * vec_rel

        ranked.append(
            {
                "entryId": rec["entryId"],
                "subject": rec["subject"],
                "sender": rec["sender"],
                "receivedTime": rec["receivedTime"],
                "snippet": rec["snippet"],
                "score": final,
                "bm25": rec["bm25"],
                "vector": rec["vector"],
                "openCommand": make_open_command(str(rec["entryId"])),
            }
        )

    ranked.sort(key=lambda x: x["score"], reverse=True)
    return ranked[:top_k]


def cmd_search(args: argparse.Namespace) -> int:
    conn = ensure_db(args.db)

    mode = args.mode
    bm25_hits: list[dict[str, Any]] = []
    vec_hits: list[dict[str, Any]] = []
    embedding_error = None

    candidate_n = max(args.top_k * max(1, args.candidate_multiplier), args.top_k)

    if mode in ("bm25", "hybrid"):
        bm25_hits = bm25_candidates(conn, args.query, candidate_n)

    if mode in ("vector", "hybrid"):
        try:
            vec_hits = vector_candidates(
                conn,
                args.query,
                candidate_n,
                args.embed_base_url,
                args.embed_api_key,
                args.embed_model,
                args.embed_timeout,
                embed_backend=getattr(args, "embed_backend", DEFAULT_EMBED_BACKEND),
                embed_script=LOCAL_EMBED_SCRIPT,
            )
        except RuntimeError as exc:
            embedding_error = str(exc)
            if mode == "vector":
                print(json.dumps({"ok": False, "error": embedding_error}, ensure_ascii=False))
                return 1
            mode = "bm25"

    hits = aggregate_hits(
        bm25_hits=bm25_hits,
        vec_hits=vec_hits,
        top_k=args.top_k,
        text_weight=args.text_weight,
        vector_weight=args.vector_weight,
        mode=mode,
    )

    print(
        json.dumps(
            {
                "ok": True,
                "query": args.query,
                "mode": mode,
                "weights": {"text": args.text_weight, "vector": args.vector_weight},
                "embeddingError": embedding_error,
                "hits": hits,
            },
            ensure_ascii=False,
        )
    )
    return 0


def cmd_status(args: argparse.Namespace) -> int:
    conn = ensure_db(args.db)
    email_count = conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0]
    chunk_count = conn.execute("SELECT COUNT(*) FROM email_chunks").fetchone()[0]
    chunk_embedded = conn.execute("SELECT COUNT(*) FROM email_chunks WHERE embedding_json IS NOT NULL").fetchone()[0]
    cache_count = conn.execute("SELECT COUNT(*) FROM embedding_cache").fetchone()[0]
    row = conn.execute("SELECT value FROM meta WHERE key='last_sync_at'").fetchone()
    profile_row = conn.execute("SELECT value FROM meta WHERE key='last_sync_profile'").fetchone()

    print(
        json.dumps(
            {
                "ok": True,
                "db": str(args.db),
                "emails": email_count,
                "chunks": chunk_count,
                "chunksEmbedded": chunk_embedded,
                "embeddingCache": cache_count,
                "lastSyncAt": row[0] if row else None,
                "lastSyncProfile": profile_row[0] if profile_row else None,
            },
            ensure_ascii=False,
        )
    )
    return 0


def cmd_embed_pending(args: argparse.Namespace) -> int:
    """Embed all chunks that have embedding_json IS NULL (no Outlook re-fetch needed)."""
    conn = ensure_db(args.db)
    embed_base_url = args.embed_base_url
    embed_api_key = args.embed_api_key
    embed_model = args.embed_model
    embed_batch = max(1, args.embed_batch)
    embed_timeout = args.embed_timeout
    embed_backend = getattr(args, "embed_backend", DEFAULT_EMBED_BACKEND)

    pending = conn.execute(
        "SELECT rowid, entry_id, chunk_index, text_hash, chunk_text FROM email_chunks WHERE embedding_json IS NULL"
    ).fetchall()

    total = len(pending)
    done = 0
    errors = 0
    print(json.dumps({"ok": True, "pending": total, "backend": embed_backend}), flush=True)

    if embed_backend == "local":
        # Use persistent server mode: model loads once, texts sent in batches of LOCAL_EMBED_BATCH_SIZE.
        all_texts = [row[4] for row in pending]

        def _progress(done_n: int, total_n: int) -> None:
            if done_n % (LOCAL_EMBED_BATCH_SIZE * 16) == 0 or done_n == total_n:
                print(json.dumps({"ok": True, "done": done_n, "total": total_n, "errors": 0}), flush=True)

        try:
            all_vectors = embed_texts_local_persistent(all_texts, LOCAL_EMBED_SCRIPT, progress_cb=_progress)
            if len(all_vectors) != len(pending):
                raise RuntimeError(f"count mismatch: expected {len(pending)}, got {len(all_vectors)}")
        except RuntimeError as exc:
            print(json.dumps({"ok": False, "error": str(exc), "done": 0, "remaining": total}), flush=True)
            return 1

        now = now_iso()
        for row, vec in zip(pending, all_vectors):
            rowid, entry_id, chunk_index, text_hash, text = row
            nrm = vec_norm(vec)
            emb_json = json.dumps(vec, separators=(",", ":"))
            conn.execute(
                "UPDATE email_chunks SET embedding_json=?, embedding_norm=?, updated_at=? WHERE rowid=?",
                (emb_json, nrm, now, rowid),
            )
            _store_chunk_vector(conn, int(rowid), vec, now)
            conn.execute(
                """
                INSERT INTO embedding_cache(text_hash, embedding_json, embedding_norm, updated_at)
                VALUES(?, ?, ?, ?)
                ON CONFLICT(text_hash) DO UPDATE SET
                  embedding_json=excluded.embedding_json,
                  embedding_norm=excluded.embedding_norm,
                  updated_at=excluded.updated_at
                """,
                (text_hash, emb_json, nrm, now),
            )
            done += 1
            if done % 500 == 0 or done == total:
                conn.commit()
                print(json.dumps({"ok": True, "done": done, "total": total, "errors": errors}), flush=True)
        conn.commit()
    else:
        # API backend: batch loop.
        for i in range(0, total, embed_batch):
            batch = pending[i : i + embed_batch]
            texts = [row[4] for row in batch]
            try:
                vectors = embed_texts(embed_base_url, embed_api_key, embed_model, texts, embed_timeout)
                if len(vectors) != len(batch):
                    raise RuntimeError(f"count mismatch: expected {len(batch)}, got {len(vectors)}")
            except RuntimeError as exc:
                errors += len(batch)
                print(json.dumps({"ok": False, "error": str(exc), "done": done, "remaining": total - done}), flush=True)
                continue

            now = now_iso()
            for (rowid, entry_id, chunk_index, text_hash, text), vec in zip(batch, vectors):
                nrm = vec_norm(vec)
                emb_json = json.dumps(vec, separators=(",", ":"))
                conn.execute(
                    "UPDATE email_chunks SET embedding_json=?, embedding_norm=?, updated_at=? WHERE rowid=?",
                    (emb_json, nrm, now, rowid),
                )
                _store_chunk_vector(conn, int(rowid), vec, now)
                conn.execute(
                    """
                    INSERT INTO embedding_cache(text_hash, embedding_json, embedding_norm, updated_at)
                    VALUES(?, ?, ?, ?)
                    ON CONFLICT(text_hash) DO UPDATE SET
                      embedding_json=excluded.embedding_json,
                      embedding_norm=excluded.embedding_norm,
                      updated_at=excluded.updated_at
                    """,
                    (text_hash, emb_json, nrm, now),
                )
            conn.commit()
            done += len(batch)
            if done % 500 == 0 or done == total:
                print(json.dumps({"ok": True, "done": done, "total": total, "errors": errors}), flush=True)

    print(json.dumps({"ok": True, "finished": True, "done": done, "errors": errors, "total": total}), flush=True)
    return 0 if errors == 0 else 1


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Local Outlook email RAG index")
    p.add_argument("--db", type=Path, default=DEFAULT_DB)
    sub = p.add_subparsers(dest="cmd", required=True)

    p_sync = sub.add_parser("sync", help="Custom sync with explicit range (default: newest-only via anchor-based widening)")
    p_sync.add_argument("count", type=_positive_int, nargs="?", default=DEFAULT_DAILY_COUNT)
    p_sync.add_argument("folder", nargs="?", default=DEFAULT_DAILY_FOLDER)
    p_sync.add_argument("hours", type=_nonnegative_int, nargs="?", default=DEFAULT_DAILY_HOURS)
    p_sync.add_argument("mode", nargs="?", default=DEFAULT_DAILY_MODE, choices=["all", "unread"])
    p_sync.add_argument("max_chars", type=_positive_int, nargs="?", default=DEFAULT_DAILY_MAX_CHARS)
    _add_paging_args(p_sync)
    _add_embedding_args(p_sync)
    _add_attachment_args(p_sync)
    _add_skip_existing_args(p_sync)
    p_sync.set_defaults(func=cmd_sync)

    p_sync_default = sub.add_parser("sync-default", help="Daily incremental sync (newest only, since last sync)")
    p_sync_default.add_argument("--count", type=_positive_int, default=DEFAULT_DAILY_COUNT)
    p_sync_default.add_argument("--folder", default=DEFAULT_DAILY_FOLDER)
    p_sync_default.add_argument("--hours", type=_nonnegative_int, default=DEFAULT_DAILY_HOURS)
    p_sync_default.add_argument("--mode", default=DEFAULT_DAILY_MODE, choices=["all", "unread"])
    p_sync_default.add_argument("--max-chars", dest="max_chars", type=_positive_int, default=DEFAULT_DAILY_MAX_CHARS)
    _add_paging_args(p_sync_default)
    _add_embedding_args(p_sync_default)
    _add_attachment_args(p_sync_default)
    _add_skip_existing_args(p_sync_default)
    p_sync_default.set_defaults(func=cmd_sync_default)

    p_sync_full = sub.add_parser(
        "sync-full",
        help="Full mailbox sync (safe upsert by default; use --rebuild to clear first)",
    )
    p_sync_full.add_argument("--count-per-folder", type=_positive_int, default=DEFAULT_FULL_COUNT_PER_FOLDER)
    p_sync_full.add_argument("--folders", default=",".join(DEFAULT_FULL_FOLDERS))
    p_sync_full.add_argument("--mode", default=DEFAULT_FULL_MODE, choices=["all", "unread"])
    p_sync_full.add_argument("--max-chars", dest="max_chars", type=_positive_int, default=DEFAULT_FULL_MAX_CHARS)
    p_sync_full.add_argument(
        "--keep-embedding-cache",
        action="store_true",
        help="Keep embedding cache during full rebuild (faster, less strict)",
    )
    rebuild_group = p_sync_full.add_mutually_exclusive_group()
    rebuild_group.add_argument(
        "--rebuild",
        action="store_true",
        help="Explicitly clear the email index after the first successful Outlook page, then rebuild.",
    )
    rebuild_group.add_argument(
        "--no-rebuild",
        dest="no_rebuild",
        action="store_true",
        help="Deprecated compatibility flag; safe upsert is already the default.",
    )
    _add_paging_args(p_sync_full)
    _add_embedding_args(p_sync_full)
    _add_attachment_args(p_sync_full)
    _add_skip_existing_args(p_sync_full)
    p_sync_full.set_defaults(func=cmd_sync_full)

    p_sync_full_safe = sub.add_parser(
        "sync-full-safe",
        help="Resumable full sync in small steps (recommended for large mailboxes)",
    )
    p_sync_full_safe.add_argument("--count-per-step", type=_positive_int, default=DEFAULT_FULL_SAFE_COUNT_PER_STEP)
    p_sync_full_safe.add_argument("--folders", default=",".join(DEFAULT_FULL_SAFE_FOLDERS))
    p_sync_full_safe.add_argument(
        "--hours-windows",
        default=",".join(str(x) for x in DEFAULT_FULL_SAFE_HOURS_WINDOWS),
        help="Comma-separated cumulative windows; 0 means all time (default: 0)",
    )
    p_sync_full_safe.add_argument("--mode", default=DEFAULT_FULL_MODE, choices=["all", "unread"])
    p_sync_full_safe.add_argument("--max-chars", dest="max_chars", type=_positive_int, default=DEFAULT_FULL_MAX_CHARS)
    p_sync_full_safe.add_argument(
        "--from-start",
        action="store_true",
        help="Ignore checkpoint and restart safe full sync from the first step.",
    )
    p_sync_full_safe.add_argument(
        "--keep-embedding-cache",
        action="store_true",
        help="Deprecated compatibility flag; sync-full-safe never clears the index.",
    )
    p_sync_full_safe.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview which plans will run vs be skipped (after dedup). Does NOT execute or touch checkpoint.",
    )
    p_sync_full_safe.add_argument(
        "--foreground",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    _add_paging_args(p_sync_full_safe)
    _add_embedding_args(p_sync_full_safe)
    _add_attachment_args(p_sync_full_safe)
    _add_skip_existing_args(p_sync_full_safe)
    p_sync_full_safe.set_defaults(
        func=cmd_sync_full_safe,
        page_size=DEFAULT_FULL_SAFE_PAGE_SIZE,
    )

    p_cleanup = sub.add_parser("cleanup", help="Delete exported attachment files/directories")
    p_cleanup.add_argument(
        "--dirs",
        nargs="*",
        type=Path,
        default=None,
        help="Attachment export directories to clean. Default: attachments and attachments_test.",
    )
    p_cleanup.set_defaults(func=cmd_cleanup)

    p_cleanup_daemon = sub.add_parser(
        "cleanup-daemon",
        help="Clean up stale Outlook daemon state (schtasks, BATs). Safe to run anytime.",
    )
    p_cleanup_daemon.add_argument(
        "--kill-daemon",
        action="store_true",
        help="Also kill the live outlook_daemon.ps1 powershell process and rotate its log. "
             "Use this when the daemon is hung, the pipe is wedged, or you want a clean slate. "
             "Next sync will relaunch the daemon.",
    )
    p_cleanup_daemon.set_defaults(func=cmd_cleanup_daemon)

    p_kill_embed = sub.add_parser(
        "kill-embed-server",
        help="Stop the persistent local embed server. Next search will respawn it.",
    )
    p_kill_embed.set_defaults(func=cmd_kill_embed_server)

    p_est = sub.add_parser("estimate", help="Estimate Outlook-side mailbox size for RAG planning")
    p_est.add_argument(
        "--folders",
        default=DEFAULT_ESTIMATE_FOLDERS,
        help="Comma-separated folders to estimate (e.g. InboxAll,Sent)",
    )
    p_est.set_defaults(func=cmd_estimate)

    p_progress = sub.add_parser("progress", help="Show latest sync progress snapshot")
    p_progress.set_defaults(func=cmd_progress)

    p_search = sub.add_parser("search", help="Search indexed emails")
    p_search.add_argument("query")
    p_search.add_argument("top_k", type=_positive_int, nargs="?", default=8)
    p_search.add_argument("--mode", choices=["hybrid", "bm25", "vector"], default="hybrid")
    p_search.add_argument("--text-weight", type=float, default=0.35)
    p_search.add_argument("--vector-weight", type=float, default=0.65)
    p_search.add_argument("--candidate-multiplier", type=_positive_int, default=5)
    p_search.add_argument("--embed-base-url", default=DEFAULT_EMBED_BASE_URL)
    p_search.add_argument("--embed-model", default=DEFAULT_EMBED_MODEL)
    p_search.add_argument("--embed-api-key", default=DEFAULT_EMBED_API_KEY)
    p_search.add_argument("--embed-timeout", type=_positive_int, default=DEFAULT_EMBED_TIMEOUT)
    p_search.add_argument(
        "--embed-backend",
        choices=["local", "api"],
        default=DEFAULT_EMBED_BACKEND,
        help="Embedding backend for query vector: 'local' (default) or 'api'",
    )
    p_search.set_defaults(func=cmd_search)

    p_status = sub.add_parser("status", help="Show index status")
    p_status.set_defaults(func=cmd_status)

    p_embed_pending = sub.add_parser(
        "embed-pending",
        help="Embed all chunks that are missing embeddings (no Outlook re-fetch needed)",
    )
    _add_embedding_args(p_embed_pending)
    p_embed_pending.set_defaults(func=cmd_embed_pending)

    return p


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return int(args.func(args))
    except RuntimeError as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())
