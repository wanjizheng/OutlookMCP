"""Explicit, fixed-scope Windows index refresh. No mailbox mutation or shell input."""
from contextlib import contextmanager
import datetime as dt
import json
import os
from pathlib import Path
import re
import subprocess
import threading
import time
import uuid

from . import cli, rag

HOURS = 72
PAGES = 2
VECTORS = 64
TIMEOUT = 600
ACTIVE = {"queued", "running"}
_state_mutex = threading.RLock()


def public(state):
    return {key: value for key, value in state.items()
            if key not in {"worker_pid", "process_created", "created_epoch"}}


def directory():
    path = rag.db_path().parent / "refresh-jobs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def job_path(job_id):
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        raise ValueError("Invalid refresh job ID")
    return rag.db_path().parent / "refresh-jobs" / (job_id + ".json")


def save(state):
    with _state_mutex:
        directory()
        path = job_path(state["job_id"])
        temporary = path.with_suffix(".tmp")
        temporary.write_text(json.dumps(state), encoding="utf-8")
        temporary.replace(path)


@contextmanager
def admission_lock():
    import msvcrt
    with (directory() / "admission.lock").open("a+b") as lock:
        lock.seek(0)
        if os.fstat(lock.fileno()).st_size == 0:
            lock.write(b"0")
            lock.flush()
        lock.seek(0)
        try:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as error:
            raise RuntimeError("Refresh admission busy; retry status shortly") from error
        try:
            yield
        finally:
            lock.seek(0)
            msvcrt.locking(lock.fileno(), msvcrt.LK_UNLCK, 1)


def process_started(pid):
    import win32api
    import win32process
    try:
        handle = win32api.OpenProcess(0x1000, False, pid)
        try:
            if win32process.GetExitCodeProcess(handle) != 259:
                return None
            return win32process.GetProcessTimes(handle)["CreationTime"].timestamp()
        finally:
            handle.Close()
    except Exception:
        return None


def status(job_id):
    path = job_path(job_id)
    if not path.exists():
        raise ValueError("Unknown refresh job ID")
    state = json.loads(path.read_text(encoding="utf-8"))
    # Report interruptions without writing over a worker's concurrently updated state.
    if state["status"] in ACTIVE:
        if state.get("worker_pid"):
            alive = process_started(state["worker_pid"])
            interrupted = alive is None or abs(alive - state["process_created"]) > 1
        else:
            interrupted = time.time() - state["created_epoch"] > 30
        if interrupted:
            state.update(status="interrupted", note="Worker stopped. Committed pages may remain; nothing is automatically replayed.")
    return state


def start():
    if os.name != "nt":
        raise RuntimeError("Native Windows is required")
    with admission_lock():
        latest = directory() / "latest.json"
        if latest.exists():
            current = status(json.loads(latest.read_text())["job_id"])
            if current["status"] in ACTIVE:
                return dict(public(current), already_running=True)
        job_id = uuid.uuid4().hex
        state = {"job_id": job_id, "status": "queued", "created_epoch": time.time(),
                 "started_at": dt.datetime.now(dt.UTC).isoformat(), "poll_after_seconds": 10,
                 "scope": {"folders": ["InboxAll", "Sent"], "hours": HOURS,
                           "max_pages_per_folder": PAGES, "max_messages_per_folder": PAGES * 100,
                           "max_vectors": VECTORS, "timeout_seconds": TIMEOUT},
                 "mailbox_modified": False, "wsl_required": False}
        save(state)
        latest.write_text(json.dumps({"job_id": job_id}), encoding="utf-8")
        # Fixed executable/module/arguments, no command strings supplied by the model.
        env = dict(os.environ)
        env.pop("CONTROL_PLANE_API_KEY", None)
        env["EMAIL_RAG_AUTO_INSTALL"] = "0"
        try:
            subprocess.Popen([str(rag.ROOT / ".venv/Scripts/pythonw.exe"),
                              "-m", "outlook_mcp.refresh", job_id], cwd=rag.ROOT, env=env,
                             stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
        except OSError:
            state.update(status="failed", error="worker_start_failed")
            save(state)
        # Keep bounded metadata only; no email bodies, subjects or credentials.
        old = sorted(directory().glob("[a-f0-9]" * 32 + ".json"), key=lambda p: p.stat().st_mtime)
        for path in old[:-30]:
            path.unlink(missing_ok=True)
        return public(state)


def counts():
    with rag.connect() as c:
        return {"emails": c.execute("SELECT count(*) FROM emails").fetchone()[0],
                "chunks": c.execute("SELECT count(*) FROM email_chunks").fetchone()[0],
                "vectors": c.execute("SELECT count(*) FROM email_chunk_vectors").fetchone()[0],
                "pending_recent_vectors": c.execute(
                    "SELECT count(*) FROM email_chunks c LEFT JOIN email_chunk_vectors v ON v.chunk_id=c.chunk_id WHERE v.chunk_id IS NULL AND julianday(c.received_time)>=julianday('now', ?)",
                    (f"-{HOURS} hours",)).fetchone()[0]}


def execute(state):
    """Whole job shares the CLI writer lock, including embedding."""
    try:
        with cli.writer_lock():
            rag.initialize()
            state.update(stage="sync", before=counts())
            save(state)
            state["folders"] = cli._sync(["InboxAll", "Sent"], HOURS, PAGES)
            state.update(stage="embedding", after=counts())
            save(state)
            if (rag.ROOT / "models/embedding.gguf").exists():
                try:
                    state["embedding"] = rag.embed_pending(VECTORS, recent_hours=HOURS)
                except Exception as error:
                    state["embedding"] = {"error": "local_embedding_failed", "error_type": type(error).__name__}
            else:
                state["embedding"] = {"error": "local_model_not_installed"}
            state["after"] = counts()
            state["has_more_mail"] = any(x["has_more"] for x in state["folders"].values())
            state["status"] = "partial" if (state["has_more_mail"] or state["after"]["pending_recent_vectors"] or "error" in state["embedding"]) else "completed"
            state["freshness"] = rag.status()["last_sync"]
            state["note"] = "Limited 72-hour refresh only; not a full historical sync. Attachment skips are reported by category."
    except Exception as error:
        state.update(status="failed", error="refresh_failed_or_writer_busy", error_type=type(error).__name__,
                     note="Earlier pages may be committed; no full-success claim. Retry only on user request.")
        try:
            state["after"] = counts()
        except Exception:
            pass
    state.update(stage="finished", finished_at=dt.datetime.now(dt.UTC).isoformat())
    save(state)


def worker(job_id):
    import win32api
    import win32job
    os.environ["EMAIL_RAG_AUTO_INSTALL"] = "0"
    state = json.loads(job_path(job_id).read_text(encoding="utf-8"))
    created = process_started(os.getpid())
    if created is None:
        state.update(status="failed", error="worker_identity_unavailable")
        save(state)
        return
    state.update(status="running", worker_pid=os.getpid(), process_created=created)
    save(state)
    job = win32job.CreateJobObject(None, "")
    info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
    info["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
    win32job.AssignProcessToJobObject(job, win32api.GetCurrentProcess())

    def expired():
        with _state_mutex:
            state.update(status="timed_out", note="Ten-minute budget expired. Committed pages may remain; no automatic retry.",
                         finished_at=dt.datetime.now(dt.UTC).isoformat())
            save(state)
            win32job.TerminateJobObject(job, 1)

    timer = threading.Timer(TIMEOUT, expired)
    timer.daemon = True
    timer.start()
    try:
        execute(state)
    finally:
        timer.cancel()
    # Closing this job also removes descendant extraction/model processes.


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        raise SystemExit("Expected a refresh job ID")
    worker(sys.argv[1])
