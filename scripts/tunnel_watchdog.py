"""Windows interactive-user supervisor. Restarts transport, never mail operations."""

import argparse
from http.client import HTTPConnection, HTTPException
import json
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import subprocess
import time
from urllib.parse import urlsplit

ROOT = Path(__file__).resolve().parents[1]
PRIVATE = ROOT / "artifacts/tunnel-client"
INTERVAL = 10
UNREADY_SECONDS = 180


def ready(path=PRIVATE / "health.url"):
    """Never follow redirects/proxies or contact a non-loopback health endpoint."""
    try:
        url = path.read_text(encoding="utf-8").strip()
        parsed = urlsplit(url)
        if (parsed.scheme != "http" or parsed.hostname != "127.0.0.1"
                or not parsed.port or parsed.username or parsed.password
                or parsed.path not in ("", "/") or parsed.query or parsed.fragment):
            return False
        # http.client does not follow redirects or consult proxy environment vars.
        connection = HTTPConnection("127.0.0.1", parsed.port, timeout=3)
        try:
            connection.request("GET", "/readyz")
            response = connection.getresponse()
            return response.status == 200 and response.read(4096).strip() == b"ready"
        finally:
            connection.close()
    except (OSError, ValueError, HTTPException):
        return False


class Recovery:
    def __init__(self, now):
        self.last_ready = now
        self.attempts = 0

    def observe(self, now, healthy):
        if healthy:
            self.last_ready = now
            self.attempts = 0
        return now - self.last_ready >= UNREADY_SECONDS

    def delay(self):
        delay = min(5 * 2 ** min(self.attempts, 6), 300)
        self.attempts += 1
        return delay


def child_command(enable_send):
    args = [str(ROOT / ".venv/Scripts/python.exe"), str(ROOT / "scripts/setup_tunnel.py"), "run"]
    return args + (["--enable-send"] if enable_send else [])


def terminate_tree(child):
    if child.poll() is None:
        subprocess.run([str(Path(os.environ["SystemRoot"]) / "System32/taskkill.exe"),
                        "/PID", str(child.pid), "/T", "/F"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       creationflags=subprocess.CREATE_NO_WINDOW, timeout=30, check=False)
        child.wait(timeout=30)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--enable-send", action="store_true")
    args = parser.parse_args()
    if os.name != "nt":
        raise SystemExit("Native Windows is required")
    PRIVATE.mkdir(parents=True, exist_ok=True)
    # Lock is released by Windows even if the supervisor is forcibly killed.
    import msvcrt
    lock = (PRIVATE / "watchdog.lock").open("a+b")
    lock.seek(0)
    try:
        msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        lock.close()
        return  # Another supervisor already owns this checkout.

    logger = logging.getLogger("watchdog")
    logger.setLevel(logging.INFO)
    handler = RotatingFileHandler(PRIVATE / "watchdog.log", maxBytes=262144, backupCount=2, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(message)s"))
    logger.addHandler(handler)
    # Every descendant belongs to this job. Killing/stopping the task closes the
    # only handle and kills the entire tunnel/backend tree, preventing orphans.
    import win32api
    import win32job
    job = win32job.CreateJobObject(None, "")
    info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
    info["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
    win32job.AssignProcessToJobObject(job, win32api.GetCurrentProcess())
    recovery = Recovery(time.monotonic())
    child = None
    restarts = 0
    try:
        while True:
            (PRIVATE / "health.url").unlink(missing_ok=True)
            child = subprocess.Popen(child_command(args.enable_send), cwd=ROOT,
                                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                     stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
            logger.info("transport_started pid=%d restart=%d", child.pid, restarts)
            recovery.last_ready = time.monotonic()
            previous = None
            while child.poll() is None:
                healthy = ready()
                if healthy != previous:
                    logger.info("ready=%s", healthy)
                    previous = healthy
                state = {"supervisor_pid": os.getpid(), "child_pid": child.pid,
                         "ready": healthy, "restarts": restarts, "enable_send": args.enable_send,
                         "checked_at": time.time()}
                temp = PRIVATE / "watchdog-status.tmp"
                temp.write_text(json.dumps(state), encoding="utf-8")
                temp.replace(PRIVATE / "watchdog-status.json")
                if recovery.observe(time.monotonic(), healthy):
                    logger.info("unready_timeout restarting_transport")
                    break
                time.sleep(INTERVAL)
            terminate_tree(child)
            restarts += 1
            delay = recovery.delay()
            logger.info("retry_in_seconds=%d", delay)
            time.sleep(delay)
    except BaseException as error:
        # Closing our own job handle would terminate us with exit code zero and
        # suppress Task Scheduler's failure recovery. Fail the whole job instead.
        logger.error("supervisor_failed error_type=%s", type(error).__name__)
        win32job.TerminateJobObject(job, 1)
        raise


if __name__ == "__main__":
    main()
