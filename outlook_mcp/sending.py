"""Opt-in immutable, single-attempt messages; no mailbox access during prepare."""

import json
import os
import re
import secrets
import subprocess
import threading
import time
from pathlib import Path

_pending = {}
_lock = threading.Lock()
TTL = 600
ROOT = Path(__file__).resolve().parents[1]
ADDRESS = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?\.[A-Za-z]{2,63}\Z")


def _enabled():
    if os.environ.get("OUTLOOK_MCP_ENABLE_SEND") != "1":
        raise RuntimeError("Sending disabled; enable explicitly on this server")


def _address(value):
    if not isinstance(value, str) or len(value) > 254 or not ADDRESS.fullmatch(value):
        raise ValueError("Use explicit SMTP addresses without display names or headers")
    return value


def prepare(sender_account, to, subject, body, cc=None, bcc=None):
    _enabled()
    sender_account = _address(sender_account)
    fields = {}
    for name, values in (("to", to), ("cc", cc or []), ("bcc", bcc or [])):
        if not isinstance(values, list) or len(values) > 30:
            raise ValueError("Recipients must be lists of up to 30 SMTP addresses")
        fields[name] = [_address(v) for v in values]
    if not fields["to"] or sum(map(len, fields.values())) > 30:
        raise ValueError("Require at least one To and at most 30 total recipients")
    if not isinstance(subject, str) or not subject.strip() or len(subject) > 255:
        raise ValueError("Subject required (up to 255 characters)")
    if any(ord(ch) < 32 for ch in subject):
        raise ValueError("Subject cannot contain control characters")
    if not isinstance(body, str) or not body.strip() or len(body) > 50000:
        raise ValueError("Plain-text body required (up to 50000 characters)")
    if any(ord(ch) < 32 and ch not in "\r\n\t" for ch in body):
        raise ValueError("Body contains unsupported control characters")
    payload = dict(sender_account=sender_account, **fields, subject=subject, body=body)
    now = time.monotonic()
    token = secrets.token_urlsafe(32)
    with _lock:
        for old in list(_pending):
            if now - _pending[old][0] >= TTL:
                del _pending[old]
        if len(_pending) >= 50:
            raise RuntimeError("Too many pending messages; wait for expiry")
        _pending[token] = (now, json.dumps(payload, ensure_ascii=False))
    return dict(preparation_id=token, expires_in_seconds=TTL, message=payload,
                attachments=[], sent=False,
                next_step="Show every field to the user and STOP. Only after the user's explicit approval in ChatGPT call send_prepared_email. No separate desktop confirmation window is used.")


def send(preparation_id, *, confirmation_test=False):
    _enabled()
    with _lock:
        entry = _pending.pop(preparation_id, None)
    if entry is None or time.monotonic() - entry[0] >= TTL:
        raise ValueError("Preparation expired, unknown, or already attempted; prepare again")
    payload = json.loads(entry[1])
    # The ChatGPT host owns user confirmation; a caller-supplied boolean is not
    # proof of consent. The backend binds the immutable preview and consumes it
    # once, and the tool remains annotated as a non-idempotent write.
    if confirmation_test:
        return {"status": "chat_confirmation_test_passed", "sent": False, "test": True,
                "message": payload, "note": "No Outlook call or desktop window. ChatGPT must obtain explicit approval before this tool call."}
    # Consumed before subprocess creation: errors/timeouts must never auto-retry.
    if os.name != "nt":
        raise RuntimeError("Sending requires native Windows Outlook")
    exe = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    try:
        result = subprocess.run(
            [str(exe), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", str(ROOT / "scripts/send_prepared.ps1")], cwd=ROOT,
            input=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            capture_output=True, timeout=90,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode:
            raise RuntimeError("Outlook submission failed or is uncertain; inspect Outlook before retrying")
        return json.loads(result.stdout.decode("utf-8-sig"))
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("Send outcome unknown after timeout. Check Outlook before preparing another message.") from exc
