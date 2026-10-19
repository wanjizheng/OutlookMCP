"""Native calendar operations: immutable previews, stale-state rejection, one attempt."""
import datetime as dt
import hashlib
import json
import os
import secrets
import subprocess
import threading
import time
from . import sending

_pending = {}
_lock = threading.Lock()
_references = {}
REFERENCE_TTL = 3600


def remember_events(items):
    """Pin one listed occurrence; no mailbox access and no series inference."""
    sending._enabled()
    now = time.monotonic()
    with _lock:
        for token in list(_references):
            if now - _references[token][0] >= REFERENCE_TTL:
                del _references[token]
        for event in items:
            if not event.get("EntryID") or not event.get("StoreID"):
                continue
            scope = event.get("EventScope")
            start = event.get("OccurrenceStart", "")
            if scope not in ("single", "occurrence") or (scope == "occurrence" and not start):
                continue
            if len(_references) >= 1000:
                del _references[next(iter(_references))]
            token = "event_" + secrets.token_urlsafe(18)
            target = dict(entry_id=event["EntryID"], store_id=event["StoreID"], scope=scope, occurrence_start=start)
            _references[token] = (now, target)
            event["event_ref"] = token
            event["read_parameters"] = dict(event_ref=token)
    return items


def target(entry_id, store_id, scope, occurrence_start, event_ref):
    if not event_ref:
        if scope == "single" and occurrence_start:
            scope = "occurrence"
        if scope == "series" and occurrence_start:
            raise ValueError("Series cannot include an occurrence start")
        return entry_id, store_id, scope, occurrence_start
    if entry_id or store_id or occurrence_start or scope not in ("single", ""):
        raise ValueError("Use event_ref alone; do not mix it with IDs or override its instance scope")
    with _lock:
        reference = _references.get(event_ref)
    if reference is None or time.monotonic() - reference[0] >= REFERENCE_TTL:
        raise ValueError("Event reference expired or unavailable; list calendar again")
    value = reference[1]
    return value["entry_id"], value["store_id"], value["scope"], value["occurrence_start"]



def invoke(operation, **args):
    sending._enabled()
    if os.name != "nt":
        raise RuntimeError("Native Windows Outlook required")
    exe = sending.Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    with _lock:
        try:
            result = subprocess.run([str(exe), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                "-File", str(sending.ROOT / "scripts/calendar_actions.ps1")],
                input=json.dumps(dict(operation=operation, **args), ensure_ascii=False).encode("utf-8"),
                capture_output=True, timeout=90, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError("Outcome uncertain; inspect Outlook before retrying") from exc
    data = json.loads(result.stdout.decode("utf-8-sig"))
    if result.returncode or "error" in data:
        raise RuntimeError(data.get("error", "Outlook operation failed; inspect before retry"))
    return data


def fingerprint(event):
    return hashlib.sha256(json.dumps(event, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def stamp(value):
    parsed = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        raise ValueError("Use ISO time with explicit UTC offset")
    return parsed.astimezone(dt.timezone.utc)


def read(entry_id="", store_id="", scope="single", occurrence_start="", event_ref=""):
    entry_id, store_id, scope, occurrence_start = target(entry_id, store_id, scope, occurrence_start, event_ref)
    if not entry_id or not store_id or scope not in ("single", "occurrence", "series"):
        raise ValueError("Require event/store IDs and explicit scope")
    if scope == "occurrence":
        stamp(occurrence_start)
    return invoke("read", entry_id=entry_id, store_id=store_id, scope=scope, occurrence_start=occurrence_start)


def prepare(action, sender_account, subject="", body="", start="", end="", location="",
            required=None, optional=None, entry_id="", store_id="", scope="single", occurrence_start="",
            add_attendees=None, remove_attendees=None, event_ref=""):
    sending._enabled()
    if action not in ("create", "update", "add_remove_attendees", "cancel", "accept", "tentative", "decline"):
        raise ValueError("Unsupported calendar action")
    sender_account = sending._address(sender_account)
    lists = {}
    for name, values in (("required", required), ("optional", optional), ("add_attendees", add_attendees), ("remove_attendees", remove_attendees)):
        values = [] if values is None else values
        if not isinstance(values, list) or len(values) > 100:
            raise ValueError("At most 100 explicit SMTP addresses per list")
        lists[name] = list(dict.fromkeys(sending._address(v).lower() for v in values))
    if set(lists["required"]) & set(lists["optional"]):
        raise ValueError("An attendee cannot be both required and optional")
    if sender_account.lower() in lists["remove_attendees"]:
        raise ValueError("Cannot remove the organizer")
    if set(lists["add_attendees"]) & set(lists["remove_attendees"]):
        raise ValueError("Cannot add and remove the same attendee")
    entry_id, store_id, scope, occurrence_start = target(entry_id, store_id, scope, occurrence_start, event_ref)
    if action == "create" and event_ref:
        raise ValueError("Creation cannot use an existing event reference")
    before = None
    if action == "create":
        validated = sending.prepare(sender_account, [sender_account], subject, body or "Meeting invitation")
        with sending._lock:
            sending._pending.pop(validated["preparation_id"], None)
        if stamp(end) <= stamp(start) or stamp(end) - stamp(start) > dt.timedelta(days=7):
            raise ValueError("Meeting duration must be positive and no more than seven days")
        if not lists["required"] and not lists["optional"]:
            raise ValueError("An invitation requires attendees")
        if scope != "single" or entry_id or store_id:
            raise ValueError("Creation currently supports single meetings only")
        if len(location) > 1000 or any(ord(c) < 32 for c in location):
            raise ValueError("Invalid location")
    else:
        if action != "update" and any((subject, body, start, end, location, lists["required"], lists["optional"])):
            raise ValueError("Existing events preserve all fields except explicitly requested attendees/status")
        if action == "update":
            if lists["required"] or lists["optional"]:
                raise ValueError("Use add/remove attendees to preserve existing recipient roles")
            if not any((subject, body, start, end, location)):
                raise ValueError("No field changes requested")
            if start or end:
                if not start or not end or stamp(end) <= stamp(start) or stamp(end)-stamp(start) > dt.timedelta(days=7):
                    raise ValueError("Supply valid start and end together")
                if scope == "series":
                    raise ValueError("Recurring series time edits are unsupported; select an occurrence")
            for value, limit in ((subject,255),(body,50000),(location,1000)):
                if len(value)>limit or any(ord(ch)<32 and ch not in "\r\n\t" for ch in value):
                    raise ValueError("Invalid event text")
            if any(ord(ch)<32 for ch in subject+location):
                raise ValueError("Invalid event header")
        before = read(entry_id, store_id, scope, occurrence_start)
        if action in ("update", "add_remove_attendees", "cancel") and before["organizer_smtp"].lower() != sender_account.lower():
            raise ValueError("Only the organizer may update or cancel")
        if action in ("accept", "tentative", "decline") and before["organizer_smtp"].lower() == sender_account.lower():
            raise ValueError("Organizer cannot respond to their own invitation")
        if action != "add_remove_attendees" and (lists["add_attendees"] or lists["remove_attendees"]):
            raise ValueError("Attendee changes require add_remove_attendees")
        if action == "add_remove_attendees":
            if not lists["add_attendees"] and not lists["remove_attendees"]:
                raise ValueError("No attendee changes requested")
            existing = {r["smtp"].lower() for r in before["attendees"]}
            if not set(lists["remove_attendees"]) <= existing:
                raise ValueError("Cannot remove an absent attendee")
            if set(lists["add_attendees"]) & existing:
                raise ValueError("Attendee already present")
    payload = dict(action=action, sender_account=sender_account, subject=subject, body=body,
        start=start, end=end, location=location, **lists, entry_id=entry_id, store_id=store_id,
        scope=scope, occurrence_start=occurrence_start, before=before)
    token = secrets.token_urlsafe(32)
    now = time.monotonic()
    with _lock:
        for key in list(_pending):
            if now - _pending[key][0] >= sending.TTL:
                del _pending[key]
        if len(_pending) >= 50:
            raise RuntimeError("Too many pending operations")
        _pending[token] = (now, json.dumps(payload, ensure_ascii=False))
    after = dict(before) if before else dict(subject=subject, body=body, start=start, end=end, location=location,
        attendees=[dict(smtp=a,type=t) for t,k in ((1,"required"),(2,"optional")) for a in lists[k]])
    if action == "update":
        after.update({k:v for k,v in dict(subject=subject,body=body,start=start,end=end,location=location).items() if v})
    if action == "add_remove_attendees":
        after["attendees"] = [r for r in before["attendees"] if r["smtp"].lower() not in lists["remove_attendees"]] + [dict(smtp=a,type=1) for a in lists["add_attendees"]]
    return dict(preparation_id=token, preview=payload, resulting_event=after, expires_in_seconds=sending.TTL, sent=False,
        notification="Meeting updates/cancellations notify all attendees; responses notify organizer.",
        next_step="Show every preview field, original event and exact changes; STOP. Obtain explicit approval in a subsequent user turn before commit_calendar_action.")


def commit(preparation_id, confirmation_test=False):
    sending._enabled()
    with _lock:
        entry = _pending.pop(preparation_id, None)
    if entry is None or time.monotonic() - entry[0] >= sending.TTL:
        raise ValueError("Unknown, expired or already attempted preparation")
    payload = json.loads(entry[1])
    if confirmation_test:
        return dict(status="chat_confirmation_test_passed", sent=False, test=True, preview=payload)
    if payload["before"] is not None:
        current = read(payload["entry_id"], payload["store_id"], payload["scope"], payload["occurrence_start"])
        if fingerprint(current) != fingerprint(payload["before"]):
            raise ValueError("Event changed after preview; prepare again")
    return invoke("commit", payload=payload)
