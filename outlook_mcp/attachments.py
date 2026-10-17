"""Bounded native attachment extraction for administrative synchronization only."""

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from defusedxml import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
MAX_FILE = 10 * 1024 * 1024
MAX_TEXT = 12000
SUPPORTED = {".pdf", ".docx", ".pptx", ".xlsx", ".csv", ".txt", ".md"}
FATAL = {"export_error", "parse_error", "timeout"}


def extract_file(path):
    """Runs in a disposable worker. No Office application, shell or macros."""
    path = Path(path)
    if path.stat().st_size > MAX_FILE:
        return {"status": "size_limit", "text": ""}
    ext = path.suffix.lower()
    if ext not in SUPPORTED:
        return {"status": "unsupported", "text": ""}
    pieces = []
    truncated = False

    def add(value):
        remaining = MAX_TEXT - sum(len(x) for x in pieces)
        if remaining > 0:
            pieces.append(str(value)[:remaining])

    if ext == ".pdf":
        from pypdf import PdfReader

        reader = PdfReader(path)
        if reader.is_encrypted:
            return {"status": "encrypted", "text": ""}
        truncated = len(reader.pages) > 200
        for page in reader.pages[:200]:
            add((page.extract_text() or "") + "\n")
            if sum(map(len, pieces)) >= MAX_TEXT:
                break
    elif ext in {".txt", ".md", ".csv"}:
        raw = path.read_bytes()
        encoding = (
            "utf-16" if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig"
        )
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            text = raw.decode("cp1252")
        add(text)
    else:
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            if len(infos) > 1000 or sum(i.file_size for i in infos) > 20 * 1024 * 1024:
                return {"status": "expanded_size_limit", "text": ""}

            def xml(name):
                return ET.fromstring(
                    z.read(name),
                    forbid_dtd=True,
                    forbid_entities=True,
                    forbid_external=True,
                )

            def tag(node):
                return node.tag.rsplit("}", 1)[-1]

            names = z.namelist()
            if ext == ".xlsx":
                shared = []
                if "xl/sharedStrings.xml" in names:
                    shared = [
                        "".join(n.text or "" for n in si.iter() if tag(n) == "t")
                        for si in xml("xl/sharedStrings.xml")
                    ]
                for name in sorted(
                    n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)
                ):
                    for cell in xml(name).iter():
                        if tag(cell) != "c":
                            continue
                        values = [
                            n.text or "" for n in cell.iter() if tag(n) in ("v", "t")
                        ]
                        if cell.get("t") == "s" and values:
                            values = [shared[int(values[0])]]
                        add(" ".join(values) + "\n")
            else:
                pattern = (
                    r"word/(document|header\d+|footer\d+)\.xml"
                    if ext == ".docx"
                    else r"ppt/slides/slide\d+\.xml"
                )
                for name in sorted(n for n in names if re.fullmatch(pattern, n)):
                    for n in xml(name).iter():
                        if tag(n) == "t":
                            add((n.text or "") + "\n")
    text = "".join(pieces).strip()
    return {
        "status": "ok" if text else "no_text",
        "text": text,
        "truncated": truncated or sum(map(len, pieces)) >= MAX_TEXT,
    }


def extract_worker(path):
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "outlook_mcp.attachments", "--extract", str(path)],
            cwd=ROOT,
            capture_output=True,
            timeout=20,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        result = json.loads(proc.stdout.decode("utf-8"))
        if not isinstance(result, dict) or result.get("status") not in {
            "ok",
            "no_text",
            "encrypted",
            "size_limit",
            "expanded_size_limit",
            "unsupported",
            "parse_error",
        }:
            raise ValueError("Invalid extraction output")
        result["text"] = str(result.get("text", ""))[:MAX_TEXT]
        return result
    except subprocess.TimeoutExpired:
        return {"status": "timeout", "text": ""}
    except (OSError, ValueError):
        return {"status": "parse_error", "text": ""}


def collect(row, data_dir):
    """Return text/status records; temporary raw attachments are always removed."""
    if not row.get("Attachments"):
        return []
    if os.name != "nt":
        raise RuntimeError("Attachment sync requires native Windows")
    root = Path(data_dir) / "attachment-tmp"
    root.mkdir(parents=True, exist_ok=True)
    # Current user + SYSTEM only, including explicitly configured data directories.
    import win32api
    import win32con
    import win32security

    token = win32security.OpenProcessToken(
        win32api.GetCurrentProcess(), win32con.TOKEN_QUERY
    )
    try:
        sid = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
    finally:
        token.Close()
    acl = win32security.ACL()
    for principal in (sid, win32security.ConvertStringSidToSid("S-1-5-18")):
        acl.AddAccessAllowedAceEx(
            win32security.ACL_REVISION, 3, win32con.GENERIC_ALL, principal
        )
    win32security.SetNamedSecurityInfo(
        str(root),
        win32security.SE_FILE_OBJECT,
        win32security.DACL_SECURITY_INFORMATION
        | win32security.PROTECTED_DACL_SECURITY_INFORMATION,
        None,
        None,
        acl,
        None,
    )
    with tempfile.TemporaryDirectory(dir=root) as tmp:
        exe = (
            Path(os.environ["SystemRoot"])
            / "System32/WindowsPowerShell/v1.0/powershell.exe"
        )
        request = {
            "entry_id": row["EntryID"],
            "store_id": row.get("StoreID", ""),
            "directory": tmp,
        }
        proc = subprocess.run(
            [
                str(exe),
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(ROOT / "scripts/attachment_export.ps1"),
            ],
            input=json.dumps(request).encode(),
            capture_output=True,
            timeout=120,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        records = json.loads(proc.stdout.decode("utf-8-sig"))
        if proc.returncode or not isinstance(records, list):
            raise RuntimeError("Attachment export failed; page not committed")
        for record in records:
            if record["status"] == "exported":
                path = Path(record.pop("path")).resolve()
                if path.parent != Path(tmp).resolve() or path.is_symlink():
                    raise RuntimeError("Unexpected attachment export path")
                record["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
                record.update(extract_worker(path))
            record.pop("path", None)
        return records


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--extract", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = extract_file(args.extract)
    except Exception:
        result = {"status": "parse_error", "text": ""}
    print(json.dumps(result, ensure_ascii=True))
