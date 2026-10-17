"""Explicit, allowlisted compatibility deployment. Never changes OpenClaw config or unrelated files."""

import argparse
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FILES = (
    "outlook_daemon.ps1",
    "mail_paging.ps1",
    "outlook-mail-read",
    "outlook-calendar",
    "read_outlook_calendar.ps1",
    "test_calendar_read.ps1",
)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--root", type=Path, required=True)
    a = p.parse_args()
    scripts = a.root / "workspace/skills/outlook_reader/scripts"
    if not (scripts / "outlook_daemon.ps1").exists():
        p.error("Not an existing deployed OpenClaw root")
    for name in FILES:
        shutil.copy2(ROOT / "compat/outlook_reader/scripts" / name, scripts / name)
    shutil.copy2(ROOT / "compat/outlook_reader/SKILL.md", scripts.parent / "SKILL.md")
    for name in ("email_rag.py", "test_email_rag.py"):
        shutil.copy2(
            ROOT / "compat/email-rag" / name,
            a.root / "workspace/skills/email-rag/scripts" / name,
        )
    paging = (
        (ROOT / "tests/paging.ps1")
        .read_text()
        .replace("../compat/outlook_reader/scripts/mail_paging.ps1", "mail_paging.ps1")
    )
    (scripts / "test_mail_paging.ps1").write_text(paging, encoding="utf-8-sig", newline="\n")
    # Preserve deployed wrapper signatures and the existing common helper.
    for name in ("outlook-mail-read", "outlook-calendar"):
        shutil.copy2(scripts / name, a.root.parent / "bin" / name)
    print(
        "Deployed allowlisted compatibility sources/tests and 2 existing wrappers; config unchanged"
    )


if __name__ == "__main__":
    main()
