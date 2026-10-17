# OpenClaw compatibility sources

These are snapshots of the user's existing Outlook/RAG skill sources, retained to preserve CLI compatibility. They are not MCP tools. Windows-native runtime imports none of the Linux Python code and exposes none of the legacy send/meeting commands.

Fixed files: `outlook_reader/scripts/{outlook_daemon.ps1,mail_paging.ps1,outlook-mail-read,outlook-calendar,read_outlook_calendar.ps1}` and `email-rag/{email_rag.py,test_email_rag.py}`. `scripts/build_bridge.py` derives the native PowerShell read libraries from them. Run it after source changes.

For an explicitly authorized existing deployment, run `scripts/deploy_openclaw.py --root <existing .openclaw root>`. It copies only the allowlisted source files, regression test, and the two existing `~/bin` wrappers. It does not change OpenClaw configuration or restart the gateway. Wait for any active email-rag sync to finish before replacing the daemon, then gracefully shut down the old daemon; the next normal read starts the new version.

Existing full-safe checkpoints use a new paging version marker; the next full-safe pass replays from step zero while retaining existing rows. This is necessary because old time-only checkpoints cannot prove coverage of subfolders or tied timestamps.
