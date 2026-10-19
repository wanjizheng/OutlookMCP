# OutlookMCP development and installation

All development, installers, environments, generated data and validation artifacts belong in this checkout. Never put mailbox contents, database files, models, client config backups, logs or credentials in Git.

For installation, follow INSTALL.md. Use native Windows Python and PowerShell; WSL is neither an installation prerequisite nor a runtime dependency. Only the explicitly requested one-time legacy import and OpenClaw compatibility deployment may touch WSL.

The default MCP is read-only. The user has explicitly authorized opt-in sending,
but only after full immutable message preview and explicit confirmation in a
subsequent ChatGPT user turn on every attempt. The user explicitly rejected a
custom desktop popup: reuse the existing PowerShell Outlook sender. Confirmation
is owned by the ChatGPT conversation/client approval mechanism; the backend does
not claim it can independently authenticate a human chat reply. Never expose a
confirmation-bypass boolean, drafts, marking
read, deleting, arbitrary commands, arbitrary filesystem paths, attachment
exports, arbitrary synchronization or model installation as MCP tools. The user
has explicitly authorized the opt-in resumable incremental and verified full sync described
in SPEC.md; arbitrary commands, scopes and paths remain prohibited. Never send real test email without approval of the
specific recipients and message.

Preserve other client settings and OpenClaw CLI signatures. Run the regression suite and real MCP stdio smoke test before claiming acceptance. Client config registration is not evidence that a running ChatGPT session has reloaded tools. Report the boundary honestly.

Do not stop the user's ChatGPT app or other chats to force a reload. The user can reload MCP servers or start a new local chat if live reload is unavailable.

Changes to compat sources must be followed by `scripts/build_bridge.py`; generated Windows libraries are committed and checked for reproducibility. The Windows backend must not import the Linux-only legacy email_rag.py.

User-authorized extension: confirmed routine calendar writes described in SPEC.md
supersede the earlier meeting-write prohibition. All other restrictions remain.
