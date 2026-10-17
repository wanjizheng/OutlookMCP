# OutlookMCP development and installation

All development, installers, environments, generated data and validation artifacts belong in this checkout. Never put mailbox contents, database files, models, client config backups, logs or credentials in Git.

For installation, follow INSTALL.md. Use native Windows Python and PowerShell; WSL is neither an installation prerequisite nor a runtime dependency. Only the explicitly requested one-time legacy import and OpenClaw compatibility deployment may touch WSL.

The MCP is read-only. Never expose sending, drafts, marking read, deleting, arbitrary commands, arbitrary filesystem paths, attachment exports, synchronization or model installation as MCP tools. Administrative sync is a separate local CLI.

Preserve other client settings and OpenClaw CLI signatures. Run the regression suite and real MCP stdio smoke test before claiming acceptance. Client config registration is not evidence that a running ChatGPT session has reloaded tools. Report the boundary honestly.

Do not stop the user's ChatGPT app or other chats to force a reload. The user can reload MCP servers or start a new local chat if live reload is unavailable.

Changes to compat sources must be followed by `scripts/build_bridge.py`; generated Windows libraries are committed and checked for reproducibility. The Windows backend must not import the Linux-only legacy email_rag.py.
