# OutlookMCP

Windows-native Outlook and local email RAG for ChatGPT desktop and other STDIO MCP clients. **Read-only by default**, with separately enabled, human-confirmed plain-text email sending. No WSL, OpenClaw, Microsoft Graph OAuth, school tenant app registration, API key, public port or cloud-hosted mailbox database is required for the local STDIO connection. Ordinary cloud Chat integration uses an optional authenticated OpenAI Secure MCP Tunnel and needs a restricted tunnel runtime key.

Requires **classic Outlook for Windows**, installed and signed in under the same interactive Windows user. New Outlook and Outlook web do not provide this COM interface. Mail returned to ChatGPT becomes part of its model context; this does not make that content local-only or override your school's data rules.

**[Complete installation instructions (中文)](INSTALL.md)** · [Architecture and limitations](docs/architecture.md) · [Validation](VALIDATION.md)

Repository: https://github.com/wanjizheng/OutlookMCP

## Ask ChatGPT to install

Give a local Windows ChatGPT/Codex session with terminal access the repository URL and this request:

> Read this repository's AGENTS.md and INSTALL.md. Clone it into my chosen Windows directory (default E:\wanjizheng\MyProject\OutlookMCP), install the native dependencies there, run Install.ps1 -WithEmbeddings -RegisterClient, build a local index or import my explicitly supplied legacy index, and run the regression and MCP smoke tests. Preserve unrelated client settings. Do not use WSL for runtime, upload email data, expose write tools, stop other chats, or claim the running client has loaded tools without checking. Report any client reload requirement.

A web-only chat cannot execute a local installer just by reading a GitHub link. Use a local session with computer/terminal access; workspace policies may still control custom MCP availability. See [official OpenAI MCP configuration](https://learn.chatgpt.com/docs/extend/mcp).

## Tools

| Tool | Purpose |
|---|---|
| `outlook_status` | Outlook availability, local index counts and freshness |
| `search_email` | Local FTS/BM25 or hybrid semantic retrieval with message IDs |
| `list_email` | Live globally ordered Outlook listing with stable pagination |
| `read_email` | Read indexed or live message body, with explicit truncation |
| `list_calendar` | Read overlapping calendar occurrences without adding appointments |

By default no write tools are exposed. Setting `OUTLOOK_MCP_ENABLE_SEND=1` adds `prepare_email` and `send_prepared_email`. Preparation never changes the mailbox. ChatGPT must show the exact immutable sender, To/CC/BCC, subject and complete body, then **stop and wait for explicit user approval in a subsequent chat turn**. After approval the existing Windows PowerShell skill script submits through Outlook COM; no custom desktop window is shown. ChatGPT/client permissions own the confirmation step; the backend validates preview identity, expiry, exact recipients/account and single-attempt submission. `confirmation_test=true` validates this chat workflow without Outlook, sending or any window. Sending attachments is not supported. No draft, delete, mark-read, meeting-change, shell or arbitrary file tools are exposed. Manual index synchronization is separately opt-in; model installation remains local-only.

See [ordinary Chat / private tunnel setup](INSTALL.md#7-普通-chatgpt-对话私有隧道和可选发送) for the distinct plugin connection. A successful local MCP call does not prove that ordinary Chat has the plugin installed.

For unattended daily use, `scripts/autostart.ps1 -Action Install` registers a
password-free, least-privilege Windows logon task. Add `-EnableSend` only when
explicitly requested; the ChatGPT confirmation requirement remains in force.
The hidden native supervisor checks readiness, restarts failed/unready transports
with bounded backoff, and cleans up the owned process tree. It never retries
email submissions. See INSTALL.md for status, restart and uninstall commands.

Manual RAG refresh is a separate opt-in: `OUTLOOK_MCP_ENABLE_REFRESH=1`, or
install the startup task with `-EnableSend -EnableRefresh`. It adds
`refresh_email_index()`, `full_sync_email_index()` and `get_refresh_status(job_id)` for
resumable incremental updates and verified full synchronization. Full cleanup is
blocked when cached Exchange cannot prove server absence. Neither changes
Outlook or sends mail. See [setup and coverage limits](docs/manual-refresh.md).

## Storage

Code and dependencies live in this checkout. Personal index files are under `data/`, model under `models/`, and private diagnostics/config backups under `artifacts/`; all are excluded from Git. `OUTLOOK_MCP_DATA_DIR` can point to another local SSD directory. The installer restricts default data directories to the current user and SYSTEM. A different configured data directory must be protected separately.

This project reuses the fixed Outlook scripts, existing email RAG SQLite/chunk schema, and local GGUF embedding implementation. Linux/OpenClaw compatibility sources are under `compat/`; the native server never imports their Linux runtime.

Native `sync` also extracts new PDF, DOCX, PPTX, XLSX, CSV, TXT and Markdown attachments into the local index by default. Raw attachments are temporary; extraction limits, explicit skipped/error states and history backfill instructions are in [INSTALL.md](INSTALL.md#新增附件文本提取windows-原生). This administrative feature adds no MCP write/export tools and no WSL dependency.
