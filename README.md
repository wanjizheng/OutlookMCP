# OutlookMCP

Windows-native, **read-only** Outlook and local email RAG for ChatGPT desktop and other STDIO MCP clients. No WSL, OpenClaw, Microsoft Graph OAuth, school tenant app registration, API key, public port or cloud-hosted mailbox database is required at runtime.

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

No send, draft, delete, mark-read, meeting-change, shell, arbitrary file or synchronization tool is exposed. Sync and model installation are separate local administrative commands.

## Storage

Code and dependencies live in this checkout. Personal index files are under `data/`, model under `models/`, and private diagnostics/config backups under `artifacts/`; all are excluded from Git. `OUTLOOK_MCP_DATA_DIR` can point to another local SSD directory. The installer restricts default data directories to the current user and SYSTEM. A different configured data directory must be protected separately.

This project reuses the fixed Outlook scripts, existing email RAG SQLite/chunk schema, and local GGUF embedding implementation. Linux/OpenClaw compatibility sources are under `compat/`; the native server never imports their Linux runtime.

Native `sync` also extracts new PDF, DOCX, PPTX, XLSX, CSV, TXT and Markdown attachments into the local index by default. Raw attachments are temporary; extraction limits, explicit skipped/error states and history backfill instructions are in [INSTALL.md](INSTALL.md#新增附件文本提取windows-原生). This administrative feature adds no MCP write/export tools and no WSL dependency.
