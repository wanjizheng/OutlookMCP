# Architecture and validation boundaries

ChatGPT → native Python MCP STDIO process → fixed-operation PowerShell bridge → classic Outlook COM, or read-only SQLite RAG connections. No process invokes wsl.exe, Bash, systemd or OpenClaw at runtime. The optional administrative importer can run once on a Linux source to make a consistent snapshot.

The local store reuses the legacy `emails`, `email_chunks`, `email_chunks_fts`, and `email_chunk_vectors` schema. Chunks use 900 characters / 150 overlap. Hybrid search combines FTS5 rankings with bounded NumPy float32 cosine batches and reciprocal-rank fusion. It uses the same GGUF embedding model as the imported index, on native Windows Node. Failures are reported as explicit BM25 fallback. Query results contain message IDs, dates, excerpts and sync timestamps; attachments already indexed by OpenClaw remain searchable.

Default MCP query connections use SQLite `mode=ro` and `PRAGMA query_only=ON`.
The administrative CLI owns full sync/backfill. The user's separately opted-in
bounded refresh adds a fixed local-index write job; no arbitrary script/file,
full backfill or mailbox mutation is exposed by that operation.

`OUTLOOK_MCP_ENABLE_REFRESH=1` adds a non-destructive write tool
`refresh_email_index()` and read-only `get_refresh_status(job_id)`. A native
worker reuses CLI sync and attachment extraction for InboxAll/Sent, 72 hours,
two pages per root, followed by at most 64 recent pending vectors. Admission
locking coalesces active requests; the existing writer lock covers both phases.
A Windows Job Object and ten-minute timer bound the worker's process tree.
Private job records let a different MCP session read progress; PID creation time
prevents mistaking a reused PID for the worker. Interrupted jobs are never
automatically replayed. Partial page commits, attachment skips, vector backlog
and limited coverage are reported; no full-history freshness claim follows.
No periodic sync is introduced. See manual-refresh.md for deployment and UX.

The default server still registers exactly five read-only tools. An opt-in server
process (`OUTLOOK_MCP_ENABLE_SEND=1`) adds immutable preparation plus a single-attempt
send operation. The pending preview is an in-memory serialized copy, expires after
ten minutes and is consumed before invoking the fixed Windows PowerShell bridge.
The bridge reuses the skill's `send_outlook_mail.ps1` and shared COM library,
matching the Outlook sender account and resolved recipient SMTP addresses.
ChatGPT must display every field, stop and obtain explicit approval in a later
user turn before the write tool. This is a host/client confirmation policy;
the backend cannot independently verify a human chat reply and does not treat a
model-supplied boolean as such proof. There is no custom desktop confirmation UI.
The confirmation-test branch does not spawn any process or create a COM object.
No attachments, HTML, stored drafts or automatic retries are supported by MCP.
Legacy skill draft/preview defaults remain intact. SMTP delivery is not
established by Outlook accepting submission.

Ordinary Chat has a separate private plugin connection via OpenAI Secure MCP
Tunnel. The local STDIO server and cloud plugin reuse the same native modules,
models and read-only RAG. The tunnel uses outbound HTTPS and account/workspace
associations; the health/admin listener remains loopback-only. A restricted runtime
key is encrypted with current-user DPAPI under ignored private artifacts, supplied
only through the tunnel process environment, and stripped at the MCP entry point.
No tunnel API key is needed for the original local connection. A local test cannot
substitute for an actual ordinary Chat invocation after plugin installation.

Administrative attachment sync uses a separate, fixed PowerShell script to save bounded by-value attachments into private temporary folders, then extracts text in timeout-bounded native Python workers. The MCP bridge operation allowlist is unchanged. Filename-derived paths are never used: exported files are numbered inside a unique temporary folder, and Python verifies the returned path stays there. Raw files are cleaned on all normal/error exits; a machine crash or forced process termination can leave a temporary folder for local inspection.

PDF uses pypdf without OCR; DOCX/PPTX/XLSX use bounded ZIP/XML parsing with defusedxml, without Office automation, macros, formulas or external entities. Each message retains bounded attachment text/status/hash records in `native_attachment_state`. Their fingerprint participates in ingest deduplication, so unchanged bodies can be backfilled and changed attachments invalidate old chunk vectors. Body-only sync preserves previously cached native attachment text. Fatal extraction errors occur before the page transaction/checkpoint; skipped formats, encrypted PDFs and files with no text remain explicit states. Compact migration also preserves native attachment state when present.

## Correctness fixes

- Paging orders every folder by `(UTC received timestamp, lowercase StoreID|EntryID)` descending. It gathers a local top page including all boundary ties from each folder, globally merges, then limits. Minute-level Outlook Restrict is only a coarse optimization; exact filtering is enforced locally even if Restrict fails.
- The cursor carries both timestamp and identity, including fractional seconds. Legacy ISO-only bounds still work. Old full-safe checkpoints are invalidated once so a formerly skipped folder is not treated as complete; existing rows are kept.
- Calendar overlap uses `End > windowStart AND Start < windowEnd`. Recurring expansion remains enabled. `GetAssociatedAppointment(false)` cannot request insertion into the default calendar. Mail reply counts are kept as evidence, never promoted to an authoritative attendee roster.

Stable keyset pagination is not a transactional snapshot of a changing mailbox. Newly arrived or moved messages can appear above an existing cursor; a later incremental pass reconciles them. IDs can change on cross-store moves. No claim of exact-once mailbox replication is made.

## Limitations

- The inherited OpenClaw CLI remains separate and unchanged in signature. Its source and runtime can continue in WSL, while native ChatGPT retrieval is independent. During migration there are deliberately two indexes; old data is not deleted without permission.
- Native incremental sync indexes supported new attachment text. Scanned/image PDFs require external OCR (not provided); encrypted PDFs, unsupported/binary Office formats and over-limit files are skipped explicitly. Mailbox deletion/move reconciliation and a transactional mailbox snapshot are not provided.
- Search is retrieval, not automatic persistent ChatGPT memory. Email content is untrusted and is returned to the model only on tool calls.
- Client registration and direct protocol acceptance do not prove that an already-running chat has hot-loaded the server. A visible client tool call is the final integration check.
- Outlook COM requires an interactive Windows desktop session and classic Outlook. Offline indexed search works without Outlook running; live mail/calendar calls require it.
