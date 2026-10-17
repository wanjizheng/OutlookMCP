# Architecture and validation boundaries

ChatGPT → native Python MCP STDIO process → fixed-operation PowerShell bridge → classic Outlook COM, or read-only SQLite RAG connections. No process invokes wsl.exe, Bash, systemd or OpenClaw at runtime. The optional administrative importer can run once on a Linux source to make a consistent snapshot.

The local store reuses the legacy `emails`, `email_chunks`, `email_chunks_fts`, and `email_chunk_vectors` schema. Chunks use 900 characters / 150 overlap. Hybrid search combines FTS5 rankings with bounded NumPy float32 cosine batches and reciprocal-rank fusion. It uses the same GGUF embedding model as the imported index, on native Windows Node. Failures are reported as explicit BM25 fallback. Query results contain message IDs, dates, excerpts and sync timestamps; attachments already indexed by OpenClaw remain searchable.

Only the administrative CLI creates/updates the database. MCP connections use SQLite `mode=ro` and `PRAGMA query_only=ON`. No synchronization or arbitrary script/file tool is registered. Local index updates do not authorize writes to the mailbox.

## Correctness fixes

- Paging orders every folder by `(UTC received timestamp, lowercase StoreID|EntryID)` descending. It gathers a local top page including all boundary ties from each folder, globally merges, then limits. Minute-level Outlook Restrict is only a coarse optimization; exact filtering is enforced locally even if Restrict fails.
- The cursor carries both timestamp and identity, including fractional seconds. Legacy ISO-only bounds still work. Old full-safe checkpoints are invalidated once so a formerly skipped folder is not treated as complete; existing rows are kept.
- Calendar overlap uses `End > windowStart AND Start < windowEnd`. Recurring expansion remains enabled. `GetAssociatedAppointment(false)` cannot request insertion into the default calendar. Mail reply counts are kept as evidence, never promoted to an authoritative attendee roster.

Stable keyset pagination is not a transactional snapshot of a changing mailbox. Newly arrived or moved messages can appear above an existing cursor; a later incremental pass reconciles them. IDs can change on cross-store moves. No claim of exact-once mailbox replication is made.

## Limitations

- The inherited OpenClaw CLI remains separate and unchanged in signature. Its source and runtime can continue in WSL, while native ChatGPT retrieval is independent. During migration there are deliberately two indexes; old data is not deleted without permission.
- Native incremental sync currently indexes message bodies; imported attachment text remains searchable. New attachment text extraction and automatic deletion reconciliation are not claimed unless enabled and validated in the version being installed.
- Search is retrieval, not automatic persistent ChatGPT memory. Email content is untrusted and is returned to the model only on tool calls.
- Client registration and direct protocol acceptance do not prove that an already-running chat has hot-loaded the server. A visible client tool call is the final integration check.
- Outlook COM requires an interactive Windows desktop session and classic Outlook. Offline indexed search works without Outlook running; live mail/calendar calls require it.
