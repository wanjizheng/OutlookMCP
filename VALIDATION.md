# Validation — 2026-10-17

Validated locally on native Windows Python 3.14, classic Outlook, and the existing WSL OpenClaw deployment. No email was sent, marked read, deleted, or edited, and no meeting was created/cancelled for acceptance.

## Passed

- 14 Windows regression tests: same-second/fractional cursor identity, global ordering across folders, legacy ISO bounds, failed Restrict handling, calendar overlap boundaries, recurrence setup, no calendar insertion request, native sync resume, exclusive writer lock, read-only SQLite connections, FTS quoting, explicit hybrid fallback, compaction invariants, tool allowlist, and preserving client config bytes.
- 19 existing/extended OpenClaw RAG tests, including old checkpoint invalidation and ordinary-token Windows script path translation.
- Native and deployed OpenClaw live two-page reads: four unique messages; UTC times and StoreID retained.
- Native Outlook email and seven-day calendar reads.
- Native local GGUF inference: 768-dimensional vectors; native synthetic/bounded live index populated and vector retrieval returned results.
- Real MCP STDIO initialization, five tools listed, successful status/list/read/calendar/search calls, and rejection of a nonexistent send tool.
- Full `Install.ps1 -WithEmbeddings`: dependency locks, model checksum verification, private data ACLs, regression tests and doctor.
- Python dependency consistency and Node audit (zero reported vulnerabilities at validation time).

All live acceptance reports contain counts/status only. Mailbox data, models, indexes, private logs and client configuration backups are excluded from Git.

## Integration boundary

The native `outlook_local` STDIO server is registered in the local desktop host configuration, with an absolute Windows Python command and Windows working directory. Direct protocol acceptance passed. The currently running chat has not hot-loaded its new tools; reload the MCP connection or start a new local chat, then invoke `outlook_status`. Registration alone is not claimed as an observed in-chat call.

The Windows runtime invokes native PowerShell/Node and reads a local Windows index. WSL was used only for the explicitly authorized compatibility deployment/tests and one-time source backup. The user's running WSL/OpenClaw services were not shut down to demonstrate this boundary.

## Coverage limits

An imported index preserves existing indexed text and vectors; it does not prove that the old mailbox scan covered every message. The corrected full backfill can repair prior omissions. Native sync now adds supported new attachment text as described below. Mailbox deletion/move reconciliation, OCR and automatic ChatGPT memory are not claimed.

The first migration used a full consistent SQLite backup. Full verification on the cross-filesystem snapshot was slow; that owned validation process was stopped, and a native compact conversion preserves retrieval data and verifies the final native database. The redundant intermediate snapshot is cleaned only after final acceptance.

## Final native index acceptance

The compact Windows index passed SQLite quick_check and preserved 38,576 emails, 111,031 chunks and 111,031 valid vectors. Its main database is 993,718,272 bytes versus the consistent legacy snapshot at 4,954,951,680 bytes (about 80% smaller). On the final Windows index, hybrid retrieval returned results with no fallback. These counts describe the imported snapshot, not total mailbox coverage.

## Native attachment extraction follow-up

- Windows suite expanded to 23 tests, covering all seven supported formats, encrypted/no-text/corrupt PDFs, page/text/file/ZIP expansion limits, DTD rejection, duplicate/path-like filenames, export count/type limits, temporary cleanup, unchanged-body backfill, content changes, body-only preservation, and failed-page checkpoint rollback.
- Bounded live Outlook acceptance extracted one existing attachment into a separate local test index, embedded its text and retrieved it with hybrid search without fallback. The guard observed only native PowerShell, Python and Node child invocations: zero WSL calls, zero temporary raw files remaining. No mailbox fixture was created or sent.
- MCP retains exactly five read-only tools. The administrative attachment export script is not part of its bridge operation allowlist or tool registry. Real MCP STDIO acceptance is repeated against the existing production index.
- The new attachment pipeline has been validated on a bounded live sample, not a full-mailbox attachment backfill. Imported historical attachment text remains intact until explicitly re-synced; skipped formats, encryption, limits and no-text/OCR cases are reported separately. Native attachment state is preserved by compact migration.

## Historical backfill and first stdio hybrid query

The subsequent native historical attachment backfill covers the existing local index, not every item in the mailbox. Detailed counts, extraction exceptions, search examples and consistent index backups remain in ignored local artifacts. No mailbox contents, attachment filenames or queries are published.

Acceptance exposed a Windows Python 3.14 startup deadlock: the first hybrid/vector MCP request loaded NumPy's native extension after stdio worker threads had started. Direct CLI retrieval did not reproduce it. NumPy now initializes before the MCP transport starts; tool names, read-only permissions and OpenClaw interfaces are unchanged. A fresh-process stdio regression exercises the first hybrid request without requiring Outlook or a downloaded model. Existing clients must reload the MCP connection to load this code change.

The expanded 24-test native Windows suite passed. Real stdio requests using the registered client command retrieved three distinct historical attachments through lexical and hybrid search, with vector-only retrieval also verified. SQLite/FTS integrity, chunk-to-vector coverage and unchanged incremental-sync behavior passed against the local backfill. These are protocol and local-data checks; an existing desktop connection still needs a reload, and an in-chat invocation is not claimed.
