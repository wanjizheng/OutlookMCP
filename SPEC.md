# OutlookMCP: confirmed sending and ordinary Chat integration

The user has approved reusing the existing Windows Outlook COM backend and local
RAG for email/attachment retrieval plus sending composed email after explicit
human confirmation. No WSL, Graph login or cloud mailbox copy is required.

Flows: search/read -> compose in conversation -> prepare immutable message ->
show recipients, sender, subject and full body -> human confirmation -> submit
to Outlook. Cancel/timeout/absent desktop must fail closed. No automatic retries
after submission. No mailbox drafts, deletion, mark-read or meeting writes.

Default STDIO remains five read-only tools. Sending is separately opted in with
OUTLOOK_MCP_ENABLE_SEND=1, adding prepare_email and send_prepared_email. The
ChatGPT must stop after the preview and obtain explicit approval in a subsequent
user turn before calling the non-idempotent write tool. The user explicitly
chose no custom desktop window. Reuse the existing Windows PowerShell sender;
the backend validates the immutable message, SMTP account/recipients and
single-attempt consumption. User confirmation is a host/client responsibility,
not something the backend can prove from a model-supplied boolean.
The first sending version supports plain text with explicit SMTP addresses,
including CC/BCC, and a required sender account. Attachments are not supported.

Ordinary Chat uses a personal custom MCP plugin through OpenAI Secure MCP Tunnel
when account permissions allow it. Reuse the same Windows Python, tools, models
and database; no public unauthenticated endpoint. Host connection/confirmation
and conversation confirmation apply. End-to-end acceptance requires a
real ordinary Chat tool call; protocol tests alone are not acceptance.

Tests must not send real email. Confirm cancellation and preview-only validation
with synthetic example.invalid recipients. Any real send requires a separately
approved concrete recipient and message; this project task does not authorize it.

## User-authorized resumable incremental and full index sync

Supersedes the previous fixed 72-hour refresh contract. With the existing
OUTLOOK_MCP_ENABLE_REFRESH opt-in, refresh_email_index resumes durable per-message
progress from the last successful sync; full_sync_email_index checks the whole
primary mailbox normal-folder scope. Neither accepts commands, paths or scopes.
Default STDIO remains five read-only tools; OpenClaw interfaces are unchanged.

A strict native source reader inventories every normal primary-mailbox folder
(excluding Deleted/Junk/Drafts/Outbox and virtual search folders). Canonical
MailItem EntryIDs are resolved from short-lived Table IDs. Changed and unknown
messages are read with attachments; identical source versions reuse cached text
and vectors. Per-message extraction and source checkpoint commit atomically.
All-folder successful watermark is separate from per-page legacy timestamps.
An interrupted job preserves its original start boundary and committed work.

A second strict inventory must match before success. Full cleanup is permitted
only when server history is verifiably complete, and only after a SQLite backup
and another matching scan. Cached Exchange visibility is NOT proof of server
absence: preserve historical index entries and report partial/full cleanup
blocked, never silently delete older server mail. A primary-store change fails
closed. Moves with stable IDs update location; moves with new IDs are ingested
before old IDs can be removed, without unsafe subject-based identity matching.

Single admission/writer locks; 24-hour process-tree budget; explicit user retry
only. Vector filling is bounded to 64 chunks and reported separately from mail
coverage. Pending vectors do not erase successful text checkpoints. All code,
data and models run on Windows, with no WSL. No Outlook mailbox mutations.
Sending still requires full preview followed by explicit subsequent human
confirmation; no real email is authorized by synchronization acceptance tests.
Migration accepts only completed legacy 72-hour jobs with both InboxAll/Sent
fully paged as an incremental boundary, not as evidence of historical coverage.
Older unmaterialized gaps are deferred to full sync and reported separately.
A complete inventory of seen IDs detects late arrivals with old timestamps on
subsequent updates. Known changed source versions are never skipped solely due
to timestamps. In cached mode full sync fills the incremental interval but defers
unverified historical reconciliation, never deleting absent IDs.

Manual embedding fills up to 64 chunks per job in committed 16-chunk batches
with a 180-second model-call timeout per batch. Existing CLI batch and
query timeout defaults remain unchanged. A narrow local vector-ID index prevents
status counting from scanning embedding blobs. New email inserts avoid an
unnecessary full FTS deletion scan; existing-row updates retain full cleanup.
