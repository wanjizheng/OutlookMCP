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

## Historical prototype: confirmed sending and private tunnel follow-up

The native-popup prototype below was tested, then explicitly rejected by the
user. It is not the final implementation; see the final chat-confirmation section.

After a client restart, a real local chat called `outlook_status`, searched a
historical attachment, and read its parent message to verify the keyword was not
in the message body. This establishes the local-chat connection, separately from
ordinary cloud Chat plugin acceptance.

The Windows suite now passes 33 tests. New coverage includes send opt-in,
header/recipient validation, immutable preview snapshots, expiry, single-attempt
consumption, timeout uncertainty, mocked cancellation/discard without Save/Send,
mocked approved submission with exact account and fields, recipient-resolution
mismatch blocking, test-mode isolation from Outlook, and write annotations.
A fresh real STDIO server with sending enabled lists seven tools, prepares a
synthetic message and rejects invalid send requests. It does not submit real
mail. A test initially assumed the SDK rejects extra tool arguments; the SDK
ignores extras. Confirmation is therefore enforced in the native worker, not by
the presence/absence of an LLM-supplied `confirmed` argument.

The original five-tool default STDIO smoke passed again against native Outlook
and the existing RAG. Python dependency consistency and diff whitespace checks
passed. The official Windows tunnel binary was downloaded and verified against
its release SHA-256; its profile uses the same checkout and a loopback-only
health listener. Credentials and account-specific tunnel/workspace identifiers
remain in private ignored artifacts, not this repository.

Cloud plugin installation, tunnel readiness and ordinary Chat search/confirmation
are separate live acceptance gates. Do not claim them complete until their actual
observed results are recorded. No real email delivery is authorized by these tests.

The authorized private tunnel was subsequently created and associated with the
target personal workspace. A restricted runtime key was entered by the user into
the native masked dialog and encrypted with DPAPI. `doctor` passed, metadata was
successfully fetched with that key, and the loopback `/readyz` returned `ready`.
Creating the custom plugin from the embedded browser was rejected by ChatGPT:
`Custom apps aren't allowed in this context. Check your workspace permissions or security settings`.
Plugin and security settings also failed to load fully in that browser. The user
was given the exact existing-tunnel form values to complete the same action in
their normal ChatGPT client. This rejection is a host/account-context gate, not
an observed MCP transport failure. Ordinary Chat acceptance remains pending until
plugin creation and actual calls succeed.

### Historical popup prototype: ordinary Chat acceptance in Chrome

The embedded-browser context rejection was resolved by signing into the same
ChatGPT account in Chrome. The personal cloud plugin was created through the
normal Tunnel form and installed; the UI reported it connected and showed it
under Installed. No API bypass or workspace/security-policy change was used.

Using the plugin's **Try in chat** button opened ordinary **Chat**, with the
plugin selected (not local Codex or Work). That chat actually called status,
BM25 historical-attachment search, parent-message read and message preparation.
The keyword appeared in the attachment-index snippet and was absent from the
complete message body. Private message IDs, query/attachment examples, chat URL
and screenshots remain in ignored artifacts; no mailbox evidence is published.

The chat displayed the full immutable synthetic sender/To/CC/BCC/subject/body/
attachments preview and stopped to request approval without calling submission.
After an explicitly authorized **confirmation-only** follow-up, the chat called
`send_prepared_email` with `confirmation_test=true`. The user personally clicked
**Confirm test only** in the native Windows window. The chat received:

```json
{"status":"confirmation_test_accepted","sent":false,"test":true}
```

This is observed ordinary Chat + private tunnel + Windows confirmation acceptance.
The positive approval path is safely exercised in test mode, which does not
instantiate Outlook; actual COM submission remains covered by mocks. No real
email was sent, no delivery is claimed, and a real message still requires its
own complete preview and explicit human confirmation. All 33 Windows tests,
the original read-only live STDIO smoke and dependency checks pass.

## Final implementation: confirmation in ChatGPT, existing PowerShell sender

The user explicitly chose full preview and approval inside ChatGPT, followed by
direct Outlook submission with no custom desktop popup. The Python/Tk sending
worker was removed. The MCP now reuses the existing skill's Windows
`send_outlook_mail.ps1`; shared COM logic validates exact sender account and
resolved SMTP recipients. Legacy default-draft and preview interfaces are kept.
The existing deployed OpenClaw checkout was not changed in this follow-up.

The personal plugin was set to `ask_before_writes` (read without asking, ask
before write) and the tool remains annotated as a destructive, non-idempotent
write. The model must show every preview field, stop and obtain explicit approval
in a subsequent user turn. This confirmation belongs to the ChatGPT host/client;
the backend does not claim to independently verify a human chat reply. It binds
the immutable preview, enforces expiry/single-attempt consumption, checks the
explicit account and recipient resolutions, and never automatically retries.

All 32 revised Windows tests passed. Real PowerShell tests exercise the same COM
function with fake Outlook objects: exact fields/To/CC/BCC/account, mismatched
recipient blocking, missing-account blocking, legacy draft default and no-write
preview. The real opt-in STDIO test prepares and consumes a confirmation-only
message, verifies `sent=false,test=true`, and rejects replay. No Outlook process
or custom window is created in this test mode. Actual mail delivery remains
untested and no real test mail is authorized.

After restarting only the owned tunnel process and refreshing plugin tools, a
fresh ordinary Chat again retrieved the historical index evidence and prepared
the synthetic message. It displayed every preview field, stated that native
Windows confirmation is no longer required, and stopped without submission.
On a later explicitly authorized confirmation-only turn, ChatGPT displayed its
own **Allow once / Deny / Always allow** approval card before the write tool.
This verifies that `ask_before_writes` is active in the real cloud Chat client,
not merely described in the server metadata. The final test result is recorded
below after the user operates that card.

The approval card was subsequently resolved by the user (the assistant did not
click it). The ordinary Chat tool invocation completed and returned:

```json
{"status":"chat_confirmation_test_passed","sent":false,"test":true}
```

The returned immutable synthetic message matched preparation. There was no
desktop popup and no Outlook invocation in this confirmation-only branch. This
completes ordinary Chat retrieval + preview/stop + ChatGPT-native approval +
no-send confirmation acceptance for the final implementation. Real delivery is
still intentionally not claimed. Credentials, mailbox evidence, account IDs,
private chat links and screenshots remain excluded from the public commit.

## Windows logon startup and recovery (2026-10-17)

Installed a native Scheduled Task for the current interactive Windows user at
limited privilege, without storing a Windows password. Read back its logon
trigger, indefinite one-minute repetition trigger, IgnoreNew instance policy,
unlimited execution duration, battery settings and explicit opt-in send argument.
The task invokes the checkout's pythonw supervisor, not WSL. The existing local
STDIO clients and OpenClaw deployment remain untouched. The ChatGPT plugin's
ask-before-write setting and full-preview/subsequent-user-turn confirmation
contract remain unchanged; no real email was sent.

Live recovery checks:

- Killed only the owned tunnel/backend process tree. The supervisor relaunched
  it after the configured backoff; the loopback endpoint returned `ready` again.
- Stopped the owned task and verified that its supervisor and tunnel processes
  were gone. Restart restored the service. Windows Job Object ownership prevents
  abandoned descendants; a second supervisor invocation exited without spawning
  a second tunnel.
- Force-killed the supervisor. RestartOnFailure alone did not requeue this
  manually started task during the observation window. Added the independent
  one-minute trigger, repeated the failure, and observed a different supervisor
  PID and a live `ready` response after the next scheduled trigger.
- After recovery, the connected cloud ChatGPT plugin successfully called
  `outlook_status` (native Windows/RAG, no WSL) and retrieved the known historical
  attachment evidence with `search_email`. Transitional requests included an
  initialization error and, after a later restart, one timeout; subsequent reads
  succeeded. Explicit STDIO initialized notification is now enabled. This does
  not promise uninterrupted calls or authorize replaying write operations.

The first supervisor startup exposed a pywin32 unnamed-job argument mismatch;
corrected it before the successful live checks. Local lifecycle diagnostics,
PIDs and private result data remain under ignored `artifacts/`.

The 180-second unready grace, capped exponential retry, successful-probe reset,
loopback-only probe, HTTP failure handling, owned-child termination and default
read-only/explicit-send tunnel bootstrap have regression coverage. Real STDIO
mail/calendar/RAG smoke passed with the original five read tools, including
rejection of an unavailable send tool. Actual reboot/logoff, physical network
disconnection and real mail delivery were not performed; logon trigger settings,
scheduled recovery and process-failure recovery were verified on the running PC.
All 39 Windows regression tests passed, as did dependency consistency, generated
bridge reproducibility and whitespace checks. Stop now disables the periodic
trigger as well as terminating the task; Restart re-enables it. Runtime state
and credentials are excluded from the commit.

## Explicit manual RAG refresh (2026-10-18)

The user explicitly requested a bounded ChatGPT manual index refresh. Added a
separate opt-in and two tools: refresh_email_index and get_refresh_status.
The installed Windows logon supervisor now retains --enable-send and
--enable-refresh. Default local STDIO remains five read-only tools; OpenClaw
interfaces and the sending preview/approval contract are unchanged.

All 51 Windows regression tests passed. New coverage includes asynchronous job
admission/coalescing, fixed worker command and credential stripping, read-only
status/path rejection, PID reuse/interruption, writer contention, actual
attachment ingest with fake Outlook reads, committed-page failure, partial
model/coverage results, recent-only bounded vector backlog and a real Windows
worker timeout that kills its descendant. Opt-in STDIO lists seven tools without
sending and rejects invalid status IDs; combined refresh+sending lists nine.
Explicit dependency auto-install is disabled for the refresh worker. Default
real STDIO mail/calendar/search smoke, bridge reproducibility and pip check pass.

Live ordinary Chat acceptance used the installed private cloud plugin, not a
local-only call. The first management Refresh tools attempt failed; retry
succeeded and the UI listed seven read/two write tools. Existing ask-before-write
permission was retained. A fresh ordinary Chat requested one bounded refresh,
displayed the native Refresh Email Index approval card, and after Allow once
started one job and polled that same job until completed. No sending tool was
called, no email was sent and no mailbox write was performed.

The real job finished in about 3 minutes 41 seconds, within its ten-minute cap:

- InboxAll: 76 messages scanned; Sent: 5; neither scope had additional pages.
- Attachment extraction: 5 ok; 56 unsupported skips, reported explicitly.
- Emails: 38,776 before and after. Chunks: 132,337 -> 132,348.
- Vectors: 132,337 -> 132,348; 31 vectors computed after changed chunks were
  rebuilt; zero pending recent vectors on completion.
- Native mail sync timestamp advanced; final completed state was visible in
  ordinary Chat. This proves only the fixed 72-hour scope and supported text,
  not complete history or support for skipped formats. Full-history reindex and
  periodic synchronization were neither exposed nor run.

The first cold baseline count was relatively slow; asynchronous status remained
responsive and correctly reported running rather than prematurely claiming
completion. Private job records, screenshots, chat URL, mailbox material and
credentials remain in ignored local directories and are not part of this commit.


## Resumable incremental and verified full sync (2026-10-18)

This supersedes the earlier fixed 72-hour refresh contract; historical validation
above remains evidence of that older version. Default STDIO still exposes exactly
five read-only tools; opt-in refresh without sending exposes eight, and combined
sending/refresh exposes ten (seven read, three write). Sending confirmation and
OpenClaw scripts/interfaces are unchanged. Native Windows runtime only.

All 67 Windows regression tests passed, plus the final seven-test refresh lifecycle
check. Coverage includes legacy successful-boundary migration, long-gap and old-dated
late arrival, empty baseline, atomic per-message progress/restart, changed messages,
known-store pinning, stable/new-ID moves, same-subject copies, strict missing/duplicate
source reads, failed folders, source mutation during verification/backup, cached
Exchange refusal, SQLite backup recovery evidence, backup failure, atomic cleanup
rollback, unsupported metadata without exports, and bounded resumable vector batches with committed-progress callbacks.
The real default MCP mail/calendar/search smoke passed, as did dependency consistency,
generated bridge reproducibility and whitespace checks.

Important failures found and corrected during development:

- Outlook Table EntryID values were short-term and did not match the historical
  index. The source now reads PR_LONGTERM_ENTRYID_FROM_TABLE (verified against
  the real MailItem.EntryID), without opening every MailItem for inventory.
- Cached Exchange exposed only approximately one year of history despite being
  connected. Absence from this view is never proof of deletion; full cleanup is
  blocked until online server coverage is verified.
- SQLite backup connections initially remained open on Windows; fixed explicit
  closure. Rollback, backup-failure and backup-readable-content tests pass. A 1,100-ID reconciliation test verifies bulk
  cleanup without SQLite parameter-limit failures or repeated full FTS scans.
- Initial migration queued historical gaps during incremental updates. It now
  accepts only a fully completed legacy window as an incremental boundary, reports
  deferred history separately and tracks seen IDs to catch later old-dated arrivals.
- The first live resume reached complete incremental text coverage but its old
  45-second embedding call timed out. It correctly reported partial with 660 pending
  vectors and retained the successful mail watermark. A subsequent single 64-chunk batch also exceeded
  180 seconds on this host. Manual embedding now uses committed 16-chunk batches
  (up to 64/job), with a 180-second call budget per batch. Old CLI/query defaults remain.
- A narrow vector-ID index avoids scanning embedding blobs just to count progress;
  new-message ingestion avoids a needless full FTS deletion scan.

Ordinary Chat acceptance used the installed private tunnel plugin after refreshing
its tools, with ask-before-write permissions unchanged and Allow once on local-index
operations. No send tool was called and no Outlook mailbox write was performed.
The initial worker was deliberately stopped after 79 committed new messages.
Ordinary Chat observed interrupted, then a subsequent explicit update reused the
same durable run boundary and retained those messages. It scanned 31 folders /
7,724 locally visible messages, reused 5,992 and processed one, advancing the mail
success watermark while reporting 1,731 historical gaps separately. The vector
timeout above was preserved as a partial result, not represented as full success.

A separate real full-sync invocation completed its local scans in about 50 seconds
and returned partial, coverage_complete=false, server_history_verified=false,
cleanup_blocked_reason=cached_exchange_cannot_prove_server_absence. All 32,836
unverified absent entries were retained: removed=0, backup_created=false, full-sync
watermark not advanced. This proves the live safe-refusal path; actual full online
server reconciliation/deletion is covered by synthetic-source tests, not claimed as
live acceptance. Outlook account/cache settings were not changed. COM does not
provide a mailbox-wide atomic snapshot; mutations after the final scan belong to
the next sync.

Private source inventories, attachment material, job IDs, screenshots, chat URLs,
models, database and credentials remain excluded from Git.


Final ordinary Chat vector acceptance used the revised 16-chunk batches. The
same job reported committed progress 16 -> 32 -> 48, then embedded=64 with no
embedding error. Vectors increased 132,348 -> 132,412 and pending vectors fell
660 -> 596. Mail count remained 38,856, chunks 133,008, removed=0; the successful
incremental watermark advanced. The terminal status correctly remained partial
because historical coverage and vector backlog are unfinished. This is evidence
of working bounded vector continuation, not a claim that every vector/history
record is complete. The source and model ran as native Windows processes.

## 2026-10-19 confirmed routine Outlook operations

Native Windows regression: 76 tests passed. Additional native PowerShell mocks
exercise create/update/cancel and invitation responses, explicit recipient roles,
UTC conversion, selected sending account, recurring occurrence/series targeting,
stale-preview rejection and single submission. Reply/reply-all/forward native
mocks verify source version and exact approved text/account. No tests send real
mail, invitations or responses. Default live STDIO smoke passed all five read-only
tools and rejected unknown writes. Generated compatibility libraries are unchanged;
pip check passed. Real native read of an existing calendar item succeeded without
Save/Send. The Windows scheduled backend was restarted with both existing opt-ins;
private tunnel readyz reports ready.

Ordinary Chat plugin refresh exposes 16 tools (12 read/preparation, 4 writes).
A real ordinary Chat call to prepare_calendar_action created the complete synthetic
meeting preview and stopped for subsequent approval. Commit acceptance is pending
human confirmation of the explicitly non-sending confirmation_test=true flow;
no real calendar write or email/invitation send is claimed. Private screenshots,
source IDs, account metadata and chat URLs stay under ignored artifacts only.

Zoom link creation is not configured. Ordinary Outlook meeting invitations require
no Zoom/Teams/Meet service. Existing links may be provided/preserved. No Outlook
account mode, mailbox contents, RAG index or OpenClaw interfaces were changed.

## Recheck of the real recurring meeting flow (2026-10-19)

The earlier synthetic create preview was insufficient to prove existing recurring
meeting edits. A fresh ordinary Chat test reproduced failure in read_calendar_event
before preparation; no commit was called. The response also confused single with
occurrence. A local native read using explicit occurrence succeeded, so the generic
COM error alone does not establish a unique underlying cause.

The repair adds Windows listing StoreID/recurrence/UTC-instance metadata, bounded
one-hour opaque references pinning a listed instance, an enum scope schema, the
narrow single-plus-explicit-start alias, and direct handling of already expanded
occurrences/exceptions. References cannot override scope or broaden to series.
Errors now distinguish ID resolution and instance resolution; read errors do not
imply a send was attempted. OpenClaw runtime sources/signatures are unchanged.
Generated Windows libraries and the compatibility test fixture remain reproducible.

79-test full regression passed, plus the new MCP schema/reference dispatch test.
Real Windows STDIO acceptance listed the target meeting, read it through event_ref,
prepared one added attendee (six to seven), verified subject/body/start/end/location
unchanged in the preview, and re-read an identical source snapshot. No commit,
Save, Send or real invitation response was executed. Default five-tool live smoke
and pip check passed. Windows background task was restarted; tunnel ready is true.

Plugin permission was found set to Allow all and restored to Allow read-only tools
(read without asking; ask before changes), consistent with the user's requirement.
Fresh ordinary Chat end-to-end retest is recorded below after completion. Private
meeting IDs, body/link material, diagnostics and screenshots are excluded from Git.

Fresh ordinary Chat retest PASSED after background restart and plugin refresh:
list_calendar returned the dated event_ref; read_calendar_event(event_ref=...) and
prepare_calendar_action(event_ref=...) both succeeded for the intended occurrence.
The preview retained all six attendees and existing title/time/body/location/link,
adding only the requested person (preview total seven). Chat stopped before commit.
A subsequent independent native read matched the original full snapshot exactly;
live attendee count was still six. This establishes real recurring-event read and
preparation acceptance, not actual Send/Save or write approval acceptance. No real
write was authorized or performed. The prior generic failure was not hidden or
reclassified as successful; the old synthetic-only evidence remains distinguished.

## Actual Outlook submission rejection investigation (2026-10-19)

The user's original separate chat was located and its error evidence retrieved.
It reported Windows Outlook's native "You don't have appropriate permission to
perform this operation", wrapped as INVALID_ARGUMENT. The old response omitted
HRESULT and operation stage, so the exact historical failing COM member cannot be
established. This is not evidence that the user failed to approve ChatGPT's card.

Independent native reads found both requested occurrences still had the original
six attendees and the original modification timestamp; the target attendee was
absent. A bounded read of the primary Sent Items and Outbox (up to 1,000 recent
rows per folder, two-day window) found no matching meeting update. This is local
cached-view evidence, not server-wide proof of non-delivery. PR_ACCESS on the
actual event was 7 (includes modify), calendar was 63; read SendingAccount matched
the requested account. Unsaved throwaway-object property probes (no Save/Send)
worked. None of these tests proves a real Send would succeed.

Removed redundant reassignment of an existing event's inherited sending account,
validated access/account/clean closed state before modifications, and preserved
stage/HRESULT/exception type and mutation/send flags through the Python boundary.
Pre-send failures attempt Close(olDiscard) only on this operation's owned unsaved
item; a returned discard call does not prove a transactional server rollback.
After Send starts or Respond can have calendar side effects, no rollback or retry
is claimed. Returned snapshots include Saved/status, explicitly not send evidence.

82-test regression and targeted native mocks passed. Mocks inject 0x80070005 before
Send and at Send, verifying phase preservation, pre-send discard, no discard/retry
after uncertain Send, and rejection of unsaved/read-only state before mutation.
A throwing setter ensures existing account reassignment cannot recur. No real
meeting write/send was performed during this diagnostic turn. Actual live sending
remains unverified and requires a fresh concrete preview and human confirmation.

Subsequent real-send acceptance: the user explicitly confirmed ONE fresh concrete
retry after the new immutable preview. commit_calendar_action returned success,
status=submitted_to_outlook, send_attempted=true, send_returned=true, sent=true and
delivery_verified=false. No second attempt was made. Independent native reads
found seven attendees on the intended occurrence and the original six on the next
weekly occurrence; original subject/body/start/end/location matched the before
snapshot exactly. Sent Items contained a matching IPM.Schedule.Meeting.Request,
Sent=true, target included, with appointment properties matching the requested
instance. Outbox showed no matching pending item. These are submission/sent-copy
and local calendar evidence, not recipient delivery or server-wide atomic proof.
An occurrence proxy's Saved flag was false after success, illustrating why that
flag alone must not be used as evidence of either persistence or notification.

The prior rejection lacked a phase/HRESULT; removing redundant existing-account
assignment is a compatibility correction validated by the successful retry, not
retrospective proof of the exact old failing member. No security policy, Outlook
account settings or WSL runtime was changed. Windows and legacy GitHub CI passed
for the implementation commit. All private IDs, recipients, body, sent timestamps,
confirmation tokens and chat screenshots remain excluded from Git.
