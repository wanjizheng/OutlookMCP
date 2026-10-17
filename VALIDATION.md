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
