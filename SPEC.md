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
