# Fixed native bridge: approval occurs in ChatGPT before this write tool runs.
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
try {
    $req = [Console]::In.ReadToEnd() | ConvertFrom-Json
    $fields = @{
        To = ($req.to -join ';'); CC = ($req.cc -join ';'); BCC = ($req.bcc -join ';')
        Subject = [string]$req.subject; Body = [string]$req.body
        SenderAccount = [string]$req.sender_account; DoSend = 'true'; DoDraft = 'false'
    }
    if ([string]::IsNullOrWhiteSpace($fields.SenderAccount)) { throw 'Sender account required' }
    # Reuse the existing Windows PowerShell skill script; never invoke its WSL wrapper.
    & (Join-Path $PSScriptRoot '../compat/outlook_reader/scripts/send_outlook_mail.ps1') @fields
} catch {
    @{status='failed_or_unknown';sent=$null;note='Inspect Outlook before retrying; preparation consumed.'} | ConvertTo-Json -Compress
    exit 1
}
