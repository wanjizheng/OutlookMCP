$ErrorActionPreference='Stop'
[Console]::InputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
function Get-MailResponseOutlook { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
function Submit-MailResponse($p) {
    if ($env:OUTLOOK_MCP_ENABLE_SEND -ne '1') {throw 'Sending disabled'}
    $app=Get-MailResponseOutlook
    $ns=$app.GetNamespace('MAPI')
    $accounts=@($ns.Accounts | Where-Object { $_.SmtpAddress -eq $p.sender_account })
    if ($accounts.Count -ne 1) {throw 'Sender account missing or ambiguous'}
    $source=$ns.GetItemFromID($p.source.entry_id,$p.source.store_id)
    if ($source.Class -ne 43 -or $source.LastModificationTime.ToUniversalTime().ToString('o') -cne $p.source.modified -or [string]$source.Subject -cne $p.source.subject -or [string]$source.Body -cne $p.source.body) {throw 'Source changed; prepare again'}
    switch ($p.mode) {
        'reply' { $mail=$source.Reply() }
        'reply_all' { $mail=$source.ReplyAll() }
        'forward' { $mail=$source.Forward() }
        default {throw 'Invalid response mode'}
    }
    while ($mail.Recipients.Count -gt 0) {$mail.Recipients.Remove(1)}
    while ($mail.Attachments.Count -gt 0) {$mail.Attachments.Remove(1)}
    foreach ($pair in @(@{type=1;values=$p.to},@{type=2;values=$p.cc},@{type=3;values=$p.bcc})) {
        foreach ($address in $pair.values) {$r=$mail.Recipients.Add($address);$r.Type=$pair.type}
    }
    if (-not $mail.Recipients.ResolveAll()) {throw 'Unresolved recipients'}
    $mail.SendUsingAccount=$accounts[0]
    $mail.Subject=$p.subject
    $mail.Body=$p.body
    $mail.Send()
    return @{status='submitted_to_outlook';sent=$true;mode=$p.mode}
}
if ($MyInvocation.InvocationName -ne '.') {
try { Submit-MailResponse ([Console]::In.ReadToEnd() | ConvertFrom-Json) | ConvertTo-Json -Compress } catch {
    @{error='Response submission failed or uncertain; inspect Outlook before retrying.'} | ConvertTo-Json -Compress
    exit 1
}

}
