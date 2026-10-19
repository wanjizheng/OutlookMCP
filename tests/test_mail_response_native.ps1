$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '..\scripts\mail_response.ps1')
$env:OUTLOOK_MCP_ENABLE_SEND='1'
$script:sends=0
$script:method=''
$script:recipients=[pscustomobject]@{Count=0}
$script:recipients | Add-Member ScriptMethod Add { param($a) return [pscustomobject]@{Type=0} }
$script:recipients | Add-Member ScriptMethod ResolveAll {return $true}
$script:outgoing=[pscustomobject]@{Recipients=$script:recipients;Attachments=[pscustomobject]@{Count=0};SendUsingAccount=$null;Subject='';Body=''}
$script:outgoing | Add-Member ScriptMethod Send {$script:sends++}
$script:source=[pscustomobject]@{Class=43;Subject='Original';Body='Source';LastModificationTime=[datetime]'2026-10-19T10:00:00Z'}
$script:source | Add-Member ScriptMethod Reply {$script:method='reply';return $script:outgoing}
$script:source | Add-Member ScriptMethod ReplyAll {$script:method='reply_all';return $script:outgoing}
$script:source | Add-Member ScriptMethod Forward {$script:method='forward';return $script:outgoing}
$script:ns=[pscustomobject]@{Accounts=@([pscustomobject]@{SmtpAddress='sender@example.invalid'})}
$script:ns | Add-Member ScriptMethod GetItemFromID {param($id,$store) return $script:source}
$script:app=[pscustomobject]@{}
$script:app | Add-Member ScriptMethod GetNamespace {param($name) return $script:ns}
function Get-MailResponseOutlook {return $script:app}
$p=[pscustomobject]@{sender_account='sender@example.invalid';to=@('to@example.invalid');cc=@();bcc=@();subject='RE: Original';body='Approved text';mode='reply';source=[pscustomobject]@{entry_id='id';store_id='store';subject='Original';body='Source';modified=$script:source.LastModificationTime.ToUniversalTime().ToString('o')}}
foreach ($mode in @('reply','reply_all','forward')) {
    $p.mode=$mode
    Submit-MailResponse $p | Out-Null
    if ($script:method -ne $mode -or $script:outgoing.Body -ne 'Approved text' -or $script:outgoing.SendUsingAccount.SmtpAddress -ne $p.sender_account) {throw 'Native response contract failed'}
}
if ($script:sends -ne 3) {throw 'Unexpected submission count'}
$p.source.body='Stale'
try {Submit-MailResponse $p;throw 'Expected stale rejection'} catch {if ($_.Exception.Message -notlike '*Source changed*') {throw}}
if ($script:sends -ne 3) {throw 'Stale response sent'}
'Native reply/reply-all/forward mock PASS; no real Outlook used.'
