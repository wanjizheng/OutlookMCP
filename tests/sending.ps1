$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../compat/outlook_reader/scripts/outlook_send_library.ps1')
class SendTestRecipients {
    [System.Collections.ArrayList]$Added = [System.Collections.ArrayList]::new()
    [bool]$Mismatch = $false
    [object] Add([string]$address) {
        $resolved = if ($this.Mismatch) { 'unexpected@example.invalid' } else { $address }
        $recipient = [pscustomobject]@{Type=0;AddressEntry=[pscustomobject]@{Address=$resolved;Type='SMTP'}}
        $recipient | Add-Member ScriptMethod Resolve { return $true }
        [void]$this.Added.Add($recipient)
        return $recipient
    }
}
class SendTestMail {
    [string]$To; [string]$CC; [string]$BCC; [string]$Subject; [string]$Body
    [object]$SendUsingAccount
    [SendTestRecipients]$Recipients = [SendTestRecipients]::new()
    [int]$Sends=0; [int]$Saves=0; [int]$Closes=0
    [void] Send() { $this.Sends++ }
    [void] Save() { $this.Saves++ }
    [void] Close([int]$mode) { if ($mode -ne 1) { throw 'Expected discard' }; $this.Closes++ }
}
function New-SendFixture {
    $account=[pscustomobject]@{SmtpAddress='sender@example.invalid'}
    $accounts=[pscustomobject]@{Count=1;Account=$account}
    $accounts | Add-Member ScriptMethod Item { param($i); return $this.Account }
    $ol=[pscustomobject]@{Session=[pscustomobject]@{Accounts=$accounts};Mail=[SendTestMail]::new()}
    $ol | Add-Member ScriptMethod CreateItem { param($kind); if($kind -ne 0){throw 'Wrong item'}; return $this.Mail }
    return $ol
}
$fields=@{To='to@example.invalid';CC='cc@example.invalid';BCC='bcc@example.invalid';Subject='Exact subject';Body="Exact body`r`nSecond line";SenderAccount='sender@example.invalid'}
$ol=New-SendFixture
$result=Invoke-OutlookComSend -Outlook $ol @fields -DoSend $true -DoDraft $false
if ($result.status -ne 'sent' -or $ol.Mail.Sends -ne 1 -or $ol.Mail.Saves -ne 0) {throw 'Submission mismatch'}
if ($ol.Mail.SendUsingAccount.SmtpAddress -ne $fields.SenderAccount -or $ol.Mail.Body -cne $fields.Body -or $ol.Mail.Subject -cne $fields.Subject) {throw 'Preview changed'}
if (($ol.Mail.Recipients.Added.Type -join ',') -ne '1,2,3') {throw 'To/CC/BCC types wrong'}
$ol=New-SendFixture; $ol.Mail.Recipients.Mismatch=$true
try { Invoke-OutlookComSend -Outlook $ol @fields -DoSend $true; throw 'Mismatch was not rejected' }
catch { if($_.Exception.Message -notlike '*differs from approved*'){throw} }
if ($ol.Mail.Sends -ne 0 -or $ol.Mail.Saves -ne 0 -or $ol.Mail.Closes -ne 1) {throw 'Mismatch wrote mail'}
$ol=New-SendFixture
try { Invoke-OutlookComSend -Outlook $ol -To 'x@example.invalid' -Subject x -Body x -SenderAccount 'missing@example.invalid' -DoSend $true; throw 'Missing sender accepted' }
catch { if($_.Exception.Message -notlike '*not found*'){throw} }
if ($ol.Mail.Sends -ne 0) {throw 'Wrong sender sent'}
$ol=New-SendFixture
$result=Invoke-OutlookComSend -Outlook $ol -To 'Legacy Display Name' -Subject x -Body y
if ($result.status -ne 'draft' -or $ol.Mail.Saves -ne 1 -or $ol.Mail.Sends -ne 0) {throw 'Legacy draft default changed'}
$ol=New-SendFixture
$result=Invoke-OutlookComSend -Outlook $ol -To 'Legacy Display Name' -Subject x -Body y -DoDraft $false
if ($result.status -ne 'preview' -or $ol.Mail.Saves -ne 0 -or $ol.Mail.Sends -ne 0) {throw 'Legacy preview wrote mail'}
'Existing PowerShell sender PASS: exact fields/account/recipients, mismatch rejection, legacy draft/preview defaults. No real COM used.'
