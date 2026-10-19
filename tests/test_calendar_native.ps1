$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '..\scripts\calendar_actions.ps1')
$env:OUTLOOK_MCP_ENABLE_SEND='1'
$script:sends=0
$script:added=@()
$script:recipients=[pscustomobject]@{Count=0}
$script:recipients | Add-Member ScriptMethod Add { param($address) $script:added += $address; return [pscustomobject]@{Type=0} }
$script:recipients | Add-Member ScriptMethod ResolveAll { return $true }
$script:item=[pscustomobject]@{SendUsingAccount=$null;MeetingStatus=0;Subject='';Body='';Location='';ForceUpdateToAllAttendees=$false;StartUTC=[datetime]::MinValue;EndUTC=[datetime]::MinValue;Recipients=$script:recipients;EntryID='synthetic'}
$script:item | Add-Member ScriptMethod Send { $script:sends++ }
$items=[pscustomobject]@{}
$items | Add-Member ScriptMethod Add { param($type) if ($type -ne 'IPM.Appointment') {throw 'Wrong type'};return $script:item }
$script:folder=[pscustomobject]@{Items=$items}
$store=[pscustomobject]@{StoreID='store'}
$store | Add-Member ScriptMethod GetDefaultFolder { param($id) if ($id -ne 9) {throw 'Wrong folder'};return $script:folder }
$script:ns=[pscustomobject]@{Accounts=@([pscustomobject]@{SmtpAddress='sender@example.invalid';DeliveryStore=$store})}
$script:app=[pscustomobject]@{}
$script:app | Add-Member ScriptMethod GetNamespace { param($name) return $script:ns }
function Get-ActionOutlook { return $script:app }
$p=[pscustomobject]@{action='create';sender_account='sender@example.invalid';subject='Synthetic';body='Full body';location='Room';start='2026-10-20T14:00:00+01:00';end='2026-10-20T15:00:00+01:00';required=@('to@example.invalid');optional=@('optional@example.invalid')}
$r=Run ([pscustomobject]@{operation='commit';payload=$p})
if ($script:sends -ne 1 -or $script:item.Subject -ne 'Synthetic' -or $script:item.StartUTC.Hour -ne 13 -or $script:added.Count -ne 2) { throw 'Create field/time/send contract failed' }
$p.sender_account='missing@example.invalid'
try { Run ([pscustomobject]@{operation='commit';payload=$p});throw 'Expected account rejection' } catch { if ($_.Exception.Message -notlike '*Sender account unavailable*') {throw} }
if ($script:sends -ne 1) { throw 'Rejected account sent something' }
# Recurrence scope never falls back to editing the whole series.
$script:occurrence=[pscustomobject]@{Class=26;IsRecurring=$true;RecurrenceState=2;StartUTC=[DateTimeOffset]::Parse('2026-10-20T13:00:00Z').UtcDateTime}
$pattern=[pscustomobject]@{}
$pattern | Add-Member ScriptMethod GetOccurrence {param($date) return $script:occurrence}
$script:master=[pscustomobject]@{Class=26;IsRecurring=$true;RecurrenceState=1;Pattern=$pattern}
$script:master | Add-Member ScriptMethod GetRecurrencePattern {return $this.Pattern}
$script:ns | Add-Member ScriptMethod GetItemFromID {param($id,$store) return $script:master}
$scopeRequest=[pscustomobject]@{entry_id='id';store_id='store';scope='single';occurrence_start='2026-10-20T14:00:00+01:00'}
try { Event $script:ns $scopeRequest;throw 'Expected scope rejection' } catch {if ($_.Exception.Message -notlike '*requires occurrence or series*') {throw}}
$scopeRequest.scope='occurrence'
if ((Event $script:ns $scopeRequest).RecurrenceState -ne 2) {throw 'Wrong occurrence'}
$scopeRequest.occurrence_start='2026-10-21T14:00:00+01:00'
try { Event $script:ns $scopeRequest;throw 'Expected occurrence mismatch' } catch {if ($_.Exception.Message -notlike '*Occurrence mismatch*') {throw}}
$scopeRequest.scope='series'
if ((Event $script:ns $scopeRequest).RecurrenceState -ne 1) {throw 'Wrong series'}
# Already expanded occurrence/exception must not call GetRecurrencePattern again.
$script:master=$script:occurrence
$scopeRequest.scope='occurrence';$scopeRequest.occurrence_start='2026-10-20T14:00:00+01:00'
if ((Event $script:ns $scopeRequest).RecurrenceState -ne 2) {throw 'Expanded occurrence failed'}
$script:master.RecurrenceState=3
if ((Event $script:ns $scopeRequest).RecurrenceState -ne 3) {throw 'Expanded exception failed'}
$scopeRequest.scope='series'
try {Event $script:ns $scopeRequest;throw 'Expected exception series rejection'} catch {if ($_.Exception.Message -notlike '*requires master EntryID*') {throw}}
# Stale comparison must occur before any mutation/send.
function Event($ns,$p) { return $script:item }
function Snapshot($item) { return [ordered]@{subject='Changed';organizer_smtp='sender@example.invalid'} }
$p | Add-Member NoteProperty store_id 'store'
$p.action='update';$p.sender_account='sender@example.invalid';$p | Add-Member NoteProperty before ([pscustomobject]@{subject='Original';organizer_smtp='sender@example.invalid'})
try { Run ([pscustomobject]@{operation='commit';payload=$p});throw 'Expected stale rejection' } catch { if ($_.Exception.Message -notlike '*Event changed since preview*') {throw} }
if ($script:sends -ne 1 -or $script:item.Subject -ne 'Synthetic') {throw 'Stale operation mutated item'}
'Calendar native mock PASS: sender, recipients, UTC conversion, single send, stale rejection; no COM used.'

function Snapshot($item) { return [ordered]@{subject='Original';organizer_smtp='sender@example.invalid'} }
$p.subject='Updated'
Run ([pscustomobject]@{operation='commit';payload=$p}) | Out-Null
if ($script:sends -ne 2 -or $script:item.Subject -ne 'Updated') {throw 'Update did not preserve approved fields'}
$p.action='cancel'
Run ([pscustomobject]@{operation='commit';payload=$p}) | Out-Null
if ($script:sends -ne 3 -or $script:item.MeetingStatus -ne 5) {throw 'Cancellation contract failed'}

$script:responseCode=0
$script:response=[pscustomobject]@{SendUsingAccount=$null}
$script:response | Add-Member ScriptMethod Send {$script:sends++}
$script:item | Add-Member ScriptMethod Respond {param($code,$noUI,$dialog) if (-not $noUI -or $dialog) {throw 'Unexpected response UI'};$script:responseCode=$code;return $script:response}
function Snapshot($item) {return [ordered]@{subject='Original';organizer_smtp='organizer@example.invalid'}}
$p.before=[pscustomobject]@{subject='Original';organizer_smtp='organizer@example.invalid'}
$script:item.MeetingStatus=3
foreach ($action in @('accept','tentative','decline')) {
    $p.action=$action
    Run ([pscustomobject]@{operation='commit';payload=$p}) | Out-Null
    $expected=@{accept=3;tentative=2;decline=4}
    if ($script:responseCode -ne $expected[$action] -or $script:response.SendUsingAccount.SmtpAddress -ne 'sender@example.invalid') {throw 'Response contract failed'}
}
if ($script:sends -ne 6) {throw 'Responses were not submitted exactly once each'}
