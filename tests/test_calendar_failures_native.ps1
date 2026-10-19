$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'test_calendar_native.ps1') | Out-Null
function Snapshot($item) {return [ordered]@{subject='Original';organizer_smtp='sender@example.invalid'}}
$p.before=[pscustomobject]@{subject='Original';organizer_smtp='sender@example.invalid'}
$p.action='update';$p.subject='Approved update'
$script:item.Saved=$true
$script:closes=0
$script:item | Add-Member ScriptMethod Close {param($mode) if($mode -ne 1){throw 'Not discard'};$script:closes++}
# Any reassignment of an existing item's inherited account must be avoided.
$script:inheritedAccount=$script:item.SendUsingAccount
$script:item.PSObject.Properties.Remove('SendUsingAccount')
$script:item | Add-Member ScriptProperty SendUsingAccount {$script:inheritedAccount} {throw 'Forbidden existing account reassignment'}
$script:forceFailure=$true
$script:item.PSObject.Properties.Remove('ForceUpdateToAllAttendees')
$script:item | Add-Member ScriptProperty ForceUpdateToAllAttendees {$false} {
    if ($script:forceFailure) {throw [Runtime.InteropServices.COMException]::new("You don't have appropriate permission to perform this operation.",-2147024891)}
}
$sentBefore=$script:sends
try {Run ([pscustomobject]@{operation='commit';payload=$p});throw 'Expected distribution failure'} catch {
    if ($_.Exception.Message -eq 'Expected distribution failure') {throw}
    $failure=Get-ActionFailure $_ 'commit'
}
if ($failure.phase -ne 'set_update_distribution' -or -not $failure.mutation_attempted -or $failure.send_attempted -or $failure.hresult -ne '0x80070005' -or $failure.discard -ne 'close_discard_returned' -or $script:closes -ne 1 -or $script:sends -ne $sentBefore) {throw 'Pre-send failure evidence contract failed'}
# Failure during Send is uncertain: never discard and never retry.
$script:forceFailure=$false
$script:item | Add-Member ScriptMethod Send {throw [Runtime.InteropServices.COMException]::new('Send denied',-2147024891)} -Force
try {Run ([pscustomobject]@{operation='commit';payload=$p});throw 'Expected send failure'} catch {
    if ($_.Exception.Message -eq 'Expected send failure') {throw}
    $failure=Get-ActionFailure $_ 'commit'
}
if ($failure.phase -ne 'send_invitation' -or -not $failure.send_attempted -or $failure.send_returned -or $failure.status -ne 'send_outcome_unknown' -or $script:closes -ne 1) {throw 'Send failure evidence contract failed'}
# Unsaved/open state is rejected before any operation-owned mutation.
$script:item.Saved=$false
try {Run ([pscustomobject]@{operation='commit';payload=$p});throw 'Expected unsaved rejection'} catch {
    if ($_.Exception.Message -eq 'Expected unsaved rejection') {throw}
    $failure=Get-ActionFailure $_ 'commit'
}
if ($failure.phase -ne 'validate_editable_state' -or $failure.mutation_attempted -or $failure.send_attempted -or $script:closes -ne 1) {throw 'Unsaved-state guard failed'}
'Failure phases/HRESULT/discard/uncertain-send contract PASS; no real Outlook calls.'

$script:item.Saved=$true
$script:accessor | Add-Member ScriptMethod GetProperty {param($name) return 2} -Force
try {Run ([pscustomobject]@{operation='commit';payload=$p});throw 'Expected modify permission rejection'} catch {
    if ($_.Exception.Message -eq 'Expected modify permission rejection') {throw}
    $failure=Get-ActionFailure $_ 'commit'
}
if ($failure.phase -ne 'validate_write_access' -or $failure.mutation_attempted -or $failure.send_attempted -or $script:closes -ne 1) {throw 'Read-only event was modified'}
