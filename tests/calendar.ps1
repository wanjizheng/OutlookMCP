$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../scripts/calendar_library.ps1')
class CalendarItems : System.Collections.ArrayList {
    [bool]$IncludeRecurrences
    [void]Sort([string]$field){}
    [void]Sort([string]$field,[bool]$desc){}
    [object]Restrict([string]$filter){
        # A fake provider interprets the actual requested interval; it does not
        # preselect expected rows for the implementation under test.
        if($filter -notmatch '^\[End\] > "([^"]+)" AND \[Start\] < "([^"]+)"$'){throw "Wrong overlap filter: $filter"}
        $start=[datetime]::Parse($matches[1]);$end=[datetime]::Parse($matches[2])
        return @($this.ToArray()|Where-Object {$_.End -gt $start -and $_.Start -lt $end})
    }
}
function Event($id,$start,$end){return [pscustomobject]@{EntryID=$id;Subject=$id;Start=$start;End=$end;Recipients=@();Organizer='';Body='';Location='';MeetingStatus=0;ResponseStatus=0;ResponseRequested=$false}}
$script:events=[CalendarItems]::new()
$today=[datetime]::Today
[void]$script:events.Add((Event 'ongoing' $today.AddDays(-1) $today.AddHours(1)))
[void]$script:events.Add((Event 'ended-boundary' $today.AddDays(-1) $today))
[void]$script:events.Add((Event 'next-window' $today.AddDays(1) $today.AddDays(1).AddHours(1)))
$script:empty=[CalendarItems]::new()
$script:ns=[pscustomobject]@{}
$script:ns|Add-Member ScriptMethod GetDefaultFolder {param($id) if($id -eq 9){return [pscustomobject]@{Items=$script:events}};return [pscustomobject]@{Items=$script:empty}}
$ol=[pscustomobject]@{}
$ol|Add-Member ScriptMethod GetNamespace {param($name) return $script:ns}
$out=Read-OutlookCalendar $ol 1 20 0 | ConvertFrom-Json
if(@($out).Count -ne 1 -or $out.EntryID -ne 'ongoing'){throw 'Ongoing/boundary calendar coverage failed'}
if(-not $script:events.IncludeRecurrences){throw 'Recurring expansion was disabled'}
'Calendar overlap and recurrence setup PASS'
