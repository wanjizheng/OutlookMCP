$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '../compat/outlook_reader/scripts/mail_paging.ps1')
class TestItems {
    [object[]]$Rows
    [int]$Count
    TestItems([object[]]$rows) { $this.Rows=$rows; $this.Count=$rows.Count }
    [void] Sort([string]$field,[bool]$desc) { }
    [object] Restrict([string]$filter) { throw 'Simulated provider filter failure' }
    [object] Item([int]$i) { return $this.Rows[$i-1] }
}
$t = [datetime]'2026-10-17T10:00:31'
function Mail($id,$time) { return [pscustomobject]@{EntryID=$id;ReceivedTime=$time;UnRead=$true} }
$folders = @(
    [pscustomobject]@{StoreID='s';Items=[TestItems]::new(@((Mail 'a' $t),(Mail 'b' $t),(Mail 'c' $t),(Mail 'd' $t),(Mail 'old' $t.AddMinutes(-1))))},
    [pscustomobject]@{StoreID='s';Items=[TestItems]::new(@((Mail 'new' $t.AddSeconds(10)),(Mail 'z' $t)))}
)
$all=@(); $cursor=''
for($i=0;$i -lt 10;$i++) {
    $page=@(Select-OutlookMailPage $folders 2 0 all $cursor)
    if(-not $page.Count){break}
    $all+=@($page | ForEach-Object {$_.Item.EntryID})
    $last=$page[-1]
    $json=@{t=$last.Time.ToString('o');k=$last.Key}|ConvertTo-Json -Compress
    $cursor='v2:'+ [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
}
if(($all -join ',') -ne 'new,z,d,c,b,a,old'){throw "Wrong coverage/order: $all"}
$before=$t.ToUniversalTime().AddSeconds(1).ToString('o')
$page=@(Select-OutlookMailPage $folders 10 0 all $before)
if($page.Count -ne 6){throw 'Legacy ISO seconds boundary lost messages'}
'Paging behavior PASS: same-second ties, subfolders, legacy boundary, Restrict failure'
