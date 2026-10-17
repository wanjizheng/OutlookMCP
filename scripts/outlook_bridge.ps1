$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'mail_library.ps1')
. (Join-Path $PSScriptRoot 'calendar_library.ps1')
function Write-Log { param($Message,$Level) [Console]::Error.WriteLine($Message) }
try {
    $req = [Console]::In.ReadToEnd() | ConvertFrom-Json
    switch ([string]$req.operation) {
        'list' {
            # Do not accept export paths or attachment-export mode from MCP.
            $safe = [pscustomobject]@{count=$req.count;folder=$req.folder;hours=$req.hours;mode=$req.mode;max_chars=$req.max_chars;before_iso=$req.cursor;attachments='meta'}
            Invoke-ListRequest $safe
        }
        'calendar' {
            $ol = Get-OutlookSession
            if (-not $ol) { throw 'Classic Outlook unavailable. Start it under this Windows user.' }
            Read-OutlookCalendar $ol $req.days $req.top $req.offset
        }
        'read' {
            $ol = Get-OutlookSession
            if (-not $ol) { throw 'Classic Outlook unavailable' }
            if ($req.store_id) { $item=$script:namespace.GetItemFromID([string]$req.entry_id,[string]$req.store_id) }
            else { $item=$script:namespace.GetItemFromID([string]$req.entry_id) }
            $body=[string]$item.Body
            @{EntryID=$item.EntryID;Subject=$item.Subject;Sender=$item.SenderName;ReceivedTime=$item.ReceivedTime.ToString('o');Body=$body.Substring(0,[Math]::Min($body.Length,[int]$req.max_chars));BodyTruncated=($body.Length -gt $req.max_chars)} | ConvertTo-Json -Depth 5 -Compress
        }
        'status' {
            @{windows=$true;outlookRunning=([bool](Get-Process OUTLOOK -ErrorAction SilentlyContinue));wslRequired=$false} | ConvertTo-Json -Compress
        }
        default { throw 'Operation is not allowed by the read-only bridge' }
    }
} catch {
    @{error=$_.Exception.Message} | ConvertTo-Json -Compress
    exit 1
}
