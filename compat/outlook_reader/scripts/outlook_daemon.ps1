# outlook_daemon.ps1
# OpenClaw Outlook COM daemon. Long-lived PowerShell process that
# holds a single Outlook.Application COM connection and serves JSON
# requests over a named pipe.
#
# IPC: \\.\pipe\OpenClawOutlookDaemon (one client at a time, sequential)
# Wire format: one JSON object per line on the pipe, both directions.
# stderr is reserved for logs only and MUST NOT corrupt the stdout JSON
# line protocol (this script writes to a log file instead).
#
# Lifecycle:
#   * Launched hidden via start_daemon.vbs -> Start-Process.
#   * Stays alive for IDLE_TIMEOUT_SEC of inactivity, or until it
#     receives {"cmd":"shutdown"}.
#   * Outlook is attached lazily on the first "list" / "estimate"
#     request and re-attached automatically if the COM session dies.
#
# Commands (request.cmd):
#   ping        -> {ok, pid, startedAt, uptime, outlookPid, pipe}
#   list        -> {folder, count, hours, mode, max_chars,
#                   attachments, attachments_dir, attachments_dir_win,
#                   before_iso}    returns email array (or {error:...})
#   estimate    -> {folders: [...]} returns {ok, folders, total, generatedAt}
#   detach      -> drops cached Outlook session (will re-attach on next cmd)
#   shutdown    -> graceful exit

param(
    [string]$PipeName = "OpenClawOutlookDaemon",
    # Default 24h. The daemon runs Outlook COM which is slow to attach (~3-5s
    # first call). Re-launching for every sync flashes a brief CMD/PS window
    # before -WindowStyle Hidden kicks in, which is jarring during normal use.
    # Keeping the daemon alive for a full day means it survives across all
    # typical daily syncs with no relaunch, eliminating the flash entirely.
    [int]$IdleTimeoutSec = 86400,
    [string]$LogDir = "",
    [string]$PidFile = ""
)

if ([string]::IsNullOrWhiteSpace($LogDir)) {
    $LogDir = Join-Path $env:LOCALAPPDATA "Temp"
}
if ([string]::IsNullOrWhiteSpace($PidFile)) {
    $PidFile = Join-Path $LogDir "outlook-daemon.pid"
}
$StdoutLog = Join-Path $LogDir "outlook-daemon-stdout.log"
$StderrLog = Join-Path $LogDir "outlook-daemon-stderr.log"

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = (Get-Date).ToString("o")
    "[$ts] [$Level] $Message" | Out-File -FilePath $StderrLog -Append -Encoding UTF8
}

try {
    if (-not (Test-Path -LiteralPath $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }
    Set-Content -Path $PidFile -Value $PID -Encoding ASCII -Force -ErrorAction Stop
} catch {
    Write-Log "Failed to write PID file '$PidFile': $($_.Exception.Message)" "WARN"
}

$startTime = Get-Date
$lastActivity = $startTime
$script:outlook = $null
$script:namespace = $null
$global:SHUTDOWN_REQUESTED = $false

Write-Log "Daemon started. PID=$PID Pipe=$PipeName IdleTimeout=$IdleTimeoutSec PidFile=$PidFile"

# Limit mailbox IPC to the Windows account running this daemon. Without an
# explicit ACL, other local accounts may be able to connect to the fixed pipe
# name and issue read requests.
$pipeSecurity = New-Object System.IO.Pipes.PipeSecurity
$currentUserSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
$pipeRule = New-Object System.IO.Pipes.PipeAccessRule(
    $currentUserSid,
    [System.IO.Pipes.PipeAccessRights]::FullControl,
    [System.Security.AccessControl.AccessControlType]::Allow
)
$pipeSecurity.SetAccessRuleProtection($true, $false)
[void]$pipeSecurity.AddAccessRule($pipeRule)

# --- Outlook attach / detach ---
function Get-OutlookSession {
    # Drop a broken cached handle first.
    if ($null -ne $script:outlook) {
        try {
            $probe = $script:outlook.Application
            if ($null -eq $probe) { throw "probe null" }
        } catch {
            $script:outlook = $null
            $script:namespace = $null
        }
    }
    if ($null -ne $script:outlook) { return $script:outlook }

    $ol = $null
    $err = ""
    try {
        $ol = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Outlook.Application")
    } catch { $err = $_.Exception.Message }
    if (-not $ol) {
        try { $ol = New-Object -ComObject Outlook.Application } catch { $err = $_.Exception.Message }
    }
    if (-not $ol) {
        Write-Log "Outlook attach failed: $err" "WARN"
        return $null
    }
    try {
        $script:outlook = $ol
        $script:namespace = $ol.GetNamespace("MAPI")
    } catch {
        $script:outlook = $null
        $script:namespace = $null
        Write-Log "Outlook namespace failed: $($_.Exception.Message)" "WARN"
        return $null
    }
    return $script:outlook
}

function Detach-OutlookSession {
    $script:outlook = $null
    $script:namespace = $null
    Write-Log "Outlook session detached by request."
}

# --- Folder resolution ---
function Resolve-Folders {
    param([string]$FolderSpec, [object]$Namespace)
    if ($null -eq $Namespace) { return $null }
    $folderKey = ([string]$FolderSpec).ToLowerInvariant()
    $folderMap = @{inbox=6; sent=5; drafts=16; deleted=3; junk=23}
    $scanInboxTree = @("inboxall","inbox-all","inbox_subfolders","inbox+subfolders") -contains $folderKey
    $inboxPathMode = $folderKey.StartsWith("inbox/")

    function Get-FoldersRecursive([object]$root) {
        $result = @()
        if ($null -eq $root) { return $result }
        $result += $root
        foreach ($sf in $root.Folders) {
            $result += Get-FoldersRecursive $sf
        }
        return $result
    }

    if ($scanInboxTree) {
        return ,(Get-FoldersRecursive ($Namespace.GetDefaultFolder(6)))
    } elseif ($inboxPathMode) {
        $cur = $Namespace.GetDefaultFolder(6)
        $parts = ($FolderSpec -split "/") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        if ($parts.Count -gt 0 -and $parts[0].ToLowerInvariant() -eq "inbox") {
            if ($parts.Count -gt 1) { $parts = $parts[1..($parts.Count-1)] } else { $parts = @() }
        }
        foreach ($p in $parts) {
            $next = $null
            foreach ($sf in $cur.Folders) {
                if ([string]$sf.Name -eq [string]$p) { $next = $sf; break }
            }
            if ($null -eq $next) { return $null }
            $cur = $next
        }
        return ,@($cur)
    } else {
        $fid = $folderMap[$folderKey]
        if (-not $fid) { return $null }
        return ,@($Namespace.GetDefaultFolder($fid))
    }
}

. (Join-Path $PSScriptRoot 'mail_paging.ps1')

# --- list handler ---
function Invoke-ListRequest {
    param($Req)
    $countVal = 1
    try { $countVal = [int]$Req.count } catch {}
    if ($countVal -le 0) { $countVal = 1 }
    if ($countVal -gt 500) { $countVal = 500 }
    $folderVal = "Inbox"
    if ($Req.PSObject.Properties.Name -contains "folder") { $folderVal = [string]$Req.folder }
    $hoursVal = 0
    if ($Req.PSObject.Properties.Name -contains "hours") { try { $hoursVal = [int]$Req.hours } catch {} }
    if ($hoursVal -lt 0) {
        return (@{ error = "hours must be non-negative" } | ConvertTo-Json -Compress)
    }
    $modeVal = "all"
    if ($Req.PSObject.Properties.Name -contains "mode") { $modeVal = ([string]$Req.mode).ToLowerInvariant() }
    if (@("all", "unread") -notcontains $modeVal) {
        return (@{ error = "Invalid mode"; valid = @("all", "unread") } | ConvertTo-Json -Compress)
    }
    $maxCharsVal = 4000
    if ($Req.PSObject.Properties.Name -contains "max_chars") { try { $maxCharsVal = [int]$Req.max_chars } catch {} }
    if ($maxCharsVal -le 0) { $maxCharsVal = 4000 }
    if ($maxCharsVal -gt 100000) { $maxCharsVal = 100000 }
    $attachmentsMode = "none"
    if ($Req.PSObject.Properties.Name -contains "attachments") { $attachmentsMode = ([string]$Req.attachments).ToLowerInvariant() }
    if (@("none", "meta", "export") -notcontains $attachmentsMode) {
        return (@{ error = "Invalid attachments mode"; valid = @("none", "meta", "export") } | ConvertTo-Json -Compress)
    }
    $attachmentsDirWsl = ""
    if ($Req.PSObject.Properties.Name -contains "attachments_dir") { $attachmentsDirWsl = [string]$Req.attachments_dir }
    $attachmentsDirWin = ""
    if ($Req.PSObject.Properties.Name -contains "attachments_dir_win") { $attachmentsDirWin = [string]$Req.attachments_dir_win }
    $beforeIsoVal = ""
    if ($Req.PSObject.Properties.Name -contains "before_iso") { $beforeIsoVal = [string]$Req.before_iso }

    $ol = Get-OutlookSession
    if (-not $ol) {
        $running = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
        $pidList = ""
        if ($running) { $pidList = (($running | ForEach-Object { $_.Id }) -join ",") }
        return (@{ error = "No Outlook instance available. Please start Outlook first."; running = $pidList } | ConvertTo-Json -Compress)
    }

    $foldersToScan = Resolve-Folders -FolderSpec $folderVal -Namespace $script:namespace
    if ($null -eq $foldersToScan) {
        return (@{ error = "Invalid or missing Outlook folder"; folder = $folderVal } | ConvertTo-Json -Compress)
    }

    $result = [System.Collections.ArrayList]::new()
$page = @(Select-OutlookMailPage -Folders $foldersToScan -Count $countVal -Hours $hoursVal -Mode $modeVal -Before $beforeIsoVal)
foreach ($candidate in $page) {
    $item = $candidate.Item
    $entryId = [string]$item.EntryID
        $bodyRaw = ""
            try { $bodyRaw = [string]$item.Body } catch { $bodyRaw = "" }
            if ($null -eq $bodyRaw) { $bodyRaw = "" }
            $bodyTruncated = $false
            $bodyOut = $bodyRaw
            if ($bodyOut.Length -gt $maxCharsVal) {
                $bodyOut = $bodyOut.Substring(0, $maxCharsVal)
                $bodyTruncated = $true
            }

            $senderEmail = ""
            try { $senderEmail = [string]$item.SenderEmailAddress } catch { $senderEmail = "" }
            $toStr = ""
            try { $toStr = [string]$item.To } catch { $toStr = "" }

            $attList = [System.Collections.ArrayList]::new()
            if ($attachmentsMode -ne "none") {
                try {
                    $attachmentIndex = 0
                    foreach ($att in $item.Attachments) {
                        if ($null -eq $att) { continue }
                        $attachmentIndex++
                        $fname = [string]$att.FileName
                        if ([string]::IsNullOrWhiteSpace($fname)) { $fname = "attachment.bin" }
                        $safeName = [System.Text.RegularExpressions.Regex]::Replace($fname, '[\\/:*?"<>|]', '_')
                        $safeName = ("{0:D3}_{1}" -f $attachmentIndex, $safeName)
                        $attObj = @{
                            FileName = $fname
                            SafeFileName = $safeName
                            Size = 0
                            Type = ""
                            ExportedPath = ""
                        }
                        try { $attObj.Size = [int]$att.Size } catch {}
                        try { $attObj.Type = [string]$att.Type } catch {}
                        if ($attachmentsMode -eq "export") {
                            try {
                                $entrySafe = [System.Text.RegularExpressions.Regex]::Replace([string]$item.EntryID, '[^A-Za-z0-9]', '_')
                                $subWsl = Join-Path $attachmentsDirWsl $entrySafe
                                $subWin = Join-Path $attachmentsDirWin $entrySafe
                                if (-not (Test-Path -LiteralPath $subWin)) {
                                    New-Item -ItemType Directory -Path $subWin -Force | Out-Null
                                }
                                $targetWin = Join-Path $subWin $safeName
                                $att.SaveAsFile($targetWin)
                                $attObj.ExportedPath = (Join-Path $subWsl $safeName).Replace('\','/')
                            } catch { $attObj.ExportedPath = "" }
                        }
                        [void]$attList.Add($attObj)
                    }
                } catch { $attList = [System.Collections.ArrayList]::new() }
            }

            $rtOut = ""
            try { $rtOut = $candidate.Time.ToString("o") } catch { $rtOut = "" }
            $isUnread = $false
            try { $isUnread = [bool]$item.UnRead } catch {}

            [void]$result.Add(@{
                EntryID = $entryId
                StoreID = $candidate.StoreID
                Subject = $item.Subject
                Sender = $item.SenderName
                SenderEmailAddress = $senderEmail
                To = $toStr
                ReceivedTime = $rtOut
                Unread = $isUnread
                Body = $bodyOut
                BodyTruncated = $bodyTruncated
                Attachments = @($attList)
            })
        }

    if ($result.Count -eq 0) { "[]" } else { ($result | ConvertTo-Json -Depth 5 -Compress) }
}

# --- estimate handler ---
function Invoke-EstimateRequest {
    param($Req)
    $folders = @()
    if ($Req.PSObject.Properties.Name -contains "folders") { $folders = @($Req.folders) }
    if ($folders.Count -eq 0) { $folders = @("InboxAll", "Sent") }

    $ol = Get-OutlookSession
    if (-not $ol) {
        return (@{ ok = $false; error = "No Outlook instance available." } | ConvertTo-Json -Compress)
    }
    $folderMap = @{inbox=6; sent=5; drafts=16; deleted=3; junk=23}

    $totals = @{}
    foreach ($folderName in $folders) {
        $toScan = Resolve-Folders -FolderSpec ([string]$folderName) -Namespace $script:namespace
        if ($null -eq $toScan) {
            $totals[$folderName] = 0
            continue
        }
        $c = 0
        foreach ($f in $toScan) {
            try { $c += [int]$f.Items.Count } catch {}
        }
        $totals[$folderName] = $c
    }
    $sum = 0
    foreach ($v in $totals.Values) { $sum += [int]$v }
    return (@{
        ok = $true
        folders = $totals
        total = $sum
        generatedAt = (Get-Date).ToUniversalTime().ToString("o")
    } | ConvertTo-Json -Compress)
}

# --- ping handler ---
function Invoke-PingRequest {
    $ol = Get-OutlookSession
    $olPid = 0
    if ($ol) {
        try {
            $proc = (Get-Process | Where-Object { $_.MainWindowTitle -ne "" -or $_.ProcessName -eq "OUTLOOK" } | Where-Object { $_.Id -ne 0 } | Select-Object -First 1)
            # Simpler: scan by process name
            $outlookProcs = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
            if ($outlookProcs) { $olPid = [int]($outlookProcs | Select-Object -First 1).Id }
        } catch {}
    }
    return (@{
        ok = $true
        pid = $PID
        startedAt = $startTime.ToString("o")
        uptime = [int]((Get-Date) - $startTime).TotalSeconds
        idleSec = [int]((Get-Date) - $lastActivity).TotalSeconds
        outlookAttached = ($null -ne $script:outlook)
        outlookPid = $olPid
        pipe = $PipeName
    } | ConvertTo-Json -Compress)
}

# --- main pipe loop ---
try {
    while (-not $global:SHUTDOWN_REQUESTED) {
        $idle = ((Get-Date) - $lastActivity).TotalSeconds
        if ($idle -gt $IdleTimeoutSec) {
            Write-Log "Idle timeout reached ($([int]$idle)s). Exiting."
            break
        }

        # Named pipe race after Dispose: Windows may not release the pipe
        # name immediately when we try to rebind in a tight loop after a
        # client disconnects. A single attempt that throws
        # "All pipe instances are busy" used to be treated as fatal and
        # killed the daemon 2-7s after start. Retry with backoff instead.
        $pipe = $null
        $pipeCreateAttempts = 0
        $pipeCreateMaxAttempts = 20
        $pipeCreateDelayMs = 250
        while (-not $pipe -and $pipeCreateAttempts -lt $pipeCreateMaxAttempts -and -not $global:SHUTDOWN_REQUESTED) {
            $pipeCreateAttempts++
            try {
                $pipe = New-Object System.IO.Pipes.NamedPipeServerStream(
                    $PipeName,
                    [System.IO.Pipes.PipeDirection]::InOut,
                    1,
                    [System.IO.Pipes.PipeTransmissionMode]::Byte,
                    [System.IO.Pipes.PipeOptions]::None,
                    4096,
                    4096,
                    $pipeSecurity
                )
            } catch {
                if ($pipeCreateAttempts -ge $pipeCreateMaxAttempts) {
                    Write-Log "Failed to create pipe after $pipeCreateAttempts attempts: $($_.Exception.Message)" "ERROR"
                    break
                }
                Write-Log "Pipe create attempt $pipeCreateAttempts/$pipeCreateMaxAttempts busy ($($_.Exception.Message)); backing off ${pipeCreateDelayMs}ms" "WARN"
                Start-Sleep -Milliseconds $pipeCreateDelayMs
            }
        }
        if (-not $pipe) {
            # break out of the outer while loop (only if we exhausted retries)
            break
        }

        try {
            $pipe.WaitForConnection()
        } catch {
            Write-Log "WaitForConnection failed: $($_.Exception.Message)" "WARN"
            try { $pipe.Dispose() } catch {}
            continue
        }

        $lastActivity = Get-Date
        # Per-connection reader/writer setup. The client may have connected
        # and then immediately disconnected (e.g. parent python exited
        # between WaitForConnection returning and the next syscall). Any
        # exception here must NOT kill the daemon -- just discard the
        # broken stream and accept a new connection.
        $reader = $null
        $writer = $null
        try {
            $reader = New-Object System.IO.StreamReader($pipe, [System.Text.Encoding]::UTF8, $false, 1024, $true)
            $writer = New-Object System.IO.StreamWriter($pipe, [System.Text.Encoding]::UTF8, 1024, $true)
            # NB: do NOT set $writer.AutoFlush = $true. The setter calls
            # Flush() on the underlying stream; if the pipe is already
            # broken (client disconnected), Flush() throws
            # "Pipe is broken" and the daemon crashes. We instead call
            # Flush() explicitly after each WriteLine, wrapped in try/catch.
        } catch {
            Write-Log "Per-connection stream init failed (broken pipe?): $($_.Exception.Message)" "WARN"
            try { $reader.Dispose() } catch {}
            try { $writer.Dispose() } catch {}
            try { $pipe.Disconnect() } catch {}
            try { $pipe.Dispose() } catch {}
            continue
        }

        try {
            $connActive = $true
            while ($connActive) {
                $line = $null
                try { $line = $reader.ReadLine() } catch { $connActive = $false; break }
                if ($null -eq $line) { $connActive = $false; break }
                $line = $line.Trim()
                if ([string]::IsNullOrEmpty($line)) { continue }

                $req = $null
                try { $req = $line | ConvertFrom-Json -ErrorAction Stop }
                catch {
                    $errResp = (@{ error = "Bad JSON: $($_.Exception.Message)" } | ConvertTo-Json -Compress)
                    try { $writer.WriteLine($errResp); $writer.Flush() } catch { $connActive = $false }
                    continue
                }

                $cmdName = ""
                if ($req.PSObject.Properties.Name -contains "cmd") { $cmdName = [string]$req.cmd }
                $lastActivity = Get-Date
                $response = $null
                switch ($cmdName.ToLowerInvariant()) {
                    "ping"     { try { $response = Invoke-PingRequest } catch { $response = (@{ error = "ping failed: $($_.Exception.Message)" } | ConvertTo-Json -Compress) } }
                    "list"     { try { $response = Invoke-ListRequest -Req $req } catch { $response = (@{ error = "list failed: $($_.Exception.Message)" } | ConvertTo-Json -Compress) } }
                    "estimate" { try { $response = Invoke-EstimateRequest -Req $req } catch { $response = (@{ error = "estimate failed: $($_.Exception.Message)" } | ConvertTo-Json -Compress) } }
                    "detach"   { Detach-OutlookSession; $response = (@{ ok = $true; msg = "outlook session detached" } | ConvertTo-Json -Compress) }
                    "shutdown" {
                        $response = (@{ ok = $true; msg = "shutting down" } | ConvertTo-Json -Compress)
                        try { $writer.WriteLine($response); $writer.Flush() } catch {}
                        $global:SHUTDOWN_REQUESTED = $true
                        $connActive = $false
                    }
                    default    { $response = (@{ error = "unknown command: $cmdName" } | ConvertTo-Json -Compress) }
                }

                if ($null -ne $response -and -not $global:SHUTDOWN_REQUESTED) {
                    try { $writer.WriteLine($response); $writer.Flush() } catch { $connActive = $false; break }
                }
            }
        } catch {
            Write-Log "Connection handler error: $($_.Exception.Message)" "WARN"
        } finally {
            try { $reader.Dispose() } catch {}
            try { $writer.Dispose() } catch {}
            try { $pipe.Disconnect() } catch {}
            try { $pipe.Dispose() } catch {}
        }
    }
} catch {
    Write-Log "Fatal daemon error: $($_.Exception.Message)" "ERROR"
} finally {
    try { Remove-Item -Path $PidFile -Force -ErrorAction SilentlyContinue } catch {}
    Write-Log "Daemon exiting. PID=$PID"
}
