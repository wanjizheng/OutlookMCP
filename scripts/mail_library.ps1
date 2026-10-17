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

. (Join-Path $PSScriptRoot '../compat/outlook_reader/scripts/mail_paging.ps1')

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
