param([int]$Top = 10, [string]$Folder = "Inbox", [int]$Hours = 0)

$ol = $null
$activeErr = ""
$newErr = ""
try { $ol = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Outlook.Application") } catch { $activeErr = $_.Exception.Message }
if (-not $ol) {
    try { $ol = New-Object -ComObject Outlook.Application } catch { $newErr = $_.Exception.Message }
}
if (-not $ol) {
    $running = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    $detail = ""
    if ($newErr) { $detail = $newErr } elseif ($activeErr) { $detail = $activeErr }

    if ($running) {
        $pidList = (($running | ForEach-Object { $_.Id }) -join ",")
        Write-Output (@{
            error = "Outlook process detected (PID $pidList), but COM connection still failed after retry. Ensure Outlook and this script run at the same privilege level (both admin or both normal user). Try restarting Outlook normally (not as admin)."
            detail = $detail
        } | ConvertTo-Json)
    } else {
        Write-Output (@{
            error = "No Outlook instance available. Please start Outlook first."
            detail = $detail
        } | ConvertTo-Json)
    }
    exit 1
}

$ns = $ol.GetNamespace("MAPI")
$folderId = @{Inbox=6; Sent=5; Drafts=16; Deleted=3; Junk=23}[$Folder]
if (-not $folderId) { $folderId = 6 }

$folder = $ns.GetDefaultFolder($folderId)
$items = $folder.Items

# Normalize ordering so index 1 is always most recent mail.
try { $items.Sort("[ReceivedTime]", $true) } catch {}

$result = @()

if ($Hours -gt 0) {
    $limitDate = (Get-Date).AddHours(-$Hours)
    $count = 0
    for ($i = 1; $i -le $items.Count -and $count -lt $Top; $i++) {
        $item = $items.Item($i)
        if ($null -eq $item) { continue }
        if ($item.ReceivedTime -ge $limitDate) {
            $toAddrs = @()
            try {
                foreach ($r in $item.Recipients) {
                    if ($null -ne $r) {
                        $addr = $r.Address
                        if ($addr -match '^/o=') {
                            try { $addr = $r.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E') } catch {}
                        }
                        $toAddrs += $addr
                    }
                }
            } catch { $toAddrs = @() }

            $result += @{
                Subject = $item.Subject
                Sender = $item.SenderName
                To = @($toAddrs)
                ReceivedTime = $item.ReceivedTime.ToString("o")
            }
            $count++
        }
    }
} else {
    # No time filter - just get the most recent items
    for ($i = 1; $i -le $items.Count -and $result.Count -lt $Top; $i++) {
        $item = $items.Item($i)
        if ($null -eq $item) { continue }
        $toAddrs = @()
        try {
            foreach ($r in $item.Recipients) {
                if ($null -ne $r) {
                    $addr = $r.Address
                    if ($addr -match '^/o=') {
                        try { $addr = $r.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E') } catch {}
                    }
                    $toAddrs += $addr
                }
            }
        } catch { $toAddrs = @() }

        $result += @{
            Subject = $item.Subject
            Sender = $item.SenderName
            To = @($toAddrs)
            ReceivedTime = $item.ReceivedTime.ToString("o")
        }
    }
}

$result | ConvertTo-Json -Depth 2
