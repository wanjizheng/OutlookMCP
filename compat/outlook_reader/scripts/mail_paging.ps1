# Shared by classic OpenClaw wrappers, daemon, and Windows-native MCP.
# Cursor v2 is a stable (UTC timestamp, store ID + entry ID) exclusive bound.
function Select-OutlookMailPage {
    param($Folders, [int]$Count, [int]$Hours = 0, [string]$Mode = 'all', [string]$Before = '')
    $boundTime = $null; $boundKey = $null
    if ($Before) {
        if ($Before.StartsWith('v2:')) {
            $cursor = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Before.Substring(3))) | ConvertFrom-Json
            $boundTime = [DateTimeOffset]::Parse($cursor.t).UtcDateTime
            $boundKey = [string]$cursor.k
        } else { $boundTime = [DateTimeOffset]::Parse($Before).UtcDateTime }
    }
    $since = $null
    if ($Hours -gt 0) { $since = [DateTime]::UtcNow.AddHours(-$Hours) }
    $candidates = [Collections.Generic.List[object]]::new()
    foreach ($folder in $Folders) {
        $items = $folder.Items
        # Restrict uses minute precision on classic Outlook. Use an inclusive
        # coarse upper window and enforce the exact bound below, never round down.
        $parts = @(); $dq = [char]34
        if ($boundTime) {
            $upper = $boundTime.ToLocalTime().AddMinutes(1).ToString('g', [Globalization.CultureInfo]::CurrentCulture)
            $parts += "[ReceivedTime] < $dq$upper$dq"
        }
        if ($since) {
            $lower = $since.ToLocalTime().ToString('g', [Globalization.CultureInfo]::CurrentCulture)
            $parts += "[ReceivedTime] >= $dq$lower$dq"
        }
        if ($Mode -eq 'unread') { $parts += '[UnRead] = True' }
        if ($parts.Count) { try { $items = $items.Restrict(($parts -join ' AND ')) } catch {} }
        $sorted = $true
        try { $items.Sort('[ReceivedTime]', $true) } catch { $sorted = $false }
        $eligible = 0; $tailTime = $null
        for ($i = 1; $i -le $items.Count; $i++) {
            $item = $items.Item($i)
            try {
                $id = [string]$item.EntryID
                if (-not $id) { continue }
                $time = ([datetime]$item.ReceivedTime).ToUniversalTime()
                $key = ([string]$folder.StoreID + '|' + $id).ToLowerInvariant()
                if ($boundTime -and ($time -gt $boundTime -or
                    ($time -eq $boundTime -and ($null -eq $boundKey -or
                    [string]::CompareOrdinal($key, $boundKey) -ge 0)))) { continue }
                if ($since -and $time -lt $since) { if ($sorted) { break }; continue }
                if ($Mode -eq 'unread' -and -not $item.UnRead) { continue }
                # Include every tie at the local page boundary before sorting.
                if ($sorted -and $eligible -ge $Count -and $time -lt $tailTime) { break }
                $candidates.Add(@{ Item=$item; Time=$time; Key=$key; StoreID=[string]$folder.StoreID })
                $eligible++; $tailTime = $time
            } catch { throw "Unable to read mail paging metadata: $($_.Exception.Message)" }
        }
    }
    $candidates.Sort([Comparison[object]]{
        param($a, $b)
        $cmp = [datetime]::Compare($b.Time, $a.Time)
        if ($cmp -eq 0) { $cmp = [string]::CompareOrdinal($b.Key, $a.Key) }
        return $cmp
    })
    $seen = @{}; $emitted = 0
    foreach ($candidate in $candidates) {
        if ($seen.ContainsKey($candidate.Key)) { continue }
        $seen[$candidate.Key] = $true
        if ($emitted -ge $Count) { break }
        $emitted++; $candidate
    }
}
