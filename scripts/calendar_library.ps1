function Read-OutlookCalendar {
param($ol,[int]$DaysVal,[int]$TopVal,[int]$StartOffsetVal)
$ns = $ol.GetNamespace("MAPI")
$cal = $ns.GetDefaultFolder(9)
$items = $cal.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true

# 计算查询窗口：StartOffset指定窗口起点（相对今天），Days指定窗口大小
# 使用 .Date 保证按自然日窗口查询，不会漏掉白天的会议
$now = (Get-Date).Date
$startDate = $now.AddDays($StartOffsetVal)
$endDate = $startDate.AddDays($DaysVal)

# 构造MAPI过滤器 (en-GB: d/M/yyyy HH:mm, double-quoted)
$dq = [char]34
$filter = "[End] > $dq" + $startDate.ToString("g", [Globalization.CultureInfo]::CurrentCulture) + "$dq AND [Start] < $dq" + $endDate.ToString("g", [Globalization.CultureInfo]::CurrentCulture) + "$dq"
$restricted = $items.Restrict($filter)

function Get-ResponseStatusLabel([int]$statusCode) {
    switch ($statusCode) {
        0 { "none" }
        1 { "organized" }
        2 { "tentative" }
        3 { "accepted" }
        4 { "declined" }
        5 { "not_responded" }
        default { "unknown" }
    }
}

function Get-RecipientTypeLabel([int]$recipientType) {
    switch ($recipientType) {
        1 { "required" }
        2 { "optional" }
        3 { "resource" }
        default { "unknown" }
    }
}

function Get-AttendeesSnapshot([object]$calendarItem) {
    $attendees = [System.Collections.ArrayList]::new()
    $summary = @{
        total = 0
        accepted = 0
        tentative = 0
        declined = 0
        notResponded = 0
        organized = 0
        none = 0
        unknown = 0
    }

    try {
        if ($calendarItem.Recipients.Count -gt 0) {
            for ($idx = 1; $idx -le $calendarItem.Recipients.Count; $idx++) {
                $rc = $calendarItem.Recipients.Item($idx)
                $displayName = ""
                $address = ""
                $typeCode = -1
                $responseStatusCode = -1

                try { $displayName = [string]$rc.Name } catch {}
                try { $address = [string]$rc.Address } catch {}
                try { $typeCode = [int]$rc.Type } catch {}
                try { $responseStatusCode = [int]$rc.MeetingResponseStatus } catch {}

                $responseStatus = Get-ResponseStatusLabel $responseStatusCode
                $recipientType = Get-RecipientTypeLabel $typeCode

                [void]$attendees.Add(@{
                    Name = $displayName
                    Address = $address
                    Type = $recipientType
                    TypeCode = $typeCode
                    ResponseStatus = $responseStatus
                    ResponseStatusCode = $responseStatusCode
                })

                switch ($responseStatus) {
                    "accepted" { $summary.accepted++ }
                    "tentative" { $summary.tentative++ }
                    "declined" { $summary.declined++ }
                    "not_responded" { $summary.notResponded++ }
                    "organized" { $summary.organized++ }
                    "none" { $summary.none++ }
                    default { $summary.unknown++ }
                }
            }
        }
    } catch {}

    $summary.total = $attendees.Count
    return @{
        attendees = @($attendees)
        summary = $summary
    }
}

function Normalize-MeetingSubject([string]$subject) {
    if ([string]::IsNullOrWhiteSpace($subject)) { return "" }
    $s = $subject.Trim()
    $s = $s -replace '^(accepted|declined|tentative|proposed new time|accepted tentative|accept|decline|tentative|已接受|已拒绝|已婉拒|已暂定|暂定|接受|拒绝)\s*[:：]\s*', ''
    return $s.Trim()
}

function Get-ResponseStatusFromMessage([string]$messageClass, [string]$subject) {
    $mc = ""
    if ($messageClass) { $mc = $messageClass.ToLowerInvariant() }
    if ($mc -like "*resp.pos*") { return "accepted" }
    if ($mc -like "*resp.tent*") { return "tentative" }
    if ($mc -like "*resp.neg*") { return "declined" }

    if ($subject -match '^(accepted|accept|已接受|接受)\s*[:：]') { return "accepted" }
    if ($subject -match '^(declined|decline|已拒绝|已婉拒|拒绝)\s*[:：]') { return "declined" }
    if ($subject -match '^(tentative|暂定|已暂定)\s*[:：]') { return "tentative" }
    return "unknown"
}

function Get-ResponseStatusCodeFromLabel([string]$label) {
    switch ($label) {
        "tentative" { 2 }
        "accepted" { 3 }
        "declined" { 4 }
        default { -1 }
    }
}

function Get-ParticipantHint([string]$responseSubject, [string]$associatedSubject) {
    $cleanResp = Normalize-MeetingSubject $responseSubject
    $cleanAssoc = Normalize-MeetingSubject $associatedSubject
    if ([string]::IsNullOrWhiteSpace($cleanResp)) { return "" }

    if (-not [string]::IsNullOrWhiteSpace($cleanAssoc)) {
        if ($cleanResp.Equals($cleanAssoc, [System.StringComparison]::OrdinalIgnoreCase)) { return "" }
        $prefix = $cleanAssoc + " - "
        if ($cleanResp.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $cleanResp.Substring($prefix.Length).Trim()
        }
    }

    if ($cleanResp -match '\s-\s') {
        return ($cleanResp -split '\s-\s', 2)[1].Trim()
    }
    return ""
}

function Get-PrimaryMatchKey([string]$entryId, [object]$startAt) {
    if ([string]::IsNullOrWhiteSpace($entryId)) { return "" }
    $dt = $null
    try { $dt = [datetime]$startAt } catch {}
    if ($dt -eq $null) { return "" }
    return ("$entryId|" + $dt.ToString("yyyy-MM-ddTHH:mm"))
}

function Get-SecondaryMatchKey([string]$subject, [object]$startAt) {
    $norm = Normalize-MeetingSubject $subject
    if ([string]::IsNullOrWhiteSpace($norm)) { return "" }
    $dt = $null
    try { $dt = [datetime]$startAt } catch {}
    if ($dt -eq $null) { return "" }
    return ("$norm|" + $dt.ToString("yyyy-MM-ddTHH:mm"))
}

function Add-MapValue([hashtable]$map, [string]$key, [object]$value) {
    if ([string]::IsNullOrWhiteSpace($key)) { return }
    if (-not $map.ContainsKey($key)) {
        $map[$key] = [System.Collections.ArrayList]::new()
    }
    [void]$map[$key].Add($value)
}

function New-EmptySummary() {
    return @{
        total = 0
        accepted = 0
        tentative = 0
        declined = 0
        notResponded = 0
        organized = 0
        none = 0
        unknown = 0
    }
}

function Get-MeetingResponseEvidence([object]$namespace, [datetime]$windowStart, [datetime]$windowEnd) {
    $evidence = [System.Collections.ArrayList]::new()
    $folderIds = @(6, 5, 3) # Inbox, Sent, Deleted Items

    foreach ($folderId in $folderIds) {
        $folder = $null
        try { $folder = $namespace.GetDefaultFolder($folderId) } catch {}
        if ($null -eq $folder) { continue }

        $itemsLocal = $null
        try {
            $itemsLocal = $folder.Items
            $itemsLocal.Sort("[ReceivedTime]", $true)
        } catch {
            continue
        }
        if ($null -eq $itemsLocal) { continue }

        $scanLimit = 0
        try { $scanLimit = [Math]::Min(2500, [int]$itemsLocal.Count) } catch { $scanLimit = 800 }
        if ($scanLimit -le 0) { $scanLimit = 800 }

        for ($i = 1; $i -le $scanLimit; $i++) {
            $msg = $null
            try { $msg = $itemsLocal.Item($i) } catch { continue }
            if ($null -eq $msg) { continue }

            $msgClass = ""
            try { $msgClass = [string]$msg.MessageClass } catch {}
            if ($msgClass -notlike "IPM.Schedule.Meeting.Resp*") { continue }

            $receivedTime = $null
            try { $receivedTime = [datetime]$msg.ReceivedTime } catch {}
            if ($receivedTime -ne $null) {
                if ($receivedTime -lt $windowStart.AddDays(-30) -or $receivedTime -gt $windowEnd.AddDays(3)) { continue }
            }

            $assoc = $null
            try { $assoc = $msg.GetAssociatedAppointment($false) } catch {}
            if ($null -eq $assoc) { continue }

            $assocStart = $null
            try { $assocStart = [datetime]$assoc.Start } catch {}
            if ($assocStart -eq $null) { continue }
            if ($assocStart -lt $windowStart.AddDays(-1) -or $assocStart -gt $windowEnd.AddDays(1)) { continue }

            $assocEntryId = ""
            try { $assocEntryId = [string]$assoc.EntryID } catch {}
            $assocSubject = ""
            try { $assocSubject = [string]$assoc.Subject } catch {}
            $responseSubject = ""
            try { $responseSubject = [string]$msg.Subject } catch {}
            $responseStatus = Get-ResponseStatusFromMessage -messageClass $msgClass -subject $responseSubject
            $responseStatusCode = Get-ResponseStatusCodeFromLabel -label $responseStatus
            $participantHint = Get-ParticipantHint -responseSubject $responseSubject -associatedSubject $assocSubject
            $responseEntryId = ""
            try { $responseEntryId = [string]$msg.EntryID } catch {}
            $receivedTimeIso = ""
            if ($receivedTime -ne $null) { $receivedTimeIso = $receivedTime.ToString("o") }

            [void]$evidence.Add(@{
                responseEntryId = $responseEntryId
                responseStatus = $responseStatus
                responseStatusCode = $responseStatusCode
                responseMessageClass = $msgClass
                responseSubject = $responseSubject
                participantHint = $participantHint
                responderName = ""
                responderEmail = ""
                receivedTime = $receivedTimeIso
                associatedEntryId = $assocEntryId
                associatedSubject = $assocSubject
                associatedStart = $assocStart.ToString("o")
                source = "mail_response"
            })
        }
    }

    return @($evidence)
}

$result = [System.Collections.ArrayList]::new()
foreach ($item in $restricted) {
    if ($null -eq $item) { continue }
    if ($result.Count -ge $TopVal) { break }
    $subjectVal = ""
    try { $subjectVal = [string]$item.Subject } catch {}
    $locationVal = ""
    try { $locationVal = [string]$item.Location } catch {}
    $organizer = ""
    try { $organizer = [string]$item.Organizer } catch {}
    $entryId = ""
    try { $entryId = [string]$item.EntryID } catch {}
    $body = ""
    try { $body = [string]$item.Body; if ($body.Length -gt 500) { $body = $body.Substring(0,500) } } catch {}
    $messageClass = ""
    try { $messageClass = [string]$item.MessageClass } catch {}
    $isMeeting = $false
    try { $isMeeting = [bool]$item.IsMeeting } catch {}
    $meetingStatusCode = -1
    try { $meetingStatusCode = [int]$item.MeetingStatus } catch {}
    $responseRequested = $false
    try { $responseRequested = [bool]$item.ResponseRequested } catch {}
    $requiredAttendeesRaw = ""
    try { $requiredAttendeesRaw = [string]$item.RequiredAttendees } catch {}
    $optionalAttendeesRaw = ""
    try { $optionalAttendeesRaw = [string]$item.OptionalAttendees } catch {}
    $myResponseStatusCode = -1
    $myResponseStatus = "unknown"
    try {
        $myResponseStatusCode = [int]$item.ResponseStatus
        $myResponseStatus = Get-ResponseStatusLabel $myResponseStatusCode
    } catch {}
    $snapshot = Get-AttendeesSnapshot -calendarItem $item

    $eventStart = $null
    $eventEnd = $null
    $eventStartIso = ""
    $eventEndIso = ""
    try { $eventStart = [datetime]$item.Start; $eventStartIso = $eventStart.ToString("o") } catch {}
    try { $eventEnd = [datetime]$item.End; $eventEndIso = $eventEnd.ToString("o") } catch {}
    if ([string]::IsNullOrWhiteSpace($subjectVal) -and [string]::IsNullOrWhiteSpace($entryId) -and [string]::IsNullOrWhiteSpace($eventStartIso)) {
        continue
    }
    $primaryKey = Get-PrimaryMatchKey -entryId $entryId -startAt $eventStart
    $secondaryKey = Get-SecondaryMatchKey -subject $subjectVal -startAt $eventStart
    $summaryToUse = $snapshot.summary
    $attendeesToUse = $snapshot.attendees
    $metadataAvailable = ($snapshot.attendees.Count -gt 0) -or (-not [string]::IsNullOrWhiteSpace($organizer)) -or (-not [string]::IsNullOrWhiteSpace($requiredAttendeesRaw)) -or (-not [string]::IsNullOrWhiteSpace($optionalAttendeesRaw))
    $recoveredFromMail = $false
    $mailSummary = New-EmptySummary

    [void]$result.Add(@{
        Subject   = $subjectVal
        Start     = $eventStartIso
        End       = $eventEndIso
        Location  = $locationVal
        Organizer = $organizer
        EntryID   = $entryId
        Body      = $body
        MyResponseStatus = $myResponseStatus
        MyResponseStatusCode = $myResponseStatusCode
        MessageClass = $messageClass
        IsMeeting = $isMeeting
        MeetingStatusCode = $meetingStatusCode
        ResponseRequested = $responseRequested
        RequiredAttendeesRaw = $requiredAttendeesRaw
        OptionalAttendeesRaw = $optionalAttendeesRaw
        AttendeeMetadataAvailable = $metadataAvailable
        ResponseRecoveredFromMail = $recoveredFromMail
        MailResponseSummary = $mailSummary
        MailResponseEvents = @()
        Attendees = $attendeesToUse
        AttendeeResponseSummary = $summaryToUse
        _MatchPrimaryKey = $primaryKey
        _MatchSecondaryKey = $secondaryKey
    })
}

if ($result.Count -gt 0) {
    $responseByPrimaryKey = @{}
    $responseBySecondaryKey = @{}
    $responseEvidence = Get-MeetingResponseEvidence -namespace $ns -windowStart $startDate -windowEnd $endDate
    foreach ($resp in $responseEvidence) {
        $respStartAt = $null
        try { $respStartAt = [datetime]$resp.associatedStart } catch {}
        if ($respStartAt -eq $null) { continue }
        $k1 = Get-PrimaryMatchKey -entryId $resp.associatedEntryId -startAt $respStartAt
        $k2 = Get-SecondaryMatchKey -subject $resp.associatedSubject -startAt $respStartAt
        Add-MapValue -map $responseByPrimaryKey -key $k1 -value $resp
        Add-MapValue -map $responseBySecondaryKey -key $k2 -value $resp
    }

    foreach ($evt in $result) {
        $primaryKey = [string]$evt._MatchPrimaryKey
        $secondaryKey = [string]$evt._MatchSecondaryKey
        $mailResponses = [System.Collections.ArrayList]::new()
        $seenResponse = @{}

        if (-not [string]::IsNullOrWhiteSpace($primaryKey) -and $responseByPrimaryKey.ContainsKey($primaryKey)) {
            foreach ($r in $responseByPrimaryKey[$primaryKey]) {
                $sig = [string]$r.responseEntryId
                if ([string]::IsNullOrWhiteSpace($sig)) { $sig = ([string]$r.responseSubject + "|" + [string]$r.receivedTime + "|" + [string]$r.responseStatus) }
                if (-not $seenResponse.ContainsKey($sig)) {
                    $seenResponse[$sig] = $true
                    [void]$mailResponses.Add($r)
                }
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($secondaryKey) -and $responseBySecondaryKey.ContainsKey($secondaryKey)) {
            foreach ($r in $responseBySecondaryKey[$secondaryKey]) {
                $sig = [string]$r.responseEntryId
                if ([string]::IsNullOrWhiteSpace($sig)) { $sig = ([string]$r.responseSubject + "|" + [string]$r.receivedTime + "|" + [string]$r.responseStatus) }
                if (-not $seenResponse.ContainsKey($sig)) {
                    $seenResponse[$sig] = $true
                    [void]$mailResponses.Add($r)
                }
            }
        }

        $mailResponseSorted = @($mailResponses | Sort-Object receivedTime)
        $mailSummary = New-EmptySummary
        foreach ($mr in $mailResponseSorted) {
            switch ([string]$mr.responseStatus) {
                "accepted" { $mailSummary.accepted++ }
                "tentative" { $mailSummary.tentative++ }
                "declined" { $mailSummary.declined++ }
                default { $mailSummary.unknown++ }
            }
        }
        $mailSummary.total = $mailResponseSorted.Count
        $evt.MailResponseSummary = $mailSummary
        $evt.MailResponseEvents = $mailResponseSorted

        $curSummary = $evt.AttendeeResponseSummary
        $curTotal = 0
        try { $curTotal = [int]$curSummary.total } catch {}

        # MailResponseEvents counts replies, never distinct attendees.

    }
}

foreach ($evt in $result) {
    [void]$evt.Remove("_MatchPrimaryKey")
    [void]$evt.Remove("_MatchSecondaryKey")
}

if ($result.Count -eq 0) { Write-Output '[]' } else { $result | ConvertTo-Json -Depth 6 }
}
