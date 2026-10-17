param([int]$Days = 7, [int]$Top = 20)

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
            error = "Outlook process detected (PID $pidList), but COM connection still failed after retry. Ensure both run at the same privilege level (both admin or both normal user) and restart Outlook normally (not as admin)."
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
$cal = $ns.GetDefaultFolder(9)
$items = $cal.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true

$startDate = (Get-Date).Date
$endDate = $startDate.AddDays($Days)
$filter = "[End] > " + [char]39 + $startDate.ToString("g") + [char]39 + " AND [Start] < " + [char]39 + $endDate.ToString("g") + [char]39
$result = @()

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

foreach ($item in $items.Restrict($filter)) {
    if ($result.Count -ge $Top) { break }
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

    $result += @{
        Subject = $item.Subject
        Start = $item.Start.ToString("o")
        End = $item.End.ToString("o")
        Location = $item.Location
        Organizer = $item.Organizer
        EntryID = $item.EntryID
        MyResponseStatus = $myResponseStatus
        MyResponseStatusCode = $myResponseStatusCode
        MessageClass = $messageClass
        IsMeeting = $isMeeting
        MeetingStatusCode = $meetingStatusCode
        ResponseRequested = $responseRequested
        RequiredAttendeesRaw = $requiredAttendeesRaw
        OptionalAttendeesRaw = $optionalAttendeesRaw
        AttendeeMetadataAvailable = ($snapshot.attendees.Count -gt 0) -or (-not [string]::IsNullOrWhiteSpace([string]$item.Organizer)) -or (-not [string]::IsNullOrWhiteSpace($requiredAttendeesRaw)) -or (-not [string]::IsNullOrWhiteSpace($optionalAttendeesRaw))
        Attendees = $snapshot.attendees
        AttendeeResponseSummary = $snapshot.summary
    }
}

$result | ConvertTo-Json -Depth 6
