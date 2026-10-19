$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
function Smtp($recipient) {
    $address = $recipient
    if ($recipient.Class -eq 4) { $address = $recipient.AddressEntry }
    if ($address.Type -eq 'EX') {
        $user = $address.GetExchangeUser()
        if ($null -ne $user) { return [string]$user.PrimarySmtpAddress }
        $group = $address.GetExchangeDistributionList()
        if ($null -ne $group) { return [string]$group.PrimarySmtpAddress }
        throw 'Unable to resolve Exchange SMTP address'
    }
    return [string]$address.Address
}
function Event($ns, $p) {
    try { $item = $ns.GetItemFromID($p.entry_id, $p.store_id) }
    catch { throw ('Resolve event ID failed (HRESULT ' + $_.Exception.HResult + '): ' + $_.Exception.Message) }
    if ($item.Class -eq 53) {
        $item=$item.GetAssociatedAppointment($false)
        if ($null -eq $item) { throw 'Invitation has no existing calendar item; process it in Outlook first' }
    }
    if ($item.Class -ne 26) { throw 'Not an appointment' }
    if ($item.IsRecurring) {
        if ($p.scope -eq 'single') { throw 'Recurring event requires occurrence or series scope' }
        if ($p.scope -eq 'occurrence') {
            $date = [DateTimeOffset]::Parse($p.occurrence_start).LocalDateTime
            # Expanded occurrences/exceptions already represent one instance.
            # GetRecurrencePattern on such objects can fail in real Outlook.
            if ($item.RecurrenceState -eq 1) {
                try { $item = $item.GetRecurrencePattern().GetOccurrence($date) }
                catch { throw ('Resolve requested occurrence failed (HRESULT ' + $_.Exception.HResult + '): ' + $_.Exception.Message) }
            } elseif ($item.RecurrenceState -notin @(2,3)) {
                throw 'Unexpected recurring event state; list calendar again'
            }
            if ([Math]::Abs(($item.StartUTC - [DateTimeOffset]::Parse($p.occurrence_start).UtcDateTime).TotalSeconds) -gt 1) { throw 'Occurrence mismatch' }
        } elseif ($item.RecurrenceState -ne 1) { throw 'Series scope requires master EntryID' }
    } elseif ($p.scope -ne 'single') { throw 'Non-recurring event requires single scope' }
    return $item
}
function Snapshot($item) {
    $organizer = $item.GetOrganizer()
    $attendees = @()
    foreach ($r in $item.Recipients) { $attendees += [ordered]@{smtp=(Smtp $r);type=[int]$r.Type;name=[string]$r.Name} }
    return [ordered]@{entry_id=[string]$item.EntryID;subject=[string]$item.Subject;body=[string]$item.Body;
        start=$item.StartUTC.ToString('o');end=$item.EndUTC.ToString('o');location=[string]$item.Location;
        organizer_smtp=(Smtp $organizer);attendees=$attendees;recurrence_state=[int]$item.RecurrenceState;
        modified=$item.LastModificationTime.ToUniversalTime().ToString('o')}
}
function Get-ActionOutlook { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
function Run($request) {
    if ($env:OUTLOOK_MCP_ENABLE_SEND -ne '1') { throw 'Sending disabled' }
    $app = Get-ActionOutlook
    $ns = $app.GetNamespace('MAPI')
    if ($request.operation -eq 'capabilities') {
        $accounts = @($ns.Accounts | ForEach-Object { @{smtp=[string]$_.SmtpAddress;store_id=[string]$_.DeliveryStore.StoreID} })
        $providers = @($app.COMAddIns | ForEach-Object { @{name=[string]$_.ProgId;connected=[bool]$_.Connect} })
        return @{accounts=$accounts;addins=$providers;ordinary_invitations=$true;zoom_link_creation='Not configured: requires supported Zoom authorization/API; existing links can be included in body/location'}
    }
    if ($request.operation -eq 'resolve') {
        if ([string]::IsNullOrWhiteSpace($request.query) -or $request.query.Length -gt 254) { throw 'Invalid contact query' }
        $r = $ns.CreateRecipient($request.query)
        if (-not $r.Resolve()) { throw 'Contact unresolved or ambiguous; supply exact SMTP address' }
        return @{smtp=(Smtp $r);name=[string]$r.Name}
    }
    if ($request.operation -eq 'mail_read') {
        $mail=$ns.GetItemFromID($request.entry_id,$request.store_id)
        if ($mail.Class -ne 43) { throw 'Not a mail message' }
        return @{entry_id=[string]$mail.EntryID;store_id=[string]$request.store_id;subject=[string]$mail.Subject;body=[string]$mail.Body;modified=$mail.LastModificationTime.ToUniversalTime().ToString('o')}
    }
    if ($request.operation -eq 'read') { return (Snapshot (Event $ns $request)) }
    if ($request.operation -ne 'commit') { throw 'Unsupported operation' }
    $p = $request.payload
    $account = @($ns.Accounts | Where-Object { $_.SmtpAddress -eq $p.sender_account })
    if ($account.Count -ne 1) { throw 'Sender account unavailable or ambiguous' }
    if ($p.action -eq 'create') {
        $item = $account[0].DeliveryStore.GetDefaultFolder(9).Items.Add('IPM.Appointment')
        $item.SendUsingAccount = $account[0]
        $item.MeetingStatus = 1
        $item.Subject = $p.subject
        $item.Body = $p.body
        $item.Location = $p.location
        $item.StartUTC = [DateTimeOffset]::Parse($p.start).UtcDateTime
        $item.EndUTC = [DateTimeOffset]::Parse($p.end).UtcDateTime
        foreach ($address in $p.required) { $r=$item.Recipients.Add($address); $r.Type=1 }
        foreach ($address in $p.optional) { $r=$item.Recipients.Add($address); $r.Type=2 }
    } else {
        if ($account[0].DeliveryStore.StoreID -ne $p.store_id) { throw "Event does not belong to selected sending account" }
        $item = Event $ns $p
        $current = Snapshot $item
        # Compare again in the same COM session immediately before modifying.
        foreach ($key in $current.Keys) {
            if (($current[$key] | ConvertTo-Json -Depth 10 -Compress) -cne ($p.before.$key | ConvertTo-Json -Depth 10 -Compress)) { throw 'Event changed since preview; prepare again' }
        }
        if ($p.action -in @('update','add_remove_attendees','cancel')) {
            if ($current.organizer_smtp -ne $p.sender_account) { throw 'Only organizer may update meeting' }
            if ($p.sender_account -notin @($ns.Accounts | ForEach-Object { $_.SmtpAddress })) { throw 'Wrong organizer account' }
        }
        switch ($p.action) {
            'update' {
                if ($p.subject) { $item.Subject=$p.subject }
                if ($p.body) { $item.Body=$p.body }
                if ($p.location) { $item.Location=$p.location }
                if ($p.start) { $item.StartUTC=[DateTimeOffset]::Parse($p.start).UtcDateTime; $item.EndUTC=[DateTimeOffset]::Parse($p.end).UtcDateTime }
                $item.ForceUpdateToAllAttendees=$true
            }
            'add_remove_attendees' {
                # Resolve all additions before touching existing recipients.
                foreach ($address in $p.add_attendees) { $r=$ns.CreateRecipient($address); if (-not $r.Resolve()) { throw 'Unresolved new attendee' } }
                for ($i=$item.Recipients.Count; $i -ge 1; $i--) { if ((Smtp $item.Recipients.Item($i)) -in $p.remove_attendees) { $item.Recipients.Remove($i) } }
                foreach ($address in $p.add_attendees) { $r=$item.Recipients.Add($address); $r.Type=1 }
                $item.ForceUpdateToAllAttendees = $true
            }
            'cancel' { $item.MeetingStatus=5 }
            { $_ -in @('accept','tentative','decline') } {
                if ($current.organizer_smtp -eq $p.sender_account) { throw 'Organizer cannot respond' }
                if ($item.MeetingStatus -ne 3) { throw 'Not a received meeting invitation' }
                $codes=@{accept=3;tentative=2;decline=4}
                $response=$item.Respond($codes[$p.action],$true,$false)
                $response.SendUsingAccount=$account[0]
                $response.Send()
                return @{status='submitted_to_outlook';action=$p.action;sent=$true}
            }
            default { throw 'Unsupported calendar action' }
        }
    }
    $item.SendUsingAccount=$account[0]
    if (-not $item.Recipients.ResolveAll()) { throw 'Unresolved recipients; nothing submitted' }
    $item.Send()
    return @{status='submitted_to_outlook';action=$p.action;sent=$true;entry_id=[string]$item.EntryID}
}
if ($MyInvocation.InvocationName -ne '.') {
    try { $request=[Console]::In.ReadToEnd() | ConvertFrom-Json; Run $request | ConvertTo-Json -Depth 20 -Compress }
    catch {
        $suffix='; read-only operation, no mailbox write requested.'
        if ($request.operation -eq 'commit') { $suffix='; outcome may be uncertain, inspect Outlook before any retry.' }
        @{error=($_.Exception.Message + $suffix)} | ConvertTo-Json -Compress; exit 1
    }
}
