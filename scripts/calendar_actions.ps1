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
        modified=$item.LastModificationTime.ToUniversalTime().ToString('o');saved=[bool]$item.Saved;meeting_status=[int]$item.MeetingStatus}
}
function Get-ActionOutlook { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
function Get-ActionFailure($record, $operation) {
    $discard='not_needed'
    if ($script:MutationAttempted -and -not $script:SendAttempted -and $script:DiscardEligible -and $null -ne $script:OwnedItem) {
        try { $script:OwnedItem.Close(1); $discard='close_discard_returned' }
        catch { $discard='discard_failed' }
    }
    $exception=$record.Exception
    while ($null -ne $exception.InnerException) { $exception=$exception.InnerException }
    $state='read_failed'
    if ($operation -eq 'commit') {
        $state='rejected_before_mutation'
        if ($script:MutationAttempted) { $state='failed_after_mutation' }
        if ($script:SendAttempted) { $state='send_outcome_unknown' }
        if ($script:SendReturned) { $state='submitted_result_incomplete' }
    }
    return @{error=$record.Exception.Message;source='windows_outlook';status=$state;
        phase=$script:ActionPhase;exception_type=$exception.GetType().FullName;hresult=('0x'+$exception.HResult.ToString('X8'));
        mutation_attempted=[bool]$script:MutationAttempted;send_attempted=[bool]$script:SendAttempted;
        send_returned=[bool]$script:SendReturned;discard=$discard;
        note='Never infer saved or sent from a live attendee list. No automatic retry; obtain fresh preview and human approval for any new attempt.'}
}
function Run($request) {
    $script:ActionPhase='connect_outlook'
    $script:MutationAttempted=$false
    $script:SendAttempted=$false
    $script:SendReturned=$false
    $script:OwnedItem=$null
    $script:DiscardEligible=$false
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
    $script:ActionPhase='validate_account'
    $p = $request.payload
    $account = @($ns.Accounts | Where-Object { $_.SmtpAddress -eq $p.sender_account })
    if ($account.Count -ne 1) { throw 'Sender account unavailable or ambiguous' }
    if ($p.action -eq 'create') {
        $script:ActionPhase='create_unsaved_item'
        $script:MutationAttempted=$true
        $item = $account[0].DeliveryStore.GetDefaultFolder(9).Items.Add('IPM.Appointment')
        $script:OwnedItem=$item
        $script:DiscardEligible=$true
        $script:ActionPhase='set_new_item_account'
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
        $script:ActionPhase='read_existing_event'
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
        $script:ActionPhase='validate_write_access'
        $access=[int]$item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x0FF40003')
        if (($access -band 1) -eq 0) { throw 'Outlook reports no modify permission on this event; no changes applied' }
        $script:ActionPhase='validate_editable_state'
        if (-not $item.Saved) { throw 'Event has unsaved changes; close or save it in Outlook before preparing again' }
        foreach ($inspector in $app.Inspectors) {
            if ([string]$inspector.CurrentItem.EntryID -eq [string]$item.EntryID) { throw 'Event is open in Outlook; close its window before preparing again' }
        }
        # An occurrence inherits its sending account. Reassigning an existing
        # event's account is unnecessary and may be rejected by Outlook.
        if ([string]$item.SendUsingAccount.SmtpAddress -ne $p.sender_account) { throw 'Existing event sending account differs; no changes applied' }
        $script:OwnedItem=$item
        $script:DiscardEligible=$true
        switch ($p.action) {
            'update' {
                $script:ActionPhase='modify_event_fields'
                $script:MutationAttempted=$true
                if ($p.subject) { $item.Subject=$p.subject }
                if ($p.body) { $item.Body=$p.body }
                if ($p.location) { $item.Location=$p.location }
                if ($p.start) { $item.StartUTC=[DateTimeOffset]::Parse($p.start).UtcDateTime; $item.EndUTC=[DateTimeOffset]::Parse($p.end).UtcDateTime }
                $script:ActionPhase='set_update_distribution'
                $item.ForceUpdateToAllAttendees=$true
            }
            'add_remove_attendees' {
                # Resolve all additions before touching existing recipients.
                foreach ($address in $p.add_attendees) { $r=$ns.CreateRecipient($address); if (-not $r.Resolve()) { throw 'Unresolved new attendee' } }
                $script:ActionPhase='modify_attendees'
                $script:MutationAttempted=$true
                for ($i=$item.Recipients.Count; $i -ge 1; $i--) { if ((Smtp $item.Recipients.Item($i)) -in $p.remove_attendees) { $item.Recipients.Remove($i) } }
                foreach ($address in $p.add_attendees) { $r=$item.Recipients.Add($address); $r.Type=1 }
                $script:ActionPhase='set_update_distribution'
                $item.ForceUpdateToAllAttendees = $true
            }
            'cancel' { $script:ActionPhase='cancel_event';$script:MutationAttempted=$true;$item.MeetingStatus=5 }
            { $_ -in @('accept','tentative','decline') } {
                if ($current.organizer_smtp -eq $p.sender_account) { throw 'Organizer cannot respond' }
                if ($item.MeetingStatus -ne 3) { throw 'Not a received meeting invitation' }
                $codes=@{accept=3;tentative=2;decline=4}
                $script:ActionPhase='respond_to_invitation'
                $script:MutationAttempted=$true
                $script:DiscardEligible=$false
                $response=$item.Respond($codes[$p.action],$true,$false)
                $script:ActionPhase='set_response_account'
                $response.SendUsingAccount=$account[0]
                $script:ActionPhase='send_response'
                $script:SendAttempted=$true
                $response.Send()
                $script:SendReturned=$true
                return @{status='submitted_to_outlook';action=$p.action;sent=$true;send_attempted=$true;send_returned=$true;delivery_verified=$false}
            }
            default { throw 'Unsupported calendar action' }
        }
    }
    $script:ActionPhase='resolve_final_recipients'
    if (-not $item.Recipients.ResolveAll()) { throw 'Unresolved recipients; nothing submitted' }
    $script:ActionPhase='send_invitation'
    $script:SendAttempted=$true
    $item.Send()
    $script:SendReturned=$true
    return @{status='submitted_to_outlook';action=$p.action;sent=$true;send_attempted=$true;send_returned=$true;delivery_verified=$false;entry_id=[string]$item.EntryID}
}
if ($MyInvocation.InvocationName -ne '.') {
    try { $request=[Console]::In.ReadToEnd() | ConvertFrom-Json; Run $request | ConvertTo-Json -Depth 20 -Compress }
    catch { Get-ActionFailure $_ $request.operation | ConvertTo-Json -Depth 6 -Compress; exit 1 }
}
