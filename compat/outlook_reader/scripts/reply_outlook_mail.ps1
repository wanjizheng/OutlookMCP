param(
    [Parameter(Mandatory = $true)]
    [string]$EntryID,

    [Parameter(Mandatory = $true)]
    [string]$Body,

    [Parameter(Mandatory = $false)]
    [string]$DoSend = "false",

    [Parameter(Mandatory = $false)]
    [string]$DoDraft = "true"
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$sendFlag = [System.Convert]::ToBoolean($DoSend)
$draftFlag = [System.Convert]::ToBoolean($DoDraft)

if ($sendFlag -and $draftFlag) {
    @{ error = "DoSend and DoDraft are mutually exclusive" } | ConvertTo-Json -Depth 2
    exit 2
}

try {
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
                error = "Outlook process detected (PID $pidList), but COM connection still failed after retry. Ensure both run at same privilege level (both normal user) and restart Outlook if needed."
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
    $original = $ns.GetItemFromID($EntryID)
    if (-not $original) {
        @{ error = "Mail not found by EntryID"; entryId = $EntryID } | ConvertTo-Json -Depth 2
        exit 1
    }

    $reply = $original.Reply()
    $reply.Body = "$Body`r`n`r`n$($reply.Body)"

    if ($sendFlag) {
        $reply.Send()
        @{
            status = "sent"
            entryId = $EntryID
            to = $reply.To
            subject = $reply.Subject
            sentAt = (Get-Date).ToString("o")
        } | ConvertTo-Json -Depth 2
    }
    elseif ($draftFlag) {
        $reply.Save()
        @{
            status = "draft"
            entryId = $EntryID
            to = $reply.To
            subject = $reply.Subject
            savedAt = (Get-Date).ToString("o")
        } | ConvertTo-Json -Depth 2
    }
    else {
        @{
            status = "preview"
            entryId = $EntryID
            to = $reply.To
            subject = $reply.Subject
            bodyPreview = $reply.Body.Substring(0, [Math]::Min(600, $reply.Body.Length))
            note = "Default is draft. Use --send to actually send."
        } | ConvertTo-Json -Depth 3
    }
}
catch {
    @{
        error = $_.Exception.Message
        type = $_.Exception.GetType().FullName
    } | ConvertTo-Json -Depth 2
    exit 1
}
