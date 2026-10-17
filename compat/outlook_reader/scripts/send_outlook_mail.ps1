param(
    [Parameter(Mandatory=$true)]
    [string]$To,

    [Parameter(Mandatory=$true)]
    [string]$Subject,

    [Parameter(Mandatory=$true)]
    [string]$Body,

    [Parameter(Mandatory=$false)]
    [string]$CC = "",

    [Parameter(Mandatory=$false)]
    [string]$DoSend = "false",

    [Parameter(Mandatory=$false)]
    [string]$DoDraft = "true"
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# Convert string parameters to boolean
$sendFlag = [System.Convert]::ToBoolean($DoSend)
$draftFlag = [System.Convert]::ToBoolean($DoDraft)

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

    $mail = $ol.CreateItem(0)  # 0 = olMailItem
    $mail.To = $To
    if (-not [string]::IsNullOrWhiteSpace($CC)) {
        $mail.CC = $CC
    }
    $mail.Subject = $Subject
    $mail.Body = $Body

    if ($sendFlag) {
        $mail.Send()
        @{
            status = "sent"
            to = $mail.To
            cc = $mail.CC
            subject = $mail.Subject
            sentAt = (Get-Date).ToString("o")
        } | ConvertTo-Json -Depth 2
    } elseif ($draftFlag) {
        $mail.Save()
        @{
            status = "draft"
            to = $mail.To
            cc = $mail.CC
            subject = $mail.Subject
            savedAt = (Get-Date).ToString("o")
        } | ConvertTo-Json -Depth 2
    } else {
        @{
            status = "preview"
            to = $mail.To
            cc = $mail.CC
            subject = $mail.Subject
            body = $mail.Body
            note = "Default is draft. Use --send to actually send."
        } | ConvertTo-Json -Depth 2
    }
} catch {
    @{
        error = $_.Exception.Message
        type = $_.Exception.GetType().FullName
    } | ConvertTo-Json -Depth 2
    exit 1
}
