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
    [string]$DoDraft = "true",

    [string]$BCC = "",

    [string]$SenderAccount = ""
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'outlook_send_library.ps1')

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

    Invoke-OutlookComSend -Outlook $ol -To $To -CC $CC -BCC $BCC -SenderAccount $SenderAccount -Subject $Subject -Body $Body -DoSend $sendFlag -DoDraft $draftFlag | ConvertTo-Json -Depth 3
} catch {
    @{
        error = $_.Exception.Message
        type = $_.Exception.GetType().FullName
    } | ConvertTo-Json -Depth 2
    exit 1
}
