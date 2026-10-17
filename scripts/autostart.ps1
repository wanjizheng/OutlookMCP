param(
    [ValidateSet('Install','Status','Restart','Stop','Uninstall')]
    [string]$Action = 'Status',
    [switch]$EnableSend
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$taskName = 'OutlookMCP-Windows-' + $identity.User.Value
$description = "OutlookMCP interactive Windows supervisor; checkout=$root"
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing -and $existing.Description -ne $description) {
    throw 'Existing OutlookMCP task belongs to another checkout; inspect it before replacing.'
}
function Stop-OwnedTask {
    Stop-ScheduledTask -TaskName $taskName
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-ScheduledTask -TaskName $taskName).State -eq 'Running') {
        if ((Get-Date) -gt $deadline) { throw 'Timed out stopping OutlookMCP task.' }
        Start-Sleep -Milliseconds 200
    }
}
if ($Action -eq 'Install') {
    foreach ($relative in @('.venv\Scripts\pythonw.exe','scripts\tunnel_watchdog.py',
        'artifacts\tunnel-client\runtime-key.dpapi','artifacts\tunnel-client\profiles\outlook.yaml',
        'artifacts\tunnel-client\tunnel-client.exe')) {
        if (-not (Test-Path -LiteralPath (Join-Path $root $relative))) { throw "Missing prerequisite: $relative" }
    }
    $arguments = '"' + (Join-Path $root 'scripts\tunnel_watchdog.py') + '"'
    if ($EnableSend) { $arguments += ' --enable-send' }
    $taskAction = New-ScheduledTaskAction -Execute (Join-Path $root '.venv\Scripts\pythonw.exe') -Argument $arguments -WorkingDirectory $root
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $identity.Name
    # A periodic trigger also covers forced termination/manual-start cases where
    # Task Scheduler's RestartOnFailure alone does not reliably requeue a run.
    $heartbeat = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
        -RepetitionInterval (New-TimeSpan -Minutes 1)
    $principal = New-ScheduledTaskPrincipal -UserId $identity.Name -LogonType Interactive -RunLevel Limited
    $settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero) `
        -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
    if ($existing) {
        Disable-ScheduledTask -TaskName $taskName | Out-Null
        Stop-OwnedTask
    }
    Register-ScheduledTask -TaskName $taskName -Action $taskAction -Trigger @($trigger,$heartbeat) `
        -Principal $principal -Settings $settings -Description $description -Force | Out-Null
    Start-ScheduledTask -TaskName $taskName
} elseif ($Action -in @('Restart','Stop','Uninstall')) {
    if (-not $existing) { throw 'OutlookMCP startup task is not installed.' }
    Disable-ScheduledTask -TaskName $taskName | Out-Null
    Stop-OwnedTask
    if ($Action -eq 'Restart') {
        Enable-ScheduledTask -TaskName $taskName | Out-Null
        Start-ScheduledTask -TaskName $taskName
    }
    if ($Action -eq 'Uninstall') { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false }
}
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($task) {
    $info = Get-ScheduledTaskInfo -TaskName $taskName
    [pscustomobject]@{ TaskName=$taskName; State=$task.State; LastTaskResult=$info.LastTaskResult;
        Enabled=$task.Settings.Enabled; NextRunTime=$info.NextRunTime;
        LogonType=$task.Principal.LogonType; RunLevel=$task.Principal.RunLevel;
        Arguments=$task.Actions.Arguments }
} else { Write-Output 'OutlookMCP startup task not installed.' }
