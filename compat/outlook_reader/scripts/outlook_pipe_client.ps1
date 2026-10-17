# outlook_pipe_client.ps1
# Thin pipe client: forwards JSON-line requests from its own stdin to
# the Outlook daemon named pipe, and writes the daemon's responses
# back to its own stdout. Loops until stdin closes (EOF).
#
# Use cases:
#   * Bash: "echo '{...}' | powershell.exe -File outlook_pipe_client.ps1"
#     (one PS flash per bash invocation, then JSON stream in/out)
#   * Python subprocess.Popen with stdin=PIPE, stdout=PIPE for
#     multi-request multiplexing in email_rag.py
#
# Wire format (both sides):
#   * stdin: one JSON object per line
#   * stdout: one JSON object per line
#   * stderr: human readable errors only (DO NOT write JSON to stderr)

param(
    [string]$PipeName = "OpenClawOutlookDaemon",
    [int]$ConnectTimeoutMs = 10000
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$ErrorActionPreference = "Continue"

$pipe = $null
$writer = $null
$reader = $null
$connected = $false

function Connect-Pipe {
    if ($script:connected) { return $true }
    try {
        $script:pipe = New-Object System.IO.Pipes.NamedPipeClientStream(
            ".", $PipeName, [System.IO.Pipes.PipeDirection]::InOut
        )
        $script:pipe.Connect($ConnectTimeoutMs)
    } catch {
        [Console]::Error.WriteLine("Pipe connect failed: $($_.Exception.Message)")
        return $false
    }
    try {
        $script:writer = New-Object System.IO.StreamWriter($script:pipe, [System.Text.Encoding]::UTF8, 1024, $true)
        $script:writer.AutoFlush = $true
        $script:reader = New-Object System.IO.StreamReader($script:pipe, [System.Text.Encoding]::UTF8, $false, 1024, $true)
        $script:connected = $true
    } catch {
        [Console]::Error.WriteLine("Pipe stream init failed: $($_.Exception.Message)")
        try { $script:pipe.Close() } catch {}
        $script:connected = $false
        return $false
    }
    return $true
}

function Disconnect-Pipe {
    if (-not $script:connected) { return }
    try { $script:reader.Dispose() } catch {}
    try { $script:writer.Dispose() } catch {}
    try { $script:pipe.Close() } catch {}
    $script:connected = $false
}

try {
    while ($true) {
        $line = $null
        try { $line = [Console]::In.ReadLine() } catch { break }
        if ($null -eq $line) { break }
        $line = $line.Trim()
        if ([string]::IsNullOrEmpty($line)) { continue }

        if (-not (Connect-Pipe)) {
            [Console]::Out.WriteLine((@{ error = "Cannot connect to Outlook daemon pipe ($PipeName)" } | ConvertTo-Json -Compress))
            continue
        }

        try {
            $script:writer.WriteLine($line)
            $response = $script:reader.ReadLine()
            if ($null -eq $response) {
                Disconnect-Pipe
                [Console]::Out.WriteLine((@{ error = "daemon disconnected" } | ConvertTo-Json -Compress))
                continue
            }
            [Console]::Out.WriteLine($response)
        } catch {
            Disconnect-Pipe
            [Console]::Out.WriteLine((@{ error = "pipe io error: $($_.Exception.Message)" } | ConvertTo-Json -Compress)
            )
        }
    }
} catch {
    [Console]::Error.WriteLine("Pipe client fatal: $($_.Exception.Message)")
} finally {
    Disconnect-Pipe
}
