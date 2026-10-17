[CmdletBinding()]
param([switch]$RegisterClient,[switch]$WithEmbeddings,[switch]$SkipTests)
$ErrorActionPreference='Stop'
Set-Location -LiteralPath $PSScriptRoot
$env:PIP_CACHE_DIR=Join-Path $PSScriptRoot '.cache/pip'
$env:npm_config_cache=Join-Path $PSScriptRoot '.cache/npm'
if (-not (Get-Command py.exe -ErrorAction SilentlyContinue)) { throw 'Install Python 3.12-3.14 for Windows (with Python launcher), then rerun. WSL Python is not supported.' }
& py.exe -3 -c 'import sys; assert (3,12) <= sys.version_info[:2] <= (3,14), "Python 3.12-3.14 required"'
if($LASTEXITCODE){throw 'Unsupported native Python version'}
if (-not (Test-Path '.venv/Scripts/python.exe')) {
    & py.exe -3 -m venv .venv
    if($LASTEXITCODE){throw 'Native Python venv creation failed'}
}
$python=Join-Path $PSScriptRoot '.venv/Scripts/python.exe'
& $python -m pip install -r requirements.lock.txt
if($LASTEXITCODE){throw 'Dependency installation failed'}
New-Item -ItemType Directory -Force data,models,artifacts | Out-Null
# Personal mail and config backups must remain private even in a shared checkout.
$sid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
foreach($dir in @('data','models','artifacts')) {
    & icacls.exe (Join-Path $PSScriptRoot $dir) /inheritance:r /grant:r "*${sid}:(OI)(CI)F" '*S-1-5-18:(OI)(CI)F' | Out-Null
    if($LASTEXITCODE){throw "Cannot protect $dir permissions"}
}
if($WithEmbeddings){
    if(-not (Get-Command node.exe -ErrorAction SilentlyContinue)){throw 'Install Node.js 22 LTS or newer, then rerun with -WithEmbeddings'}
    Push-Location scripts
    try { & npm.cmd ci --omit=dev; if($LASTEXITCODE){throw 'Embedding dependency installation failed'} } finally { Pop-Location }
    & $python scripts/install_model.py
    if($LASTEXITCODE){throw 'Embedding model setup failed'}
}
if(-not $SkipTests){
    & $python -m unittest discover -s tests -v
    if($LASTEXITCODE){throw 'Regression tests failed'}
}
& $python -m outlook_mcp.cli doctor
if($LASTEXITCODE){throw 'Doctor failed'}
if($RegisterClient){
    & $python scripts/register_client.py
    if($LASTEXITCODE){throw 'Client registration needs attention'}
}
Write-Host 'Installed native Windows OutlookMCP. Run the smoke test after indexing, then reload the client MCP connection.'
