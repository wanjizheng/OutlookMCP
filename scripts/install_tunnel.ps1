$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$target = Join-Path $root 'artifacts\tunnel-client'
New-Item -ItemType Directory -Force -Path $target | Out-Null
$release = Invoke-RestMethod 'https://api.github.com/repos/openai/tunnel-client/releases/latest'
$arch = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq 'Arm64') { 'arm64' } else { 'amd64' }
$asset = $release.assets | Where-Object { $_.name -eq "tunnel-client-$($release.tag_name)-windows-$arch.zip" }
if (-not $asset -or -not $asset.digest.StartsWith('sha256:')) { throw 'Official release asset/digest unavailable' }
$archive = Join-Path $target 'download.zip'
Invoke-WebRequest $asset.browser_download_url -OutFile $archive
$digest = 'sha256:' + (Get-FileHash -Algorithm SHA256 -LiteralPath $archive).Hash.ToLower()
if ($digest -ne $asset.digest) { throw 'Official release digest mismatch' }
Expand-Archive -LiteralPath $archive -DestinationPath $target -Force
@{version=$release.tag_name;sha256=$digest} | ConvertTo-Json | Set-Content (Join-Path $target 'release.json')
& (Join-Path $target 'tunnel-client.exe') help quickstart
