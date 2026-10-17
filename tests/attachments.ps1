$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../scripts/attachment_export.ps1')
$dir=Join-Path (Join-Path $PSScriptRoot '../artifacts') ([guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $dir|Out-Null
try {
    $script:saves=0
    $script:rows=@()
    foreach($i in 1..8){
        $att=[pscustomobject]@{FileName='../../same.txt';Size=10;Type=1}
        $att|Add-Member ScriptMethod SaveAsFile {param($p) $script:saves++;[IO.File]::WriteAllText($p,'fixture')}
        $script:rows+=$att
    }
    $collection=[pscustomobject]@{Count=8}
    $collection|Add-Member ScriptMethod Item {param($i) return $script:rows[$i-1]}
    $item=[pscustomobject]@{Attachments=$collection}
    $out=@(Export-IndexAttachments $item $dir)
    if($script:saves -ne 6 -or @($out|Where-Object {$_.status -eq 'count_limit'}).Count -ne 2){throw 'Count limit failed'}
    if(@(Get-ChildItem -LiteralPath $dir -File).Count -ne 6){throw 'Colliding filename handling failed'}
    if(@($out|Where-Object {$_.path -and [IO.Path]::GetDirectoryName($_.path) -ne [IO.Path]::GetFullPath($dir)}).Count){throw 'Path escaped temporary directory'}
    $script:rows[0].Size=10485761
    $script:rows[1].Type=5
    $script:rows[2].FileName='image.png'
    $out=@(Export-IndexAttachments $item $dir)
    if($out[0].status -ne 'size_limit' -or $out[1].status -ne 'unsupported_type' -or $out[2].status -ne 'unsupported'){throw 'Unsupported/oversize policy failed'}
    'Administrative export filename/type/size/count behavior PASS'
} finally {
    # The only files here were created by this isolated fixture.
    Get-ChildItem -LiteralPath $dir -File|Remove-Item
    Remove-Item -LiteralPath $dir
}
