# Administrative sync only. This script is never called by an MCP tool.
function Export-IndexAttachments {
    param($Item, [string]$Directory)
    $Directory=[IO.Path]::GetFullPath($Directory)
    $allowed = @('.pdf','.docx','.pptx','.xlsx','.csv','.txt','.md')
    $results = @(); $selected = 0; $total = 0
    for ($i=1; $i -le $Item.Attachments.Count; $i++) {
        $att=$Item.Attachments.Item($i)
        $ext=[IO.Path]::GetExtension([string]$att.FileName).ToLowerInvariant()
        $entry=@{index=$i;name=[string]$att.FileName;status='unsupported';path=''}
        if ($allowed -notcontains $ext) { $results += $entry; continue }
        if ($selected -ge 6) { $entry.status='count_limit'; $results += $entry; continue }
        $selected++
        if ([int64]$att.Size -gt 10485760 -or $total+[int64]$att.Size -gt 41943040) {
            $entry.status='size_limit'; $results += $entry; continue
        }
        if ([int]$att.Type -ne 1) { $entry.status='unsupported_type'; $results += $entry; continue }
        # Never use the untrusted original filename as an output path.
        $target=Join-Path $Directory (('{0:D3}' -f $i)+$ext)
        try {
            $att.SaveAsFile($target)
            if ((Get-Item -LiteralPath $target).Length -gt 10485760) { throw 'Export exceeded size limit' }
            $total+=(Get-Item -LiteralPath $target).Length
            $entry.status='exported'; $entry.path=$target
        } catch { $entry.status='export_error' }
        $results += $entry
    }
    return $results
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference='Stop'
    [Console]::InputEncoding=[Text.UTF8Encoding]::new($false)
    [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
    try {
        $req=[Console]::In.ReadToEnd()|ConvertFrom-Json
        $dir=[IO.Path]::GetFullPath([string]$req.directory)
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) { throw 'Missing sync temporary directory' }
        $ol=[Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        $ns=$ol.GetNamespace('MAPI')
        if ($req.store_id) { $item=$ns.GetItemFromID([string]$req.entry_id,[string]$req.store_id) }
        else { $item=$ns.GetItemFromID([string]$req.entry_id) }
        $result=@(Export-IndexAttachments $item $dir)
        ConvertTo-Json -InputObject $result -Depth 4 -Compress
    } catch {
        @{error='Attachment export failed; verify the Outlook session and retry this page'}|ConvertTo-Json -Compress
        exit 1
    }
}
