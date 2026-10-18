# Strict read-only source for native synchronization. Never saves Outlook items.
$ErrorActionPreference='Stop'
[Console]::InputEncoding=[Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
try {
    $req=[Console]::In.ReadToEnd()|ConvertFrom-Json
    $ol=[Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $ns=$ol.GetNamespace('MAPI')
    $store=$ns.DefaultStore
    if ($store.ExchangeStoreType -ne 0) { throw 'Primary Exchange mailbox required' }
    if ($ns.Offline) { throw 'Outlook offline' }
    $connectionMode=[int]$ns.ExchangeConnectionMode
    function Stamp($value) {
        if ($null -eq $value) { return '' }
        return ([datetime]$value).ToUniversalTime().ToString('o')
    }
    if ($req.operation -eq 'inventory') {
        $excluded=@{}
        foreach ($n in @(3,4,16,23)) {
            try { $excluded[[string]$store.GetDefaultFolder($n).EntryID]=$true } catch {
                if ($n -ne 23) { throw }
            }
        }
        # Drafts=16, Junk=23. Search folders are excluded using folder type.
        $rows=[Collections.Generic.List[object]]::new()
        $folders=[Collections.Generic.List[string]]::new()
        function Scan($folder) {
            $id=[string]$folder.EntryID
            if ($excluded.ContainsKey($id)) { return }
            $kind=$folder.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003')
            if ([int]$kind -eq 2) { return }
            if ([int]$folder.DefaultItemType -eq 0) {
                $folders.Add($id)
                $table=$folder.GetTable()
                $table.Columns.RemoveAll()
                foreach ($column in @('http://schemas.microsoft.com/mapi/proptag/0x66700102','MessageClass','Subject','ReceivedTime','LastModificationTime')) {
                    [void]$table.Columns.Add($column)
                }
                while (-not $table.EndOfTable) {
                    $batch=$table.GetArray(500)
                    for ($i=0; $i -lt $batch.GetLength(0); $i++) {
                        if ([string]$batch.GetValue($i,1) -notlike 'IPM.Note*') { continue }
                        # PR_LONGTERM_ENTRYID_FROM_TABLE avoids short-term IDs,
                        # without opening thousands of MailItem COM objects.
                        $canonical=[BitConverter]::ToString([byte[]]$batch.GetValue($i,0)).Replace('-','')
                        if (-not $canonical) { throw 'Missing canonical entry ID' }
                        $rows.Add(@{EntryID=$canonical;StoreID=[string]$store.StoreID;
                            FolderID=$id;Subject=[string]$batch.GetValue($i,2);
                            ReceivedTime=(Stamp $batch.GetValue($i,3));ModifiedTime=(Stamp $batch.GetValue($i,4))})
                    }
                }
            }
            for ($j=1; $j -le $folder.Folders.Count; $j++) { Scan ($folder.Folders.Item($j)) }
        }
        Scan ($store.GetRootFolder())
        if ($ns.Offline) { throw 'Outlook became offline' }
        # Cached Exchange tables can omit older server mail even while connected.
        # Never use that absence as proof of server-side deletion.
        @{complete=$true;server_complete=($connectionMode -eq 800 -and [int]$ns.ExchangeConnectionMode -eq 800);
            connection_mode=[int]$ns.ExchangeConnectionMode;
            store_id=[string]$store.StoreID;folders=@($folders.ToArray());items=@($rows.ToArray())} |
            ConvertTo-Json -Depth 6 -Compress
    } elseif ($req.operation -eq 'read_batch') {
        if (@($req.items).Count -gt 25) { throw 'Batch too large' }
        $result=@()
        foreach ($row in $req.items) {
            if ([string]$row.StoreID -ne [string]$store.StoreID) { throw 'Wrong mailbox' }
            $item=$ns.GetItemFromID([string]$row.EntryID,[string]$row.StoreID)
            if ([int]$item.Class -ne 43) { throw 'Not a mail item' }
            $body=[string]$item.Body
            $atts=@()
            for ($k=1;$k -le $item.Attachments.Count;$k++) {
                $a=$item.Attachments.Item($k)
                $atts+=@{FileName=[string]$a.FileName;Size=[int64]$a.Size;Type=[int]$a.Type}
            }
            $result+=@{EntryID=[string]$item.EntryID;StoreID=[string]$store.StoreID;
                Subject=[string]$item.Subject;Sender=[string]$item.SenderName;
                SenderEmailAddress=[string]$item.SenderEmailAddress;
                ReceivedTime=(Stamp $item.ReceivedTime);ModifiedTime=(Stamp $item.LastModificationTime);
                Body=$body.Substring(0,[Math]::Min(100000,$body.Length));Attachments=$atts}
        }
        @{items=@($result)}|ConvertTo-Json -Depth 6 -Compress
    } else { throw 'Invalid operation' }
} catch {
    # No message content in job logs; errors fail the entire inventory closed.
    @{error='Native source scan/read failed';error_type=$_.Exception.GetType().Name}|ConvertTo-Json -Compress
    exit 1
}
