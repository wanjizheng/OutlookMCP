param(
    [Parameter(Mandatory = $false)]
    [string]$QueryBase64 = "",

    [Parameter(Mandatory = $false)]
    [string]$Query = "",

    [Parameter(Mandatory = $false)]
    [ValidateSet("all", "contacts", "groups", "members")]
    [string]$Mode = "all",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 2000)]
    [int]$Top = 100,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 2000)]
    [int]$MemberTop = 500
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$ErrorActionPreference = "Stop"

if (-not [string]::IsNullOrWhiteSpace($QueryBase64)) {
    try {
        $Query = [System.Text.Encoding]::UTF8.GetString(
            [System.Convert]::FromBase64String($QueryBase64)
        )
    } catch {
        @{
            error = "Invalid QueryBase64 value"
            detail = $_.Exception.Message
        } | ConvertTo-Json -Depth 3
        exit 2
    }
}
$Query = $Query.Trim()

function Get-SafeString {
    param([object]$Object, [string]$PropertyName)
    if ($null -eq $Object) { return "" }
    try {
        $value = $Object.$PropertyName
        if ($null -eq $value) { return "" }
        return [string]$value
    } catch {
        return ""
    }
}

function Get-SafeInt {
    param([object]$Object, [string]$PropertyName, [int]$DefaultValue = -1)
    if ($null -eq $Object) { return $DefaultValue }
    try { return [int]$Object.$PropertyName } catch { return $DefaultValue }
}

function Test-TextMatch {
    param([string]$Value, [string]$Needle)
    if ([string]::IsNullOrWhiteSpace($Needle)) { return $true }
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return $Value.IndexOf($Needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
}

function Get-SmtpAddressFromAddressEntry {
    param([object]$AddressEntry)
    if ($null -eq $AddressEntry) { return "" }

    $smtp = ""
    try {
        $smtp = [string]$AddressEntry.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        )
    } catch {}
    if (-not [string]::IsNullOrWhiteSpace($smtp)) { return $smtp }

    $type = Get-SafeString $AddressEntry "Type"
    if ($type.Equals("EX", [System.StringComparison]::OrdinalIgnoreCase)) {
        try {
            $exchangeUser = $AddressEntry.GetExchangeUser()
            if ($null -ne $exchangeUser) {
                $smtp = Get-SafeString $exchangeUser "PrimarySmtpAddress"
            }
        } catch {}
        if (-not [string]::IsNullOrWhiteSpace($smtp)) { return $smtp }

        try {
            $exchangeList = $AddressEntry.GetExchangeDistributionList()
            if ($null -ne $exchangeList) {
                $smtp = Get-SafeString $exchangeList "PrimarySmtpAddress"
            }
        } catch {}
        if (-not [string]::IsNullOrWhiteSpace($smtp)) { return $smtp }
    }

    try {
        $contact = $AddressEntry.GetContact()
        if ($null -ne $contact) {
            $smtp = Get-SafeString $contact "Email1Address"
        }
    } catch {}
    if (-not [string]::IsNullOrWhiteSpace($smtp)) { return $smtp }

    $rawAddress = Get-SafeString $AddressEntry "Address"
    if ($rawAddress -match "^[^@\s]+@[^@\s]+\.[^@\s]+$") {
        return $rawAddress
    }
    return ""
}

function Get-ContactEmail {
    param([object]$Contact, [int]$Slot, [object]$Namespace)

    $address = Get-SafeString $Contact ("Email{0}Address" -f $Slot)
    $addressType = Get-SafeString $Contact ("Email{0}AddressType" -f $Slot)
    $displayName = Get-SafeString $Contact ("Email{0}DisplayName" -f $Slot)

    if ($addressType.Equals("EX", [System.StringComparison]::OrdinalIgnoreCase)) {
        try {
            $entryId = Get-SafeString $Contact ("Email{0}EntryID" -f $Slot)
            if (-not [string]::IsNullOrWhiteSpace($entryId)) {
                $addressEntry = $Namespace.GetAddressEntryFromID($entryId)
                $smtp = Get-SmtpAddressFromAddressEntry $addressEntry
                if (-not [string]::IsNullOrWhiteSpace($smtp)) { $address = $smtp }
            }
        } catch {}
    }

    if ([string]::IsNullOrWhiteSpace($address)) { return $null }
    return @{
        Slot = $Slot
        Address = $address
        AddressType = $addressType
        DisplayName = $displayName
    }
}

function Convert-ContactItem {
    param([object]$Contact, [string]$FolderPath, [object]$Namespace)

    $emails = [System.Collections.ArrayList]::new()
    foreach ($slot in 1..3) {
        $email = Get-ContactEmail -Contact $Contact -Slot $slot -Namespace $Namespace
        if ($null -ne $email) { [void]$emails.Add($email) }
    }

    return @{
        Type = "contact"
        EntryID = Get-SafeString $Contact "EntryID"
        FullName = Get-SafeString $Contact "FullName"
        FirstName = Get-SafeString $Contact "FirstName"
        LastName = Get-SafeString $Contact "LastName"
        CompanyName = Get-SafeString $Contact "CompanyName"
        JobTitle = Get-SafeString $Contact "JobTitle"
        Department = Get-SafeString $Contact "Department"
        Emails = @($emails)
        BusinessTelephoneNumber = Get-SafeString $Contact "BusinessTelephoneNumber"
        MobileTelephoneNumber = Get-SafeString $Contact "MobileTelephoneNumber"
        HomeTelephoneNumber = Get-SafeString $Contact "HomeTelephoneNumber"
        BusinessAddress = Get-SafeString $Contact "BusinessAddress"
        HomeAddress = Get-SafeString $Contact "HomeAddress"
        Categories = Get-SafeString $Contact "Categories"
        FolderPath = $FolderPath
    }
}

function Convert-GroupMember {
    param([object]$Recipient)

    $resolved = $false
    try { $resolved = [bool]$Recipient.Resolved } catch {}

    $addressEntry = $null
    try { $addressEntry = $Recipient.AddressEntry } catch {}
    $smtp = Get-SmtpAddressFromAddressEntry $addressEntry
    if ([string]::IsNullOrWhiteSpace($smtp)) {
        try {
            $smtp = [string]$Recipient.PropertyAccessor.GetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
            )
        } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($smtp)) {
        $raw = Get-SafeString $Recipient "Address"
        if ($raw -match "^[^@\s]+@[^@\s]+\.[^@\s]+$") { $smtp = $raw }
    }

    return @{
        Name = Get-SafeString $Recipient "Name"
        EmailAddress = $smtp
        AddressType = Get-SafeString $addressEntry "Type"
        Resolved = $resolved
        EntryID = Get-SafeString $addressEntry "ID"
    }
}

$outlook = $null
$attachError = ""
for ($attempt = 1; $attempt -le 3 -and $null -eq $outlook; $attempt++) {
    try {
        $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        $attachError = $_.Exception.Message
    }
    if ($null -eq $outlook) {
        try {
            $outlook = New-Object -ComObject Outlook.Application
        } catch {
            $attachError = $_.Exception.Message
        }
    }
    if ($null -eq $outlook -and $attempt -lt 3) {
        Start-Sleep -Milliseconds (300 * $attempt)
    }
}

if ($null -eq $outlook) {
    $running = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
    $pidList = ""
    if ($running) { $pidList = (($running | ForEach-Object { $_.Id }) -join ",") }
    @{
        error = "Unable to connect to classic Outlook after 3 attempts. Start or restart classic Outlook and ensure it runs at the same privilege level."
        detail = $attachError
        outlookProcessIds = $pidList
    } | ConvertTo-Json -Depth 3
    exit 1
}

try {
    $namespace = $outlook.GetNamespace("MAPI")
} catch {
    @{
        error = "Connected to Outlook but could not open the MAPI namespace"
        detail = $_.Exception.Message
    } | ConvertTo-Json -Depth 3
    exit 1
}

$contactFolders = [System.Collections.ArrayList]::new()
$folderKeys = @{}
$folderQueue = New-Object System.Collections.Queue

try { $folderQueue.Enqueue($namespace.GetDefaultFolder(10)) } catch {}
try {
    for ($storeIndex = 1; $storeIndex -le $namespace.Stores.Count; $storeIndex++) {
        try {
            $store = $namespace.Stores.Item($storeIndex)
            $folderQueue.Enqueue($store.GetDefaultFolder(10))
        } catch {}
    }
} catch {}

while ($folderQueue.Count -gt 0) {
    $folder = $folderQueue.Dequeue()
    if ($null -eq $folder) { continue }
    $folderKey = Get-SafeString $folder "EntryID"
    if ([string]::IsNullOrWhiteSpace($folderKey)) {
        $folderKey = Get-SafeString $folder "FolderPath"
    }
    if ([string]::IsNullOrWhiteSpace($folderKey) -or $folderKeys.ContainsKey($folderKey)) {
        continue
    }
    $folderKeys[$folderKey] = $true
    [void]$contactFolders.Add($folder)
    try {
        for ($childIndex = 1; $childIndex -le $folder.Folders.Count; $childIndex++) {
            $folderQueue.Enqueue($folder.Folders.Item($childIndex))
        }
    } catch {}
}

if ($contactFolders.Count -eq 0) {
    @{
        error = "No Outlook Contacts folder is available in the current profile"
        query = $Query
        mode = $Mode
    } | ConvertTo-Json -Depth 3
    exit 1
}

$contacts = [System.Collections.ArrayList]::new()
$groups = [System.Collections.ArrayList]::new()
$warnings = [System.Collections.ArrayList]::new()
$scannedItems = 0
$contactsTruncated = $false
$groupsTruncated = $false

foreach ($folder in $contactFolders) {
    $folderPath = Get-SafeString $folder "FolderPath"
    try {
        $items = $folder.Items
        for ($itemIndex = 1; $itemIndex -le $items.Count; $itemIndex++) {
            $item = $null
            try { $item = $items.Item($itemIndex) } catch { continue }
            if ($null -eq $item) { continue }
            $scannedItems++

            $classCode = Get-SafeInt $item "Class"
            $messageClass = Get-SafeString $item "MessageClass"
            $isContact = ($classCode -eq 40) -or $messageClass.StartsWith(
                "IPM.Contact",
                [System.StringComparison]::OrdinalIgnoreCase
            )
            $isGroup = ($classCode -eq 69) -or $messageClass.StartsWith(
                "IPM.DistList",
                [System.StringComparison]::OrdinalIgnoreCase
            )

            if ($isContact -and @("all", "contacts") -contains $Mode) {
                $contact = Convert-ContactItem -Contact $item -FolderPath $folderPath -Namespace $namespace
                $searchValues = @(
                    $contact.FullName,
                    $contact.FirstName,
                    $contact.LastName,
                    $contact.CompanyName,
                    $contact.JobTitle,
                    $contact.Department
                )
                foreach ($email in $contact.Emails) { $searchValues += [string]$email.Address }
                $matches = [string]::IsNullOrWhiteSpace($Query)
                if (-not $matches) {
                    foreach ($value in $searchValues) {
                        if (Test-TextMatch $value $Query) { $matches = $true; break }
                    }
                }
                if ($matches) {
                    if ($contacts.Count -lt $Top) {
                        [void]$contacts.Add($contact)
                    } else {
                        $contactsTruncated = $true
                    }
                }
                continue
            }

            if ($isGroup -and @("all", "groups", "members") -contains $Mode) {
                $groupName = Get-SafeString $item "DLName"
                if ([string]::IsNullOrWhiteSpace($groupName)) {
                    $groupName = Get-SafeString $item "Subject"
                }
                $nameMatched = Test-TextMatch $groupName $Query
                $matches = $nameMatched
                if ($Mode -eq "members") {
                    $matches = -not [string]::IsNullOrWhiteSpace($Query) -and $nameMatched
                }
                if (-not $matches) { continue }

                $memberCount = Get-SafeInt $item "MemberCount" 0
                $members = [System.Collections.ArrayList]::new()
                $membersTruncated = $false
                $includeMembers = ($Mode -eq "members") -or
                    ($Mode -eq "all" -and -not [string]::IsNullOrWhiteSpace($Query))

                if ($includeMembers) {
                    for ($memberIndex = 1; $memberIndex -le $memberCount; $memberIndex++) {
                        $recipient = $null
                        try { $recipient = $item.GetMember($memberIndex) } catch { continue }
                        if ($null -eq $recipient) { continue }
                        $member = Convert-GroupMember $recipient
                        if ($members.Count -lt $MemberTop) {
                            [void]$members.Add($member)
                        } else {
                            $membersTruncated = $true
                        }
                    }
                }

                if ($matches) {
                    $matchType = "contains"
                    if ($groupName.Equals($Query, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $matchType = "exact"
                    }
                    $group = @{
                        Type = "group"
                        Name = $groupName
                        EntryID = Get-SafeString $item "EntryID"
                        FolderPath = $folderPath
                        MemberCount = $memberCount
                        MembersReturned = $members.Count
                        MembersTruncated = $membersTruncated
                        MatchType = $matchType
                        Members = @($members)
                    }
                    if ($groups.Count -lt $Top) {
                        [void]$groups.Add($group)
                    } else {
                        $groupsTruncated = $true
                    }
                }
            }
        }
    } catch {
        [void]$warnings.Add(@{
            FolderPath = $folderPath
            Error = $_.Exception.Message
        })
    }
}

$folderPaths = @($contactFolders | ForEach-Object { Get-SafeString $_ "FolderPath" })
$response = @{
    Query = $Query
    Mode = $Mode
    ContactFolders = $folderPaths
    ScannedItems = $scannedItems
    MatchedContactCount = $contacts.Count
    MatchedGroupCount = $groups.Count
    ContactsTruncated = $contactsTruncated
    GroupsTruncated = $groupsTruncated
    Contacts = @($contacts)
    Groups = @($groups)
    Warnings = @($warnings)
}
$response | ConvertTo-Json -Depth 8
