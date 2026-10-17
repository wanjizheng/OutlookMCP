# Shared by the existing skill send script and the native MCP bridge.
function Invoke-OutlookComSend {
    param($Outlook, [string]$To, [string]$Subject, [string]$Body,
        [string]$CC='', [string]$BCC='', [string]$SenderAccount='',
        [bool]$DoSend=$false, [bool]$DoDraft=$true)
    $mail = $null
    try {
        $account = $null
        if ($SenderAccount) {
            $accounts = $Outlook.Session.Accounts
            for ($i=1; $i -le $accounts.Count; $i++) {
                $candidate = $accounts.Item($i)
                if ([string]$candidate.SmtpAddress -ieq $SenderAccount) { $account=$candidate; break }
            }
            if (-not $account) { throw 'Sender SMTP account not found' }
        }
        $mail = $Outlook.CreateItem(0)
        if ($account) { $mail.SendUsingAccount=$account }
        if ($SenderAccount) {
            # Native MCP uses exact SMTP recipients, not Outlook display-name guesses.
            $sets = @(@{Addresses=$To;Type=1},@{Addresses=$CC;Type=2},@{Addresses=$BCC;Type=3})
            foreach ($set in $sets) {
                foreach ($address in ($set.Addresses -split ';' | Where-Object { $_ })) {
                    $recipient = $mail.Recipients.Add($address)
                    $recipient.Type = $set.Type
                    if (-not $recipient.Resolve()) { throw 'Recipient cannot be resolved' }
                    $entry = $recipient.AddressEntry
                    $resolved = [string]$entry.Address
                    if ([string]$entry.Type -eq 'EX') {
                        $user = $entry.GetExchangeUser()
                        if (-not $user) { throw 'Distribution list is not supported by the native sender' }
                        $resolved = [string]$user.PrimarySmtpAddress
                    }
                    if ($resolved -ine $address) { throw 'Resolved SMTP recipient differs from approved preview' }
                }
            }
        } else {
            # Preserve legacy skill display-name/default-account behavior.
            $mail.To=$To
            if ($CC) { $mail.CC=$CC }
            if ($BCC) { $mail.BCC=$BCC }
        }
        $mail.Subject=$Subject
        $mail.Body=$Body
        if ($DoSend) {
            $mail.Send()
            $mail = $null
            return @{status='sent';sent=$true;to=$To;cc=$CC;bcc=$BCC;subject=$Subject;
                sentAt=(Get-Date).ToString('o');note='Submitted to Outlook; final delivery is not established. Do not auto-retry.'}
        } elseif ($DoDraft) {
            $mail.Save()
            return @{status='draft';to=$To;cc=$CC;subject=$Subject;savedAt=(Get-Date).ToString('o')}
        } else {
            return @{status='preview';to=$To;cc=$CC;subject=$Subject;body=$Body;note='Default is draft. Use --send to actually send.'}
        }
    } finally {
        if ($null -ne $mail) { $mail.Close(1) }
    }
}
