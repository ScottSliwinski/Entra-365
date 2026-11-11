Connect-ExchangeOnline
$disabledUsers = Get-Mailbox -RecipientTypeDetails UserMailbox | Where-Object {
    (Get-User $_.UserPrincipalName).Enabled -eq $false
}
$disabledUsers = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

$results = @()

foreach ($mailbox in $disabledUsers) {
    $permissions = Get-MailboxPermission -Identity $mailbox.PrimarySmtpAddress | Where-Object {
        $_.AccessRights -contains "FullAccess" -and $_.IsInherited -eq $false
    }
    foreach ($permission in $permissions) {
        $results += [PSCustomObject]@{
            Mailbox       = $mailbox.PrimarySmtpAddress
            Delegate      = $permission.User
            AccessRights  = $permission.AccessRights -join ", "
        }
    }
}
$results | Export-Csv -Path "$env:USERPROFILE\OneDrive - Andrew Wommack Ministries\Documents\PowerShell\Scripts\DisabledUsersWithDelegates.csv" -NoTypeInformation -Encoding UTF8
