#Requires -Version 5.1
<#
.SYNOPSIS
Replaces direct Full Access and Send As delegates on one Exchange Online shared mailbox.
.EXAMPLE
.\Replace-SharedMailboxMembers.ps1 -Mailbox shared@contoso.com -CsvPath .\members.csv
.EXAMPLE
.\Replace-SharedMailboxMembers.ps1 -Mailbox shared@contoso.com
.NOTES
CSV must have a UserPrincipalName header and at least one row:
UserPrincipalName
alice@contoso.com
bob@contoso.com
Requires the ExchangeOnlineManagement module and rights to manage mailbox permissions.
Use -WhatIf to preview commands without making changes. Review the displayed plan
and type the mailbox address to confirm a live replacement.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)][string]$Mailbox,
    [string]$CsvPath,
    [string]$AdminUPN
)

$ErrorActionPreference = 'Stop'

function Get-PrincipalKey {
    param($Entry)
    foreach ($name in @('UserPrincipalName', 'PrimarySmtpAddress', 'WindowsEmailAddress')) {
        if ($Entry.PSObject.Properties[$name] -and $Entry.$name) {
            return ([string]$Entry.$name).Trim().ToLowerInvariant()
        }
    }
    return ([string]$Entry.Identity).Trim().ToLowerInvariant()
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop
$connect = @{}
if ($AdminUPN) { $connect.UserPrincipalName = $AdminUPN }
Connect-ExchangeOnline @connect -ShowBanner:$false -ErrorAction Stop

try {
    $target = Get-EXOMailbox -Identity $Mailbox -Properties RecipientTypeDetails -ErrorAction Stop
    if ($target.RecipientTypeDetails -ne 'SharedMailbox') {
        throw "'$Mailbox' is not a shared mailbox."
    }
    $mailboxAddress = [string]$target.PrimarySmtpAddress

    if ($CsvPath) {
        if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { throw "CSV not found: $CsvPath" }
        $rows = @(Import-Csv -LiteralPath $CsvPath -ErrorAction Stop)
        if (-not $rows.Count -or -not $rows[0].PSObject.Properties['UserPrincipalName']) {
            throw 'CSV must contain at least one row and a UserPrincipalName column.'
        }
        $requested = @($rows | ForEach-Object { ([string]$_.UserPrincipalName).Trim() })
        if (@($requested | Where-Object { -not $_ }).Count) { throw 'CSV contains a blank UserPrincipalName.' }
    }
    else {
        Write-Host 'Enter replacement user UPNs, one at a time. Press Enter on a blank line when finished.'
        $requested = @()
        while ($true) {
            $value = (Read-Host 'User UPN').Trim()
            if (-not $value) { break }
            $requested += $value
        }
    }
    $requested = @($requested | Select-Object -Unique)
    if (-not $requested.Count) { throw 'No replacement users were provided. Nothing changed.' }

    # Resolve every replacement user before touching any existing permissions.
    $users = @()
    foreach ($upn in $requested) {
        if ($upn -notmatch '^[^@\s]+@[^@\s]+$') { throw "Invalid UPN: $upn" }
        $recipient = Get-EXOMailbox -UserPrincipalName $upn -ErrorAction Stop
        if ($recipient.RecipientTypeDetails -notin @('UserMailbox', 'SharedMailbox')) {
            throw "'$upn' is not a user or shared mailbox recipient."
        }
        $resolvedUPN = [string]$recipient.UserPrincipalName
        if ($resolvedUPN -ine $upn) { throw "'$upn' resolved to '$resolvedUPN'; use the exact UPN in the CSV." }
        if ($resolvedUPN -ieq [string]$target.UserPrincipalName) { throw 'The shared mailbox cannot be its own delegate.' }
        $users += $resolvedUPN
    }
    $users = @($users | Select-Object -Unique)

    # Only direct, non-deny Full Access and direct Send As are in scope.
    $full = @(Get-EXOMailboxPermission -Identity $mailboxAddress -ErrorAction Stop |
        Where-Object { -not $_.IsInherited -and -not $_.Deny -and $_.AccessRights -contains 'FullAccess' -and
            [string]$_.User -notmatch '^(NT AUTHORITY\\|S-1-5-(10|18)$)' -and [string]$_.User -ne 'SELF' })
    $send = @(Get-EXORecipientPermission -Identity $mailboxAddress -ErrorAction Stop |
        Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'SendAs' -and
            [string]$_.Trustee -notmatch '^(NT AUTHORITY\\|S-1-5-(10|18)$)' -and [string]$_.Trustee -ne 'SELF' })

    Write-Host "`nShared mailbox: $mailboxAddress"
    Write-Host "Existing direct Full Access: $($full.Count)"
    $full | ForEach-Object { Write-Host "  $($_.User)" }
    Write-Host "Existing direct Send As: $($send.Count)"
    $send | ForEach-Object { Write-Host "  $($_.Trustee)" }
    Write-Host "Replacement users (each receives Full Access and Send As): $($users.Count)"
    $users | ForEach-Object { Write-Host "  $_" }

    if (-not $WhatIfPreference) {
        $answer = Read-Host "To replace these permissions, type the mailbox address ($mailboxAddress)"
        if ($answer.Trim() -ine $mailboxAddress) { throw 'Confirmation did not match. Nothing changed.' }
    }

    $failures = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $full) {
        $principal = [string]$entry.User
        if ($PSCmdlet.ShouldProcess("$mailboxAddress / $principal", 'Remove direct Full Access')) {
            try { Remove-MailboxPermission -Identity $mailboxAddress -User $principal -AccessRights FullAccess -Confirm:$false -ErrorAction Stop }
            catch { $failures.Add("Remove Full Access [$principal]: $($_.Exception.Message)") }
        }
    }
    foreach ($entry in $send) {
        $principal = [string]$entry.Trustee
        if ($PSCmdlet.ShouldProcess("$mailboxAddress / $principal", 'Remove direct Send As')) {
            try { Remove-RecipientPermission -Identity $mailboxAddress -Trustee $principal -AccessRights SendAs -Confirm:$false -ErrorAction Stop }
            catch { $failures.Add("Remove Send As [$principal]: $($_.Exception.Message)") }
        }
    }

    foreach ($upn in $users) {
        if ($PSCmdlet.ShouldProcess("$mailboxAddress / $upn", 'Add Full Access')) {
            try { Add-MailboxPermission -Identity $mailboxAddress -User $upn -AccessRights FullAccess -InheritanceType All -ErrorAction Stop | Out-Null }
            catch { $failures.Add("Add Full Access [$upn]: $($_.Exception.Message)") }
        }
        if ($PSCmdlet.ShouldProcess("$mailboxAddress / $upn", 'Add Send As')) {
            try { Add-RecipientPermission -Identity $mailboxAddress -Trustee $upn -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null }
            catch { $failures.Add("Add Send As [$upn]: $($_.Exception.Message)") }
        }
    }

    if ($WhatIfPreference) { Write-Host '`nWhatIf complete. No permissions changed.'; return }
    Write-Host '`nCurrent direct Full Access:'
    Get-EXOMailboxPermission -Identity $mailboxAddress | Where-Object { -not $_.IsInherited -and -not $_.Deny -and $_.AccessRights -contains 'FullAccess' } |
        Select-Object User,AccessRights | Format-Table -AutoSize
    Write-Host 'Current direct Send As:'
    Get-EXORecipientPermission -Identity $mailboxAddress | Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'SendAs' } |
        Select-Object Trustee,AccessRights | Format-Table -AutoSize
    if ($failures.Count) {
        $failures | ForEach-Object { Write-Warning $_ }
        throw "$($failures.Count) permission operation(s) failed. Review the current permissions above."
    }
    Write-Host 'Replacement complete.'
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
