#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
Reports direct shared-mailbox delegate permissions for users in a CSV.
.EXAMPLE
.\Check-BulkSharedMailboxAccess.ps1 -InputCsv .\Disabled_Users.csv -OutputCsv .\SharedMailboxAccess.csv
.NOTES
Input CSV must contain a UserPrincipalName column. Reports direct FullAccess,
SendAs, and SendOnBehalf permissions. Group membership and mailbox folder
permissions are outside this report.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$InputCsv,

    [string]$OutputCsv = (Join-Path (Get-Location) 'SharedMailboxAccess.csv')
)

$ErrorActionPreference = 'Stop'
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false

$inputRows = @(Import-Csv -LiteralPath $InputCsv)
if ($inputRows.Count -eq 0 -or 'UserPrincipalName' -notin $inputRows[0].PSObject.Properties.Name) {
    throw 'The input CSV needs a UserPrincipalName header and at least one account.'
}

# Resolve the input UPNs to primary SMTP addresses, as Exchange may return
# aliases or display names in its permission lists.
$targets = @{}
$upns = @($inputRows | ForEach-Object { [string]$_.UserPrincipalName } |
    ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
foreach ($upn in $upns) {
    try {
        $recipient = Get-Recipient -Identity $upn -ErrorAction Stop
        $key = $recipient.PrimarySmtpAddress.ToString().ToLowerInvariant()
        $targets[$key] = $upn
    }
    catch {
        Write-Warning "Could not resolve input account '$upn': $($_.Exception.Message)"
    }
}
if ($targets.Count -eq 0) { throw 'No input accounts could be resolved in Exchange Online.' }

$recipientCache = @{}
function Resolve-DelegateAddress {
    param([Parameter(Mandatory = $true)][string]$Identity)
    $cacheKey = $Identity.ToLowerInvariant()
    if ($recipientCache.ContainsKey($cacheKey)) { return $recipientCache[$cacheKey] }
    try {
        $resolved = Get-Recipient -Identity $Identity -ErrorAction Stop
        $recipientCache[$cacheKey] = $resolved.PrimarySmtpAddress.ToString().ToLowerInvariant()
    }
    catch {
        $recipientCache[$cacheKey] = $null
    }
    return $recipientCache[$cacheKey]
}

$report = [System.Collections.Generic.List[object]]::new()
$mailboxes = @(Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited)
$index = 0
foreach ($mailbox in $mailboxes) {
    $index++
    $address = $mailbox.PrimarySmtpAddress.ToString()
    Write-Progress -Activity 'Checking shared mailboxes' `
        -Status "$index of $($mailboxes.Count): $address" `
        -PercentComplete (($index / $mailboxes.Count) * 100)

    try {
        $permissionSets = @(
            @{ Type = 'FullAccess'; Entries = @(Get-EXOMailboxPermission -Identity $address |
                Where-Object { -not $_.IsInherited -and -not $_.Deny -and
                    $_.AccessRights -contains 'FullAccess' }); IdentityProperty = 'User' },
            @{ Type = 'SendAs'; Entries = @(Get-EXORecipientPermission -Identity $address |
                Where-Object { -not $_.IsInherited -and -not $_.Deny -and
                    $_.AccessRights -contains 'SendAs' }); IdentityProperty = 'Trustee' },
            @{ Type = 'SendOnBehalf'; Entries = @((Get-Mailbox -Identity $address).GrantSendOnBehalfTo);
                IdentityProperty = $null }
        )
        foreach ($set in $permissionSets) {
            foreach ($entry in $set.Entries) {
                if ($null -eq $entry) { continue }
                $identity = if ($set.IdentityProperty) {
                    $entry.($set.IdentityProperty).ToString()
                } else { $entry.ToString() }
                $delegateAddress = Resolve-DelegateAddress -Identity $identity
                if ($delegateAddress -and $targets.ContainsKey($delegateAddress)) {
                    $report.Add([pscustomobject]@{
                        UserPrincipalName = $targets[$delegateAddress]
                        SharedMailbox     = $address
                        AccessType        = $set.Type
                    })
                }
            }
        }
    }
    catch {
        Write-Progress -Activity 'Checking shared mailboxes' -Completed
        throw "Failed checking '$address'. The report was not exported because it would be incomplete. $($_.Exception.Message)"
    }
}
Write-Progress -Activity 'Checking shared mailboxes' -Completed

# Always write a headed CSV, even if no direct assignments were found.
$columns = 'UserPrincipalName', 'SharedMailbox', 'AccessType'
$sorted = @($report | Sort-Object -Property $columns -Unique)
if ($sorted.Count) {
    $sorted | Select-Object $columns | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
} else {
    '"UserPrincipalName","SharedMailbox","AccessType"' |
        Set-Content -LiteralPath $OutputCsv -Encoding UTF8
}

Write-Host "Checked $($mailboxes.Count) shared mailboxes for $($targets.Count) resolved users."
Write-Host "Found $($sorted.Count) direct permissions. Report: $OutputCsv"
