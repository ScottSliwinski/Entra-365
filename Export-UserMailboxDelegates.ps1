#requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
Exports direct Full Access, Send As, and Send on Behalf delegates for one user mailbox.
#>
$ErrorActionPreference = 'Stop'

$upn = (Read-Host 'Enter the UPN of the user mailbox to audit').Trim()
if ($upn -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') {
    throw 'Enter a valid user principal name, such as person@example.com.'
}

Import-Module ExchangeOnlineManagement
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

$mailbox = Get-EXOMailbox -Identity $upn -ErrorAction Stop
if ($mailbox.RecipientTypeDetails -ne 'UserMailbox') {
    throw "'$upn' is a $($mailbox.RecipientTypeDetails), not a user mailbox."
}

$mailboxAddress = [string]$mailbox.PrimarySmtpAddress
$rows = [System.Collections.Generic.List[object]]::new()

function Add-DelegateRow {
    param([string]$Identity, [string]$Permission)
    $delegateUpn = ''
    $delegateName = ''
    try {
        $recipient = Get-Recipient -Identity $Identity -ErrorAction Stop
        $delegateName = [string]$recipient.DisplayName
        if ($recipient.PSObject.Properties['UserPrincipalName']) {
            $delegateUpn = [string]$recipient.UserPrincipalName
        }
    } catch {
        # Preserve the raw identity for deleted or unresolvable delegates.
    }
    $rows.Add([pscustomobject]@{
        MailboxDisplayName    = [string]$mailbox.DisplayName
        MailboxUPN            = [string]$mailbox.UserPrincipalName
        MailboxEmail          = $mailboxAddress
        DelegateDisplayName   = $delegateName
        DelegateUPN           = $delegateUpn
        DelegateIdentity      = $Identity
        Permission            = $Permission
    })
}

Write-Host "Checking Full Access on $mailboxAddress..."
foreach ($entry in @(Get-EXOMailboxPermission -Identity $mailboxAddress -ErrorAction Stop)) {
    $identity = [string]$entry.User
    if (-not $entry.IsInherited -and -not $entry.Deny -and
        $entry.AccessRights -contains 'FullAccess' -and
        $identity -ne 'NT AUTHORITY\SELF') {
        Add-DelegateRow -Identity $identity -Permission 'Full Access'
    }
}

Write-Host "Checking Send As on $mailboxAddress..."
foreach ($entry in @(Get-EXORecipientPermission -Identity $mailboxAddress -ErrorAction Stop)) {
    $identity = [string]$entry.Trustee
    if (-not $entry.IsInherited -and -not $entry.Deny -and
        $entry.AccessRights -contains 'SendAs' -and
        $identity -ne 'NT AUTHORITY\SELF') {
        Add-DelegateRow -Identity $identity -Permission 'Send As'
    }
}

Write-Host "Checking Send on Behalf on $mailboxAddress..."
$behalfMailbox = Get-Mailbox -Identity $mailboxAddress -ErrorAction Stop
foreach ($delegate in @($behalfMailbox.GrantSendOnBehalfTo)) {
    if ($null -ne $delegate -and -not [string]::IsNullOrWhiteSpace([string]$delegate)) {
        Add-DelegateRow -Identity ([string]$delegate) -Permission 'Send on Behalf'
    }
}

$defaultName = 'UserMailboxDelegates_{0}_{1}.csv' -f ($mailboxAddress -replace '[^a-zA-Z0-9._-]', '_'), (Get-Date -Format 'yyyyMMdd_HHmmss')
$saveAt = (Read-Host "CSV path (press Enter for .\$defaultName)").Trim().Trim('"')
if (-not $saveAt) { $saveAt = Join-Path (Get-Location).Path $defaultName }
$parent = Split-Path -Parent $saveAt
if ($parent -and -not (Test-Path -LiteralPath $parent -PathType Container)) {
    throw "Output folder does not exist: $parent"
}

# Always create a CSV with headers, including when no delegates are found.
if ($rows.Count -gt 0) {
    $rows | Sort-Object DelegateIdentity, Permission | Export-Csv -LiteralPath $saveAt -NoTypeInformation -Encoding UTF8
} else {
    '"MailboxDisplayName","MailboxUPN","MailboxEmail","DelegateDisplayName","DelegateUPN","DelegateIdentity","Permission"' |
        Set-Content -LiteralPath $saveAt -Encoding UTF8
}
Write-Host "Exported $($rows.Count) permission entries to $saveAt"
Write-Host 'Group-based access is listed by group; this script does not expand group membership.'
