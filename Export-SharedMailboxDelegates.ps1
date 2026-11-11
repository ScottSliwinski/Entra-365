#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
Exports direct shared mailbox delegates and their permissions to CSV.
.EXAMPLE
.\Export-SharedMailboxDelegates.ps1 -Mailbox shared@contoso.com
.NOTES
Full Access and Send As may also be granted through groups. This report lists
the directly assigned trustees; it does not expand group membership or inherited
permissions. Send on Behalf is read from the mailbox's delegate list.
#>
[CmdletBinding()]
param(
    [string]$Mailbox,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
if (-not $Mailbox) { $Mailbox = Read-Host 'Enter the shared mailbox email address or UPN' }
if ([string]::IsNullOrWhiteSpace($Mailbox)) { throw 'A mailbox identity is required.' }

Import-Module ExchangeOnlineManagement -ErrorAction Stop
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

$mb = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
if ($mb.RecipientTypeDetails -ne 'SharedMailbox') {
    throw "'$Mailbox' is not a shared mailbox."
}
if (-not $OutputPath) {
    $safeName = ([string]$mb.PrimarySmtpAddress) -replace '[^a-zA-Z0-9._@-]', '_'
    $OutputPath = Join-Path (Get-Location).Path "SharedMailboxDelegates-$safeName.csv"
}

$entries = @{}
function Add-DelegatePermission {
    param([string]$Identity, [string]$Permission)
    if ([string]::IsNullOrWhiteSpace($Identity)) { return }
    try {
        $recipient = Get-Recipient -Identity $Identity -ErrorAction Stop
        $upn = [string]$recipient.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($upn)) { $upn = [string]$recipient.PrimarySmtpAddress }
        $name = [string]$recipient.DisplayName
        $type = [string]$recipient.RecipientTypeDetails
    } catch {
        Write-Warning "Could not resolve trustee '$Identity': $($_.Exception.Message)"
        $upn = ''
        $name = $Identity
        $type = 'Unresolved'
    }
    $key = if ($upn) { $upn.ToLowerInvariant() } else { "unresolved:$($Identity.ToLowerInvariant())" }
    if (-not $entries.ContainsKey($key)) {
        $entries[$key] = [ordered]@{
            SharedMailbox = [string]$mb.PrimarySmtpAddress
            DisplayName = $name
            UPN = $upn
            Trustee = $Identity
            RecipientType = $type
            FullAccess = $false
            SendAs = $false
            SendOnBehalf = $false
        }
    }
    $entries[$key][$Permission] = $true
}

Write-Host "Checking Full Access on $($mb.PrimarySmtpAddress)..."
Get-MailboxPermission -Identity $mb.Identity -ResultSize Unlimited |
    Where-Object { -not $_.IsInherited -and -not $_.Deny -and
        $_.AccessRights -contains 'FullAccess' -and
        [string]$_.User -notmatch '^(NT AUTHORITY\\SELF|S-1-5-10)$' } |
    ForEach-Object { Add-DelegatePermission -Identity ([string]$_.User) -Permission FullAccess }

Write-Host 'Checking Send As...'
Get-RecipientPermission -Identity $mb.Identity -ResultSize Unlimited |
    Where-Object { -not $_.IsInherited -and -not $_.Deny -and
        $_.AccessRights -contains 'SendAs' -and
        [string]$_.Trustee -notmatch '^(NT AUTHORITY\\SELF|S-1-5-10)$' } |
    ForEach-Object { Add-DelegatePermission -Identity ([string]$_.Trustee) -Permission SendAs }

Write-Host 'Checking Send on Behalf...'
foreach ($delegate in @($mb.GrantSendOnBehalfTo)) {
    Add-DelegatePermission -Identity ([string]$delegate) -Permission SendOnBehalf
}

$report = @($entries.Values | ForEach-Object {
    [pscustomobject]@{
        SharedMailbox = $_.SharedMailbox
        DisplayName = $_.DisplayName
        UPN = $_.UPN
        Trustee = $_.Trustee
        RecipientType = $_.RecipientType
        FullAccess = $_.FullAccess
        SendAs = $_.SendAs
        SendOnBehalf = $_.SendOnBehalf
    }
} | Sort-Object UPN, DisplayName)

if ($report.Count -eq 0) {
    Write-Warning 'No direct delegates were found; no CSV was created.'
    return
}
$report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported $($report.Count) delegate(s) to $OutputPath" -ForegroundColor Green
