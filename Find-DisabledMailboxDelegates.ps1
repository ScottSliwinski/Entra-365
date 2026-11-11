#Requires -Version 5.1
<#
.SYNOPSIS
  Exports disabled user delegates on user and shared mailboxes whose mailbox UPN
  ends in @awmcharis.com or @awmi.net.
.NOTES
  Install-Module ExchangeOnlineManagement, Microsoft.Graph.Authentication,
  Microsoft.Graph.Users -Scope CurrentUser
  Requires Exchange permission to read mailbox permissions and Graph User.Read.All
  (admin consent may be required). Does not modify permissions.
#>
param(
    [string]$OutputPath = (Join-Path (Get-Location) ("Disabled-Mailbox-Delegates-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss')))
)
$ErrorActionPreference = 'Stop'
foreach ($module in 'ExchangeOnlineManagement','Microsoft.Graph.Authentication','Microsoft.Graph.Users') {
    if (-not (Get-Module -ListAvailable -Name $module)) { throw "Missing module $module. Install-Module $module -Scope CurrentUser" }
}
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
# Authenticate Graph before loading Exchange Online to avoid MSAL assembly
# version conflicts between modules in the same PowerShell process.
try {
    Connect-MgGraph -Scopes 'User.Read.All' -UseDeviceCode -NoWelcome -ErrorAction Stop
} catch {
    throw "Graph sign-in failed: $($_.Exception.Message) Start a fresh PowerShell 7 session and run this script again. If the error persists, update Microsoft.Graph.Authentication and Microsoft.Graph.Users together, then restart PowerShell."
}
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

# Fetch users once so the scan does not make a Graph call for each permission.
Write-Host 'Loading Entra users and account states...'
$users = @(Get-MgUser -All -Property 'id,displayName,userPrincipalName,mail,proxyAddresses,accountEnabled')
$byKey = @{}
foreach ($user in $users) {
    $keys = @($user.Id, $user.UserPrincipalName, $user.Mail) + @($user.ProxyAddresses | ForEach-Object { $_ -replace '^smtp:', '' })
    foreach ($key in $keys) {
        if (-not [string]::IsNullOrWhiteSpace($key)) { $byKey[$key.Trim().ToLowerInvariant()] = $user }
    }
}
$recipientCache = @{}
function Resolve-DelegateUser([string]$Identity) {
    if ([string]::IsNullOrWhiteSpace($Identity)) { return $null }
    $key = $Identity.Trim().ToLowerInvariant()
    if ($byKey.ContainsKey($key)) { return $byKey[$key] }
    if ($recipientCache.ContainsKey($key)) { return $recipientCache[$key] }
    $found = $null
    try {
        $recipient = Get-EXORecipient -Identity $Identity -ErrorAction Stop
        $id = [string]$recipient.ExternalDirectoryObjectId
        if ($id -and $byKey.ContainsKey($id.ToLowerInvariant())) { $found = $byKey[$id.ToLowerInvariant()] }
        if (-not $found) {
            foreach ($address in @([string]$recipient.PrimarySmtpAddress, [string]$recipient.WindowsLiveID)) {
                if ($address -and $byKey.ContainsKey($address.ToLowerInvariant())) { $found = $byKey[$address.ToLowerInvariant()]; break }
            }
        }
    } catch { Write-Verbose "Cannot resolve delegate '$Identity': $($_.Exception.Message)" }
    $recipientCache[$key] = $found
    return $found
}

Write-Host 'Loading user and shared mailboxes...'
$mailboxes = @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox |
    Where-Object { [string]$_.UserPrincipalName -match '(?i)@(awmcharis\.com|awmi\.net)$' })
$rows = [System.Collections.Generic.List[object]]::new()
$failures = 0
for ($i = 0; $i -lt $mailboxes.Count; $i++) {
    $mailbox = $mailboxes[$i]
    $identity = [string]$mailbox.PrimarySmtpAddress
    Write-Progress -Activity 'Checking mailbox delegates' -Status "$($i + 1) of $($mailboxes.Count): $identity" -PercentComplete ([int](100 * ($i + 1) / [math]::Max(1,$mailboxes.Count)))
    $permissions = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($entry in @(Get-EXOMailboxPermission -Identity $identity -ResultSize Unlimited -ErrorAction Stop)) {
            if (-not $entry.IsInherited -and -not $entry.Deny -and @($entry.AccessRights) -contains 'FullAccess' -and
                [string]$entry.User -ne 'NT AUTHORITY\SELF') {
                $permissions.Add([pscustomobject]@{ Name = [string]$entry.User; Right = 'FullAccess' })
            }
        }
    } catch { $failures++; Write-Warning "FullAccess lookup failed for ${identity}: $($_.Exception.Message)" }
    try {
        foreach ($entry in @(Get-EXORecipientPermission -Identity $identity -ResultSize Unlimited -ErrorAction Stop)) {
            if (-not $entry.Deny -and @($entry.AccessRights) -contains 'SendAs' -and
                [string]$entry.Trustee -ne 'NT AUTHORITY\SELF') {
                $permissions.Add([pscustomobject]@{ Name = [string]$entry.Trustee; Right = 'SendAs' })
            }
        }
    } catch { $failures++; Write-Warning "SendAs lookup failed for ${identity}: $($_.Exception.Message)" }
    foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
        if ($null -ne $delegate) {
            $permissions.Add([pscustomobject]@{ Name = [string]$delegate; Right = 'SendOnBehalf' })
        }
    }
    foreach ($permission in $permissions) {
        $user = Resolve-DelegateUser $permission.Name
        if ($null -ne $user -and $user.AccountEnabled -eq $false) {
            $rows.Add([pscustomobject]@{
                MailboxDisplayName = $mailbox.DisplayName
                MailboxUPN = $mailbox.UserPrincipalName
                MailboxEmail = $identity
                MailboxType = $mailbox.RecipientTypeDetails
                DelegateDisplayName = $user.DisplayName
                DelegateUPN = $user.UserPrincipalName
                DelegateAccountEnabled = $user.AccountEnabled
                Permission = $permission.Right
            })
        }
    }
}
Write-Progress -Activity 'Checking mailbox delegates' -Completed
$columns = 'MailboxDisplayName','MailboxUPN','MailboxEmail','MailboxType','DelegateDisplayName','DelegateUPN','DelegateAccountEnabled','Permission'
# Write headers even when there are no matches.
if ($rows.Count -eq 0) { Set-Content -LiteralPath $OutputPath -Value ($columns -join ',') -Encoding UTF8 }
else { $rows | Sort-Object MailboxUPN,DelegateUPN,Permission -Unique | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8 }
Write-Host "Checked $($mailboxes.Count) mailboxes; found $($rows.Count) disabled delegate permissions. CSV: $OutputPath"
if ($failures) { Write-Warning "$failures permission lookups failed. Review warnings above; report may be incomplete." }
