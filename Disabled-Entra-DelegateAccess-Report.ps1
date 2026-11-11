#requires -Version 7.0

<#
.SYNOPSIS
Finds disabled Entra ID users that retain Exchange Online delegate permissions.

.DESCRIPTION
Reports explicit Full Access, Send As, and Send on Behalf permissions assigned
to disabled Entra ID users on Exchange Online mailboxes and mail-enabled groups.

The script is read-only. It does not remove or change permissions.

.PARAMETER OutputPath
CSV output path. The default creates a timestamped CSV in the current folder.

.PARAMETER IncludeInherited
Includes inherited Full Access and Send As permission entries. By default,
inherited permission entries are excluded.

.PARAMETER KeepConnectionsOpen
Does not disconnect Microsoft Graph or Exchange Online when the scan finishes.

.PARAMETER UseDeviceCode
Uses device-code authentication for both Microsoft Graph and Exchange Online.
Use this when interactive browser authentication reports a token acquisition,
WAM, window handle, or broker error.

.PARAMETER UserPrincipalNameDomains
Only disabled users whose UPN suffix matches one of these domains are audited.
Defaults to awmi.net and awmcharis.com. Do not include the @ character.

.EXAMPLE
.\Disabled-Entra-DelegateAccess-Report.ps1

.EXAMPLE
.\Disabled-Entra-DelegateAccess-Report.ps1 -OutputPath C:\Reports\DisabledDelegates.csv
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath (
        'Disabled-Entra-DelegateAccess-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    )),

    [Parameter()]
    [switch]$IncludeInherited,

    [Parameter()]
    [switch]$KeepConnectionsOpen,

    [Parameter()]
    [switch]$UseDeviceCode,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string[]]$UserPrincipalNameDomains = @('awmi.net', 'awmcharis.com')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Stage {
    param([string]$Message)
    Write-Host ('[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor Cyan
}

function Add-LookupValue {
    param(
        [hashtable]$Table,
        [AllowNull()][object]$Value,
        [Parameter(Mandatory)][object]$User
    )

    $key = [string]$Value
    if ([string]::IsNullOrWhiteSpace($key)) { return }
    $key = $key.Trim().ToLowerInvariant()

    # Keep only unambiguous matches. A null value marks a duplicate key.
    if ($Table.ContainsKey($key)) {
        if ($null -ne $Table[$key] -and $Table[$key].Id -ne $User.Id) {
            $Table[$key] = $null
        }
    }
    else {
        $Table[$key] = $User
    }
}

function Resolve-DisabledUser {
    param(
        [AllowNull()][object]$Trustee,
        [hashtable]$Lookup
    )

    if ($null -eq $Trustee) { return $null }

    $candidateKeys = [System.Collections.Generic.List[string]]::new()
    foreach ($propertyName in @(
        'ExternalDirectoryObjectId', 'PrimarySmtpAddress', 'WindowsEmailAddress',
        'UserPrincipalName', 'Guid', 'Identity', 'Name', 'Alias',
        'DistinguishedName', 'SamAccountName'
    )) {
        if ($Trustee.PSObject.Properties.Name -contains $propertyName) {
            $value = [string]$Trustee.$propertyName
            if (-not [string]::IsNullOrWhiteSpace($value)) { $candidateKeys.Add($value) }
        }
    }

    $textValue = [string]$Trustee
    if (-not [string]::IsNullOrWhiteSpace($textValue)) { $candidateKeys.Add($textValue) }

    foreach ($candidate in $candidateKeys) {
        $key = $candidate.Trim().ToLowerInvariant()
        if ($Lookup.ContainsKey($key) -and $null -ne $Lookup[$key]) {
            return $Lookup[$key]
        }

        # Permission output can use DOMAIN\samAccountName.
        if ($key.Contains('\')) {
            $shortKey = $key.Substring($key.LastIndexOf('\') + 1)
            if ($Lookup.ContainsKey($shortKey) -and $null -ne $Lookup[$shortKey]) {
                return $Lookup[$shortKey]
            }
        }
    }

    return $null
}

function Get-SafePropertyValue {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$PropertyName
    )

    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Add-ReportRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [Parameter(Mandatory)][object]$DisabledUser,
        [Parameter(Mandatory)][object]$Target,
        [Parameter(Mandatory)][string]$PermissionType,
        [AllowNull()][object]$IsInherited,
        [AllowNull()][string]$PermissionDetails
    )

    $primarySmtpAddress = Get-SafePropertyValue -InputObject $Target -PropertyName 'PrimarySmtpAddress'
    $windowsEmailAddress = Get-SafePropertyValue -InputObject $Target -PropertyName 'WindowsEmailAddress'
    $targetAddress = if ($primarySmtpAddress) {
        [string]$primarySmtpAddress
    }
    elseif ($windowsEmailAddress) {
        [string]$windowsEmailAddress
    }
    else { '' }

    $Rows.Add([pscustomobject][ordered]@{
        DisabledUserDisplayName = $DisabledUser.DisplayName
        DisabledUserUPN         = $DisabledUser.UserPrincipalName
        DisabledUserMail        = $DisabledUser.Mail
        DisabledUserObjectId    = $DisabledUser.Id
        AccountEnabled          = $DisabledUser.AccountEnabled
        TargetDisplayName       = $Target.DisplayName
        TargetEmailAddress      = $targetAddress
        TargetRecipientType     = $Target.RecipientTypeDetails
        PermissionType          = $PermissionType
        IsInherited             = $IsInherited
        PermissionDetails       = $PermissionDetails
    })
}

$connectedGraph = $false
$connectedExchange = $false
$reportRows = [System.Collections.Generic.List[object]]::new()
$scanErrors = [System.Collections.Generic.List[string]]::new()

try {
    Write-Stage 'Checking required PowerShell modules...'
    foreach ($moduleName in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'ExchangeOnlineManagement')) {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            throw "Required module '$moduleName' is not installed. Install it with: Install-Module $moduleName -Scope CurrentUser"
        }
    }

    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Users
    Import-Module ExchangeOnlineManagement

    Write-Stage 'Connecting to Microsoft Graph...'
    try {
        if ($UseDeviceCode) {
            Connect-MgGraph -Scopes 'User.Read.All' -UseDeviceAuthentication -NoWelcome
        }
        else {
            Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome
        }
        $connectedGraph = $true
        Write-Host 'Microsoft Graph connection successful.' -ForegroundColor Green
    }
    catch {
        throw @"
Microsoft Graph authentication failed.
Error: $($_.Exception.Message)

Close this PowerShell window, open a new PowerShell 7 window, and run:
  .\Disabled-Entra-DelegateAccess-Report.ps1 -UseDeviceCode
"@
    }

    Write-Stage 'Connecting to Exchange Online...'
    try {
        if ($UseDeviceCode) {
            Connect-ExchangeOnline -Device -ShowBanner:$false
        }
        else {
            # Recent ExchangeOnlineManagement versions use WAM by default.
            # DisableWAM avoids common token broker and window-handle failures.
            $connectCommand = Get-Command Connect-ExchangeOnline
            if ($connectCommand.Parameters.ContainsKey('DisableWAM')) {
                Connect-ExchangeOnline -DisableWAM -ShowBanner:$false
            }
            else {
                Connect-ExchangeOnline -ShowBanner:$false
            }
        }
        $connectedExchange = $true
        Write-Host 'Exchange Online connection successful.' -ForegroundColor Green
    }
    catch {
        throw @"
Exchange Online authentication failed.
Error: $($_.Exception.Message)

Close this PowerShell window, open a new PowerShell 7 window, and run:
  .\Disabled-Entra-DelegateAccess-Report.ps1 -UseDeviceCode

If the error continues, update the module first:
  Update-Module ExchangeOnlineManagement
"@
    }

    $normalizedDomains = @(
        $UserPrincipalNameDomains |
            ForEach-Object { $_.Trim().TrimStart('@').ToLowerInvariant() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
    if ($normalizedDomains.Count -eq 0) {
        throw 'At least one valid UPN domain must be supplied in -UserPrincipalNameDomains.'
    }

    Write-Stage ("Getting disabled Entra ID users with UPN suffix: {0}" -f ($normalizedDomains -join ', '))
    $disabledUsers = @(
        Get-MgUser -All -Filter 'accountEnabled eq false' `
            -Property 'id,displayName,userPrincipalName,mail,accountEnabled,onPremisesSamAccountName' |
            Where-Object {
                $upn = [string]$_.UserPrincipalName
                $separatorIndex = $upn.LastIndexOf('@')
                if ($separatorIndex -lt 0) { return $false }
                $upnDomain = $upn.Substring($separatorIndex + 1).ToLowerInvariant()
                return $upnDomain -in $normalizedDomains
            }
    )

    if ($disabledUsers.Count -eq 0) {
        Write-Warning ("No disabled Entra ID users were found for: {0}" -f ($normalizedDomains -join ', '))
    }
    else {
        Write-Host ("Found {0:N0} disabled users." -f $disabledUsers.Count) -ForegroundColor Green
    }

    Write-Stage 'Loading Exchange Online mailboxes and mail-enabled groups...'
    $recipients = @(
        # Get-Recipient returns the identity and GrantSendOnBehalfTo fields used
        # by this report. Get-EXORecipient does not accept several of these
        # fields in its -Properties parameter in all module/service versions.
        Get-Recipient -ResultSize Unlimited
    )

    # Create a lookup containing only disabled users and their Exchange identities.
    $disabledById = @{}
    $disabledLookup = @{}
    foreach ($user in $disabledUsers) {
        $disabledById[$user.Id.ToLowerInvariant()] = $user
        Add-LookupValue -Table $disabledLookup -Value $user.Id -User $user
        Add-LookupValue -Table $disabledLookup -Value $user.UserPrincipalName -User $user
        Add-LookupValue -Table $disabledLookup -Value $user.Mail -User $user
        Add-LookupValue -Table $disabledLookup -Value $user.OnPremisesSamAccountName -User $user
    }

    foreach ($recipient in $recipients) {
        $externalId = [string](Get-SafePropertyValue -InputObject $recipient -PropertyName 'ExternalDirectoryObjectId')
        if (-not [string]::IsNullOrWhiteSpace($externalId)) {
            $idKey = $externalId.ToLowerInvariant()
            if ($disabledById.ContainsKey($idKey)) {
                $disabledUser = $disabledById[$idKey]
                foreach ($value in @(
                    (Get-SafePropertyValue $recipient 'ExternalDirectoryObjectId'),
                    (Get-SafePropertyValue $recipient 'PrimarySmtpAddress'),
                    (Get-SafePropertyValue $recipient 'WindowsEmailAddress'),
                    (Get-SafePropertyValue $recipient 'UserPrincipalName'),
                    (Get-SafePropertyValue $recipient 'Alias'),
                    (Get-SafePropertyValue $recipient 'Guid'),
                    (Get-SafePropertyValue $recipient 'DistinguishedName'),
                    (Get-SafePropertyValue $recipient 'Name'),
                    (Get-SafePropertyValue $recipient 'DisplayName')
                )) {
                    Add-LookupValue -Table $disabledLookup -Value $value -User $disabledUser
                }
            }
        }
    }

    $mailboxTypes = @(
        'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox',
        'DiscoveryMailbox', 'SchedulingMailbox', 'TeamMailbox'
    )
    $groupTypes = @(
        'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup',
        'DynamicDistributionGroup', 'GroupMailbox'
    )
    $mailboxes = @($recipients | Where-Object { $_.RecipientTypeDetails -in $mailboxTypes })
    $permissionTargets = @(
        $recipients | Where-Object {
            $_.RecipientTypeDetails -in $mailboxTypes -or
            $_.RecipientTypeDetails -in $groupTypes
        }
    )

    Write-Stage ("Scanning Full Access permissions on {0:N0} mailboxes..." -f $mailboxes.Count)
    for ($i = 0; $i -lt $mailboxes.Count; $i++) {
        $target = $mailboxes[$i]
        $percent = if ($mailboxes.Count) { [math]::Floor((($i + 1) / $mailboxes.Count) * 100) } else { 100 }
        Write-Progress -Id 1 -Activity 'Scanning mailbox Full Access permissions' `
            -Status ("{0:N0} of {1:N0}: {2}" -f ($i + 1), $mailboxes.Count, $target.DisplayName) `
            -PercentComplete $percent

        try {
            $targetSmtp = [string](Get-SafePropertyValue $target 'PrimarySmtpAddress')
            $permissions = @(Get-EXOMailboxPermission -Identity $targetSmtp -ErrorAction Stop)
            foreach ($permission in $permissions) {
                if (-not $IncludeInherited -and $permission.IsInherited) { continue }
                if ('FullAccess' -notin @($permission.AccessRights)) { continue }
                if ($permission.Deny) { continue }

                $disabledUser = Resolve-DisabledUser -Trustee $permission.User -Lookup $disabledLookup
                if ($disabledUser) {
                    Add-ReportRow -Rows $reportRows -DisabledUser $disabledUser -Target $target `
                        -PermissionType 'Full Access' -IsInherited $permission.IsInherited `
                        -PermissionDetails ((@($permission.AccessRights) -join '; '))
                }
            }
        }
        catch {
            $scanErrors.Add("Full Access | $targetSmtp | $($_.Exception.Message)")
        }
    }
    Write-Progress -Id 1 -Activity 'Scanning mailbox Full Access permissions' -Completed

    Write-Stage ("Scanning Send As permissions on {0:N0} mailboxes and groups..." -f $permissionTargets.Count)
    for ($i = 0; $i -lt $permissionTargets.Count; $i++) {
        $target = $permissionTargets[$i]
        $percent = if ($permissionTargets.Count) { [math]::Floor((($i + 1) / $permissionTargets.Count) * 100) } else { 100 }
        Write-Progress -Id 2 -Activity 'Scanning Send As permissions' `
            -Status ("{0:N0} of {1:N0}: {2}" -f ($i + 1), $permissionTargets.Count, $target.DisplayName) `
            -PercentComplete $percent

        try {
            $targetSmtp = [string](Get-SafePropertyValue $target 'PrimarySmtpAddress')
            $permissions = @(Get-RecipientPermission -Identity $targetSmtp -ErrorAction Stop)
            foreach ($permission in $permissions) {
                if (-not $IncludeInherited -and $permission.IsInherited) { continue }
                if ('SendAs' -notin @($permission.AccessRights)) { continue }
                if ($permission.Deny) { continue }

                $disabledUser = Resolve-DisabledUser -Trustee $permission.Trustee -Lookup $disabledLookup
                if ($disabledUser) {
                    Add-ReportRow -Rows $reportRows -DisabledUser $disabledUser -Target $target `
                        -PermissionType 'Send As' -IsInherited $permission.IsInherited `
                        -PermissionDetails ((@($permission.AccessRights) -join '; '))
                }
            }
        }
        catch {
            $scanErrors.Add("Send As | $targetSmtp | $($_.Exception.Message)")
        }
    }
    Write-Progress -Id 2 -Activity 'Scanning Send As permissions' -Completed

    Write-Stage 'Checking Send on Behalf assignments...'
    for ($i = 0; $i -lt $permissionTargets.Count; $i++) {
        $target = $permissionTargets[$i]
        $percent = if ($permissionTargets.Count) { [math]::Floor((($i + 1) / $permissionTargets.Count) * 100) } else { 100 }
        Write-Progress -Id 3 -Activity 'Scanning Send on Behalf assignments' `
            -Status ("{0:N0} of {1:N0}: {2}" -f ($i + 1), $permissionTargets.Count, $target.DisplayName) `
            -PercentComplete $percent

        $sendOnBehalfDelegates = Get-SafePropertyValue $target 'GrantSendOnBehalfTo'
        foreach ($delegate in @($sendOnBehalfDelegates)) {
            $disabledUser = Resolve-DisabledUser -Trustee $delegate -Lookup $disabledLookup
            if ($disabledUser) {
                Add-ReportRow -Rows $reportRows -DisabledUser $disabledUser -Target $target `
                    -PermissionType 'Send on Behalf' -IsInherited $false `
                    -PermissionDetails ([string]$delegate)
            }
        }
    }
    Write-Progress -Id 3 -Activity 'Scanning Send on Behalf assignments' -Completed

    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    $sortedRows = @(
        $reportRows | Sort-Object DisabledUserUPN, TargetRecipientType, TargetEmailAddress, PermissionType -Unique
    )

    if ($sortedRows.Count -gt 0) {
        $sortedRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8BOM
    }
    else {
        # Export headers even when the report has no findings.
        [pscustomobject][ordered]@{
            DisabledUserDisplayName = $null; DisabledUserUPN = $null; DisabledUserMail = $null
            DisabledUserObjectId = $null; AccountEnabled = $null; TargetDisplayName = $null
            TargetEmailAddress = $null; TargetRecipientType = $null; PermissionType = $null
            IsInherited = $null; PermissionDetails = $null
        } | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8BOM
        # Remove the blank data row, retaining only the header line.
        $header = Get-Content -LiteralPath $OutputPath -First 1
        Set-Content -LiteralPath $OutputPath -Value $header -Encoding utf8BOM
    }

    Write-Host ''
    Write-Host 'Scan complete.' -ForegroundColor Green
    Write-Host ("Disabled users checked : {0:N0}" -f $disabledUsers.Count)
    Write-Host ("Mailboxes checked      : {0:N0}" -f $mailboxes.Count)
    Write-Host ("Mailbox/group targets  : {0:N0}" -f $permissionTargets.Count)
    Write-Host ("Permission findings    : {0:N0}" -f $sortedRows.Count)
    Write-Host ("Recipient scan errors  : {0:N0}" -f $scanErrors.Count)
    Write-Host ("CSV report             : {0}" -f (Resolve-Path -LiteralPath $OutputPath))

    if ($scanErrors.Count -gt 0) {
        $errorPath = [System.IO.Path]::ChangeExtension($OutputPath, '.errors.log')
        $scanErrors | Set-Content -Path $errorPath -Encoding utf8BOM
        Write-Warning "Some recipients could not be scanned. Details: $errorPath"
    }
}
finally {
    Write-Progress -Id 1 -Activity 'Scanning mailbox Full Access permissions' -Completed
    Write-Progress -Id 2 -Activity 'Scanning Send As permissions' -Completed
    Write-Progress -Id 3 -Activity 'Scanning Send on Behalf assignments' -Completed

    if (-not $KeepConnectionsOpen) {
        if ($connectedExchange) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
        if ($connectedGraph) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
}
