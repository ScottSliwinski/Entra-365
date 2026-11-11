#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Test-IdentityMatch {
    param(
        [AllowNull()]
        [object]$PermissionIdentity,

        [Parameter(Mandatory)]
        [string[]]$UserIdentities
    )

    if ($null -eq $PermissionIdentity) {
        return $false
    }

    $value = $PermissionIdentity.ToString().Trim()
    return ($UserIdentities -contains $value)
}

try {
    if (-not $UserPrincipalName) {
        $UserPrincipalName = Read-Host 'Enter the user UPN (example: user@domain.com)'
    }

    if ([string]::IsNullOrWhiteSpace($UserPrincipalName) -or
        $UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        throw "'$UserPrincipalName' does not appear to be a valid UPN."
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "The ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Import-Module ExchangeOnlineManagement

    Write-Host 'Connecting to Exchange Online...' -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false

    Write-Host "Validating user: $UserPrincipalName" -ForegroundColor Cyan
    $targetUser = Get-User -Identity $UserPrincipalName -ErrorAction Stop

    # Exchange permission output may identify the same user by different names.
    $identityCandidates = @(
        $UserPrincipalName
        $targetUser.UserPrincipalName
        $targetUser.WindowsEmailAddress
        $targetUser.PrimarySmtpAddress
        $targetUser.ExternalDirectoryObjectId
        $targetUser.Guid
        $targetUser.Alias
        $targetUser.Name
        $targetUser.DisplayName
    ) | ForEach-Object {
        if ($_ -and -not [string]::IsNullOrWhiteSpace($_.ToString())) {
            $_.ToString().Trim()
        }
    } | Sort-Object -Unique

    Write-Host 'Loading shared mailboxes...' -ForegroundColor Cyan
    $sharedMailboxes = @(
        Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
            Sort-Object DisplayName
    )

    if ($sharedMailboxes.Count -eq 0) {
        Write-Warning 'No shared mailboxes were found.'
        return
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $total = $sharedMailboxes.Count

    for ($index = 0; $index -lt $total; $index++) {
        $mailbox = $sharedMailboxes[$index]
        $number = $index + 1
        $percent = [math]::Round(($number / $total) * 100)

        Write-Progress `
            -Activity "Checking shared mailbox access for $UserPrincipalName" `
            -Status "[$number of $total] $($mailbox.DisplayName)" `
            -PercentComplete $percent

        $fullAccess = $false
        $sendAs = $false
        $sendOnBehalf = $false

        try {
            $fullAccess = @(
                Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction Stop |
                    Where-Object {
                        -not $_.Deny -and
                        -not $_.IsInherited -and
                        $_.AccessRights -contains 'FullAccess' -and
                        (Test-IdentityMatch -PermissionIdentity $_.User -UserIdentities $identityCandidates)
                    }
            ).Count -gt 0
        }
        catch {
            Write-Warning "Could not check Full Access on '$($mailbox.DisplayName)': $($_.Exception.Message)"
        }

        try {
            $sendAs = @(
                Get-RecipientPermission -Identity $mailbox.Identity -ErrorAction Stop |
                    Where-Object {
                        -not $_.Deny -and
                        $_.AccessRights -contains 'SendAs' -and
                        (Test-IdentityMatch -PermissionIdentity $_.Trustee -UserIdentities $identityCandidates)
                    }
            ).Count -gt 0
        }
        catch {
            Write-Warning "Could not check Send As on '$($mailbox.DisplayName)': $($_.Exception.Message)"
        }

        $sendOnBehalf = @(
            $mailbox.GrantSendOnBehalfTo |
                Where-Object {
                    Test-IdentityMatch -PermissionIdentity $_ -UserIdentities $identityCandidates
                }
        ).Count -gt 0

        if ($fullAccess -or $sendAs -or $sendOnBehalf) {
            $permissionTypes = @(
                if ($fullAccess)    { 'Full Access' }
                if ($sendAs)        { 'Send As' }
                if ($sendOnBehalf)  { 'Send on Behalf' }
            ) -join '; '

            $results.Add([PSCustomObject]@{
                UserPrincipalName      = $UserPrincipalName
                SharedMailboxName      = $mailbox.DisplayName
                SharedMailboxAddress   = $mailbox.PrimarySmtpAddress.ToString()
                FullAccess             = $fullAccess
                SendAs                 = $sendAs
                SendOnBehalf           = $sendOnBehalf
                PermissionTypes        = $permissionTypes
            })
        }
    }

    Write-Progress -Activity "Checking shared mailbox access for $UserPrincipalName" -Completed

    if (-not $OutputPath) {
        $safeUserName = ($UserPrincipalName -replace '[^a-zA-Z0-9._-]', '_')
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $OutputPath = Join-Path -Path ([Environment]::GetFolderPath('Desktop')) `
            -ChildPath "SharedMailboxAccess-$safeUserName-$timestamp.csv"
    }

    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }

    if ($results.Count -gt 0) {
        $results |
            Sort-Object SharedMailboxName |
            Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

        Write-Host "`nFound access to $($results.Count) shared mailbox(es)." -ForegroundColor Green
        $results | Sort-Object SharedMailboxName | Format-Table -AutoSize
    }
    else {
        # Create a CSV with headers even when no permissions are found.
        [PSCustomObject]@{
            UserPrincipalName    = $null
            SharedMailboxName    = $null
            SharedMailboxAddress = $null
            FullAccess           = $null
            SendAs               = $null
            SendOnBehalf         = $null
            PermissionTypes      = $null
        } | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

        # Remove the empty data row while retaining the CSV header.
        $header = Get-Content -Path $OutputPath -TotalCount 1
        Set-Content -Path $OutputPath -Value $header -Encoding UTF8

        Write-Host "`nNo direct shared-mailbox permissions were found for $UserPrincipalName." -ForegroundColor Yellow
    }

    Write-Host "CSV exported to: $OutputPath" -ForegroundColor Green
}
catch {
    Write-Progress -Activity 'Checking shared mailbox access' -Completed
    Write-Error $_.Exception.Message
}
finally {
    # Disconnect only when an Exchange Online session exists.
    try {
        if (Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
    catch {
        # Do not replace the script's main result with a disconnect error.
    }
}
