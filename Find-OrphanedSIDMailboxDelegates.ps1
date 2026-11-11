#Requires -Version 5.1
#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
Finds and optionally removes mailbox delegate entries whose unresolved identity
begins with S-1-5-21.

.DESCRIPTION
Scans Exchange Online mailboxes for Full Access, Send As, and Send on Behalf
delegate permissions. Results are displayed and exported to CSV, and progress
is shown while each mailbox is checked. After the complete scan, the script can
interactively review each result. Every removal requires its own prompt.

Microsoft Graph does not expose these Exchange mailbox delegation permissions,
so the ExchangeOnlineManagement module is required.

.PARAMETER SidPrefix
SID prefix to locate. The default is S-1-5-21.

.PARAMETER OutputPath
CSV export path. By default, the file is created in the current directory with
a timestamp in its name.

.PARAMETER UserPrincipalName
Optional admin UPN passed to Connect-ExchangeOnline.

.EXAMPLE
.\Find-OrphanedSIDMailboxDelegates.ps1

.EXAMPLE
.\Find-OrphanedSIDMailboxDelegates.ps1 -UserPrincipalName admin@contoso.com `
    -OutputPath C:\Reports\Orphaned-Mailbox-Delegates.csv
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidatePattern('^S-1-5-21')]
    [string]$SidPrefix = 'S-1-5-21',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath (
        'Orphaned-SID-Mailbox-Delegates_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    )),

    [Parameter()]
    [string]$UserPrincipalName
)

$ErrorActionPreference = 'Stop'
$connectedHere = $false
$results = [System.Collections.Generic.List[object]]::new()
$scanErrors = [System.Collections.Generic.List[string]]::new()

function Test-SidPrefix {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Identity,

        [Parameter(Mandatory)]
        [string]$Prefix
    )

    if ($null -eq $Identity) {
        return $false
    }

    $valuesToTest = [System.Collections.Generic.List[string]]::new()
    $valuesToTest.Add([string]$Identity)

    foreach ($propertyName in 'SecurityIdentifier', 'Sid', 'Name', 'DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress') {
        if ($Identity.PSObject.Properties.Name -contains $propertyName) {
            $propertyValue = $Identity.$propertyName
            if ($null -ne $propertyValue) {
                $valuesToTest.Add([string]$propertyValue)
            }
        }
    }

    foreach ($value in $valuesToTest) {
        if ($value.Trim() -like "$Prefix*") {
            return $true
        }
    }

    return $false
}

function Add-DelegateResult {
    param(
        [Parameter(Mandatory)] [object]$Mailbox,
        [Parameter(Mandatory)] [string]$PermissionType,
        [Parameter(Mandatory)] [object]$DelegateIdentity,
        [Parameter()] [string]$AccessRights,
        [Parameter()] [object]$IsInherited,
        [Parameter()] [object]$Deny,
        [Parameter(Mandatory)] [object]$RemovalTarget,
        [Parameter(Mandatory)] [string]$MailboxIdentity
    )

    $results.Add([pscustomobject][ordered]@{
        MailboxDisplayName    = [string]$Mailbox.DisplayName
        MailboxPrimarySmtp   = [string]$Mailbox.PrimarySmtpAddress
        MailboxUPN           = [string]$Mailbox.UserPrincipalName
        MailboxType          = [string]$Mailbox.RecipientTypeDetails
        PermissionType       = $PermissionType
        DelegateIdentity     = [string]$DelegateIdentity
        AccessRights         = $AccessRights
        IsInherited          = $IsInherited
        Deny                 = $Deny
        RemovalDecision      = 'Not reviewed'
        RemovalResult        = $null
        RemovalTarget        = $RemovalTarget
        MailboxIdentity      = $MailboxIdentity
    })
}

function Export-DelegateResults {
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.IEnumerable]$InputResults,
        [Parameter(Mandatory)] [string]$Path
    )

    $columns = @(
        'MailboxDisplayName', 'MailboxPrimarySmtp', 'MailboxUPN', 'MailboxType',
        'PermissionType', 'DelegateIdentity', 'AccessRights', 'IsInherited', 'Deny',
        'RemovalDecision', 'RemovalResult'
    )

    $resultArray = @($InputResults)
    if ($resultArray.Count -gt 0) {
        $resultArray |
            Sort-Object MailboxPrimarySmtp, PermissionType, DelegateIdentity |
            Select-Object $columns |
            Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    }
    else {
        $headerOnlyRow = [pscustomobject][ordered]@{}
        foreach ($column in $columns) {
            $headerOnlyRow | Add-Member -NotePropertyName $column -NotePropertyValue $null
        }
        $csvLines = $headerOnlyRow | Select-Object $columns | ConvertTo-Csv -NoTypeInformation
        $csvLines[0] | Set-Content -LiteralPath $Path -Encoding UTF8
    }
}

try {
    $existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' } |
        Select-Object -First 1

    if (-not $existingConnection) {
        Write-Host 'Connecting to Exchange Online...' -ForegroundColor Cyan
        if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            Connect-ExchangeOnline -ShowBanner:$false
        }
        else {
            Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false
        }
        $connectedHere = $true
    }

    Write-Host 'Retrieving Exchange Online mailboxes...' -ForegroundColor Cyan
    $mailboxes = @(
        Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum |
            Sort-Object DisplayName
    )

    if ($mailboxes.Count -eq 0) {
        Write-Warning 'No mailboxes were returned by Exchange Online.'
    }

    for ($index = 0; $index -lt $mailboxes.Count; $index++) {
        $mailbox = $mailboxes[$index]
        $mailboxIdentity = if ($mailbox.ExternalDirectoryObjectId) {
            [string]$mailbox.ExternalDirectoryObjectId
        }
        else {
            [string]$mailbox.PrimarySmtpAddress
        }

        $percentComplete = if ($mailboxes.Count -gt 0) {
            [math]::Floor((($index + 1) / $mailboxes.Count) * 100)
        }
        else { 100 }

        Write-Progress -Activity 'Scanning mailbox delegate permissions' `
            -Status ('Mailbox {0} of {1}: {2}' -f ($index + 1), $mailboxes.Count, $mailbox.PrimarySmtpAddress) `
            -PercentComplete $percentComplete

        # Full Access permissions
        try {
            $fullAccessPermissions = @(
                Get-MailboxPermission -Identity $mailboxIdentity -ErrorAction Stop |
                    Where-Object {
                        -not $_.IsInherited -and
                        $_.AccessRights -contains 'FullAccess'
                    }
            )

            foreach ($permission in $fullAccessPermissions) {
                if (Test-SidPrefix -Identity $permission.User -Prefix $SidPrefix) {
                    Add-DelegateResult -Mailbox $mailbox -PermissionType 'Full Access' `
                        -DelegateIdentity $permission.User `
                        -AccessRights ($permission.AccessRights -join ';') `
                        -IsInherited $permission.IsInherited -Deny $permission.Deny `
                        -RemovalTarget $permission.User -MailboxIdentity $mailboxIdentity
                }
            }
        }
        catch {
            $scanErrors.Add(('Full Access scan failed for {0}: {1}' -f $mailbox.PrimarySmtpAddress, $_.Exception.Message))
        }

        # Send As permissions
        try {
            $sendAsPermissions = @(
                Get-RecipientPermission -Identity $mailboxIdentity -ErrorAction Stop |
                    Where-Object {
                        -not $_.IsInherited -and
                        $_.AccessRights -contains 'SendAs'
                    }
            )

            foreach ($permission in $sendAsPermissions) {
                if (Test-SidPrefix -Identity $permission.Trustee -Prefix $SidPrefix) {
                    Add-DelegateResult -Mailbox $mailbox -PermissionType 'Send As' `
                        -DelegateIdentity $permission.Trustee `
                        -AccessRights ($permission.AccessRights -join ';') `
                        -IsInherited $permission.IsInherited -Deny $permission.Deny `
                        -RemovalTarget $permission.Trustee -MailboxIdentity $mailboxIdentity
                }
            }
        }
        catch {
            $scanErrors.Add(('Send As scan failed for {0}: {1}' -f $mailbox.PrimarySmtpAddress, $_.Exception.Message))
        }

        # Send on Behalf permissions
        try {
            $mailboxDetails = Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
            foreach ($delegate in @($mailboxDetails.GrantSendOnBehalfTo)) {
                if (Test-SidPrefix -Identity $delegate -Prefix $SidPrefix) {
                    Add-DelegateResult -Mailbox $mailbox -PermissionType 'Send on Behalf' `
                        -DelegateIdentity $delegate -AccessRights 'SendOnBehalf' `
                        -IsInherited $false -Deny $false `
                        -RemovalTarget $delegate -MailboxIdentity $mailboxIdentity
                }
            }
        }
        catch {
            $scanErrors.Add(('Send on Behalf scan failed for {0}: {1}' -f $mailbox.PrimarySmtpAddress, $_.Exception.Message))
        }
    }

    Write-Progress -Activity 'Scanning mailbox delegate permissions' -Completed

    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }

    if ($results.Count -gt 0) {
        Write-Host "`nFound $($results.Count) matching delegate entr$(if ($results.Count -eq 1) { 'y' } else { 'ies' })." -ForegroundColor Yellow
        $results |
            Sort-Object MailboxPrimarySmtp, PermissionType, DelegateIdentity |
            Format-Table MailboxPrimarySmtp, PermissionType, DelegateIdentity -AutoSize

        # Write the discovery report before asking about any changes.
        Export-DelegateResults -InputResults $results -Path $OutputPath
        Write-Host "`nInitial CSV exported to: $OutputPath" -ForegroundColor Green

        $reviewResponse = Read-Host "`nReview each result for optional removal? [Y/N] (default N)"
        if ($reviewResponse -match '^(y|yes)$') {
            $orderedResults = @($results | Sort-Object MailboxPrimarySmtp, PermissionType, DelegateIdentity)
            $quitReview = $false

            for ($resultIndex = 0; $resultIndex -lt $orderedResults.Count -and -not $quitReview; $resultIndex++) {
                $result = $orderedResults[$resultIndex]

                Write-Host "`n------------------------------------------------------------" -ForegroundColor DarkGray
                Write-Host ('Result {0} of {1}' -f ($resultIndex + 1), $orderedResults.Count) -ForegroundColor Cyan
                Write-Host ('Mailbox:        {0} ({1})' -f $result.MailboxDisplayName, $result.MailboxPrimarySmtp)
                Write-Host ('Mailbox type:   {0}' -f $result.MailboxType)
                Write-Host ('Permission:     {0}' -f $result.PermissionType)
                Write-Host ('Orphaned SID:   {0}' -f $result.DelegateIdentity) -ForegroundColor Yellow

                do {
                    $choice = (Read-Host 'Choose [R]emove, [S]kip, or [Q]uit review (default S)').Trim().ToUpperInvariant()
                    if ([string]::IsNullOrWhiteSpace($choice)) {
                        $choice = 'S'
                    }
                } until ($choice -in @('R', 'S', 'Q'))

                switch ($choice) {
                    'S' {
                        $result.RemovalDecision = 'Skipped'
                        $result.RemovalResult = 'No change made'
                        Write-Host 'Skipped. No change was made.' -ForegroundColor DarkYellow
                    }
                    'Q' {
                        $result.RemovalDecision = 'Review stopped'
                        $result.RemovalResult = 'No change made'
                        $quitReview = $true
                        Write-Host 'Review stopped. Remaining entries were not changed.' -ForegroundColor DarkYellow
                    }
                    'R' {
                        $result.RemovalDecision = 'Remove requested'
                        try {
                            switch ($result.PermissionType) {
                                'Full Access' {
                                    $removeParameters = @{
                                        Identity        = $result.MailboxIdentity
                                        User            = $result.RemovalTarget
                                        AccessRights    = 'FullAccess'
                                        InheritanceType = 'All'
                                        Confirm         = $false
                                        ErrorAction     = 'Stop'
                                    }
                                    if ($result.Deny -eq $true) {
                                        $removeParameters.Deny = $true
                                    }
                                    Remove-MailboxPermission @removeParameters | Out-Null
                                }
                                'Send As' {
                                    Remove-RecipientPermission -Identity $result.MailboxIdentity `
                                        -Trustee $result.RemovalTarget -AccessRights SendAs `
                                        -Confirm:$false -ErrorAction Stop | Out-Null
                                }
                                'Send on Behalf' {
                                    Set-Mailbox -Identity $result.MailboxIdentity `
                                        -GrantSendOnBehalfTo @{ Remove = $result.RemovalTarget } `
                                        -Confirm:$false -ErrorAction Stop
                                }
                                default {
                                    throw "Unsupported permission type: $($result.PermissionType)"
                                }
                            }

                            $result.RemovalResult = 'Removed successfully'
                            Write-Host 'Permission removed successfully.' -ForegroundColor Green
                        }
                        catch {
                            $result.RemovalResult = 'Removal failed: {0}' -f $_.Exception.Message
                            Write-Warning $result.RemovalResult
                        }
                    }
                }

                # Refresh the CSV after every decision so the audit trail survives an interruption.
                Export-DelegateResults -InputResults $results -Path $OutputPath
            }
        }
        else {
            Write-Host 'Removal review was not started. No changes were made.' -ForegroundColor Cyan
        }
    }
    else {
        Export-DelegateResults -InputResults @() -Path $OutputPath
        Write-Host "`nNo delegate identities beginning with '$SidPrefix' were found." -ForegroundColor Green
    }

    Write-Host "`nFinal CSV available at: $OutputPath" -ForegroundColor Green

    if ($scanErrors.Count -gt 0) {
        Write-Warning "$($scanErrors.Count) mailbox permission check(s) encountered an error."
        $errorPath = [System.IO.Path]::ChangeExtension($OutputPath, '.errors.txt')
        $scanErrors | Set-Content -LiteralPath $errorPath -Encoding UTF8
        Write-Warning "Error details exported to: $errorPath"
    }
}
finally {
    Write-Progress -Activity 'Scanning mailbox delegate permissions' -Completed
    if ($connectedHere) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}
