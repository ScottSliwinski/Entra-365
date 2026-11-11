#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$Domain = 'charisstudent.org',
    [string]$OutputPath = (Join-Path -Path $PWD -ChildPath 'CharisStudent-GAL-Visibility-Report.csv')
)

$ErrorActionPreference = 'Stop'

try {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "The ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        throw "The Microsoft Graph PowerShell module is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    # Load and authenticate Microsoft Graph before Exchange Online. Both modules
    # include Microsoft identity libraries, and loading Exchange first can cause
    # an InteractiveBrowserCredential/MSAL method-version conflict.
    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Users

    $requiredGraphScopes = @('User.Read.All', 'AuditLog.Read.All')
    $graphContext = Get-MgContext
    $missingGraphScopes = @(
        $requiredGraphScopes | Where-Object {
            $null -eq $graphContext -or $_ -notin $graphContext.Scopes
        }
    )

    if ($missingGraphScopes.Count -gt 0) {
        Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Cyan
        Connect-MgGraph -Scopes $requiredGraphScopes -NoWelcome
    }

    Import-Module ExchangeOnlineManagement

    if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        Write-Host 'Connecting to Exchange Online...' -ForegroundColor Cyan

        $connectExchangeCommand = Get-Command Connect-ExchangeOnline
        if (-not $connectExchangeCommand.Parameters.ContainsKey('DisableWAM')) {
            throw "Your ExchangeOnlineManagement module does not support -DisableWAM. Update it in a new PowerShell window with: Update-Module ExchangeOnlineManagement"
        }

        Connect-ExchangeOnline -ShowBanner:$false -DisableWAM
    }

    Write-Host "Retrieving Exchange users for @$Domain..." -ForegroundColor Cyan

    # Get-User is used because UserPrincipalName is a native property of this
    # Exchange object. Get-EXORecipient does not accept UserPrincipalName in
    # its -Properties parameter.
    $recipients = @(
        Get-User -ResultSize Unlimited |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) -and
            $_.UserPrincipalName.EndsWith("@$Domain", [System.StringComparison]::OrdinalIgnoreCase)
        }
    )

    Write-Host 'Retrieving authoritative GAL visibility settings...' -ForegroundColor Cyan

    $exchangeRecipients = @(
        Get-EXORecipient -ResultSize Unlimited -Properties HiddenFromAddressListsEnabled
    )

    $recipientByEntraId = @{}
    foreach ($exchangeRecipient in $exchangeRecipients) {
        $objectId = [string]$exchangeRecipient.ExternalDirectoryObjectId
        if (-not [string]::IsNullOrWhiteSpace($objectId)) {
            $recipientByEntraId[$objectId] = $exchangeRecipient
        }
    }

    Write-Host 'Retrieving account creation and sign-in dates from Microsoft Graph...' -ForegroundColor Cyan

    $graphUsers = @(
        Get-MgUser -All -Property Id,UserPrincipalName,CreatedDateTime,SignInActivity |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) -and
            $_.UserPrincipalName.EndsWith("@$Domain", [System.StringComparison]::OrdinalIgnoreCase)
        }
    )

    $graphUserById = @{}
    $graphUserByUpn = @{}
    foreach ($graphUser in $graphUsers) {
        if (-not [string]::IsNullOrWhiteSpace([string]$graphUser.Id)) {
            $graphUserById[[string]$graphUser.Id] = $graphUser
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$graphUser.UserPrincipalName)) {
            $graphUserByUpn[[string]$graphUser.UserPrincipalName] = $graphUser
        }
    }

    $reportRows = [System.Collections.Generic.List[object]]::new()
    $total = $recipients.Count

    for ($index = 0; $index -lt $total; $index++) {
        $user = $recipients[$index]

        if ($total -gt 0) {
            $percentComplete = [math]::Round((($index + 1) / $total) * 100)
            Write-Progress -Activity 'Checking Global Address List visibility' `
                -Status "$($index + 1) of $total - $($user.UserPrincipalName)" `
                -PercentComplete $percentComplete
        }

        $objectId = [string]$user.ExternalDirectoryObjectId
        $exchangeRecipient = $null
        if (-not [string]::IsNullOrWhiteSpace($objectId)) {
            $exchangeRecipient = $recipientByEntraId[$objectId]
        }

        if ($null -eq $exchangeRecipient) {
            $galVisibility = 'Unknown - Exchange recipient not found'
            $hiddenSetting = $null
        }
        else {
            $hiddenSetting = $exchangeRecipient.HiddenFromAddressListsEnabled -eq $true
            $galVisibility = if ($hiddenSetting) { 'Hidden' } else { 'Visible' }
        }

        $graphUser = $null
        if (-not [string]::IsNullOrWhiteSpace($objectId)) {
            $graphUser = $graphUserById[$objectId]
        }
        if ($null -eq $graphUser) {
            $graphUser = $graphUserByUpn[[string]$user.UserPrincipalName]
        }

        $createdDateTime = $null
        $lastSuccessfulSignIn = $null
        $daysSinceLastSignIn = $null

        if ($graphUser) {
            if ($graphUser.CreatedDateTime) {
                $createdDateTime = ([datetimeoffset]$graphUser.CreatedDateTime).ToUniversalTime()
            }

            if ($graphUser.SignInActivity.LastSuccessfulSignInDateTime) {
                $lastSuccessfulSignIn = ([datetimeoffset]$graphUser.SignInActivity.LastSuccessfulSignInDateTime).ToUniversalTime()
                $daysSinceLastSignIn = [math]::Floor(
                    ((Get-Date).ToUniversalTime() - $lastSuccessfulSignIn.UtcDateTime).TotalDays
                )
            }
        }

        $reportRows.Add([pscustomobject]@{
            DisplayName                   = $user.DisplayName
            UserPrincipalName             = $user.UserPrincipalName
            PrimarySmtpAddress            = if ($exchangeRecipient.PrimarySmtpAddress) {
                                                [string]$exchangeRecipient.PrimarySmtpAddress
                                            }
                                            elseif ($user.WindowsEmailAddress) {
                                                [string]$user.WindowsEmailAddress
                                            }
                                            else {
                                                [string]$user.UserPrincipalName
                                            }
            Alias                         = if ($exchangeRecipient) { $exchangeRecipient.Alias } else { $user.Alias }
            RecipientType                 = if ($exchangeRecipient) { $exchangeRecipient.RecipientTypeDetails } else { $user.RecipientTypeDetails }
            GALVisibility                 = $galVisibility
            HiddenFromAddressListsEnabled = $hiddenSetting
            AccountCreatedDateUTC          = if ($createdDateTime) { $createdDateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
            LastSuccessfulLoginUTC         = if ($lastSuccessfulSignIn) { $lastSuccessfulSignIn.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
            DaysSinceLastSuccessfulLogin   = $daysSinceLastSignIn
            EntraObjectId                 = $objectId
        })
    }

    Write-Progress -Activity 'Checking Global Address List visibility' -Completed

    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    if ($reportRows.Count -gt 0) {
        $reportRows |
            Sort-Object DisplayName, UserPrincipalName |
            Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    }
    else {
        '"DisplayName","UserPrincipalName","PrimarySmtpAddress","Alias","RecipientType","GALVisibility","HiddenFromAddressListsEnabled","AccountCreatedDateUTC","LastSuccessfulLoginUTC","DaysSinceLastSuccessfulLogin","EntraObjectId"' |
            Set-Content -Path $OutputPath -Encoding UTF8
    }

    Write-Host ''
    $hiddenCount = @($reportRows | Where-Object GALVisibility -eq 'Hidden').Count
    $visibleCount = @($reportRows | Where-Object GALVisibility -eq 'Visible').Count
    $unknownCount = @($reportRows | Where-Object GALVisibility -like 'Unknown*').Count
    Write-Host "Domain recipients checked: $total" -ForegroundColor White
    Write-Host "Visible recipients:        $visibleCount" -ForegroundColor Green
    Write-Host "Hidden recipients:         $hiddenCount" -ForegroundColor Yellow
    Write-Host "Unknown recipients:        $unknownCount" -ForegroundColor Magenta
    Write-Host "CSV exported to:           $OutputPath" -ForegroundColor Cyan

    if ($reportRows.Count -gt 0) {
        $reportRows |
            Sort-Object DisplayName, UserPrincipalName |
            Format-Table DisplayName, UserPrincipalName, RecipientType, GALVisibility -AutoSize
    }
}
catch {
    Write-Error "The report could not be completed. $($_.Exception.Message)"
    exit 1
}
