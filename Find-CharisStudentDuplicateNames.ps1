#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Finds duplicate display names among Charis student accounts and exports them to CSV.

.DESCRIPTION
    Retrieves Entra ID users whose UPN ends with @charisstudent.org, normalizes
    their display names, identifies names used by more than one account, displays
    the results, and exports a detailed CSV report.

.PARAMETER StudentDomain
    UPN domain to examine. The default is charisstudent.org.

.PARAMETER OutputPath
    Destination CSV path. The default is CharisStudent-Duplicate-Names.csv in
    the current directory.

.EXAMPLE
    .\Find-CharisStudentDuplicateNames.ps1

.EXAMPLE
    .\Find-CharisStudentDuplicateNames.ps1 -OutputPath C:\Reports\Duplicates.csv

.NOTES
    Required delegated Microsoft Graph permissions:
      User.Read.All
      AuditLog.Read.All
#>

[CmdletBinding()]
param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$StudentDomain = "charisstudent.org",

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath "CharisStudent-Duplicate-Names.csv")
)

$ErrorActionPreference = "Stop"

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
}
catch {
    throw "Microsoft Graph PowerShell is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
}

try {
    Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All" -NoWelcome

    Write-Host ""
    Write-Host "Retrieving @$StudentDomain accounts..." -ForegroundColor Cyan

    $Properties = @(
        "Id"
        "DisplayName"
        "GivenName"
        "Surname"
        "UserPrincipalName"
        "Mail"
        "AccountEnabled"
        "CreatedDateTime"
        "EmployeeId"
        "JobTitle"
        "SignInActivity"
    )

    $Users = Get-MgUser `
        -Filter "endsWith(userPrincipalName,'@$StudentDomain')" `
        -ConsistencyLevel eventual `
        -CountVariable UserCount `
        -All `
        -Property $Properties

    Write-Host "Found $($Users.Count) student accounts." -ForegroundColor Green
    Write-Host "Checking for duplicate display names..." -ForegroundColor Cyan

    $PreparedUsers = foreach ($User in $Users) {
        $NormalizedName = if ([string]::IsNullOrWhiteSpace($User.DisplayName)) {
            $null
        }
        else {
            ($User.DisplayName.Trim() -replace '\s+', ' ').ToLowerInvariant()
        }

        [PSCustomObject]@{
            NormalizedName     = $NormalizedName
            DisplayName        = $User.DisplayName
            GivenName          = $User.GivenName
            Surname            = $User.Surname
            UserPrincipalName  = $User.UserPrincipalName
            Mail               = $User.Mail
            AccountEnabled     = $User.AccountEnabled
            CreatedDateTime    = $User.CreatedDateTime
            EmployeeId         = $User.EmployeeId
            JobTitle           = $User.JobTitle
            LastSignInDateTime = $User.SignInActivity.LastSignInDateTime
            EntraObjectId      = $User.Id
        }
    }

    $DuplicateGroups = $PreparedUsers |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.NormalizedName) } |
        Group-Object -Property NormalizedName |
        Where-Object { $_.Count -gt 1 } |
        Sort-Object -Property @(
            @{ Expression = "Count"; Descending = $true }
            @{ Expression = "Name"; Ascending = $true }
        )

    $Report = foreach ($Group in $DuplicateGroups) {
        foreach ($User in ($Group.Group | Sort-Object -Property CreatedDateTime)) {
            [PSCustomObject]@{
                DuplicateNameCount = $Group.Count
                DisplayName        = $User.DisplayName
                GivenName          = $User.GivenName
                Surname            = $User.Surname
                UserPrincipalName  = $User.UserPrincipalName
                Mail               = $User.Mail
                AccountEnabled     = $User.AccountEnabled
                CreatedDateTime    = $User.CreatedDateTime
                EmployeeId         = $User.EmployeeId
                JobTitle           = $User.JobTitle
                LastSignInDateTime = $User.LastSignInDateTime
                EntraObjectId      = $User.EntraObjectId
            }
        }
    }

    if ($Report.Count -gt 0) {
        $OutputDirectory = Split-Path -Path $OutputPath -Parent
        if ($OutputDirectory -and -not (Test-Path -LiteralPath $OutputDirectory)) {
            New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
        }

        $Report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

        Write-Host ""
        Write-Host "Duplicate-name groups found: $($DuplicateGroups.Count)" -ForegroundColor Yellow
        Write-Host "Accounts in duplicate groups: $($Report.Count)" -ForegroundColor Yellow
        Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
        Write-Host ""

        $Report | Format-Table DuplicateNameCount, DisplayName, UserPrincipalName,
            AccountEnabled, CreatedDateTime -AutoSize
    }
    else {
        Write-Host ""
        Write-Host "No duplicate display names were found for @$StudentDomain." -ForegroundColor Green
    }
}
catch {
    Write-Error "The duplicate-name report failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if (Get-MgContext) {
        Disconnect-MgGraph | Out-Null
    }
}
