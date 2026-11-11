# =====================================================================
# Entra ID Account Activity & License Audit V2
# Domain: @awmcharis.com
# =====================================================================

$Domain = "@awmcharis.com"
$OutputFile = ".\AWMCharis_Account_Audit.csv"

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {

    Write-Host "Installing Microsoft.Graph..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes `
    "User.Read.All",
    "AuditLog.Read.All",
    "Directory.Read.All",
    "Organization.Read.All"

Write-Host "Retrieving tenant license information..." -ForegroundColor Cyan

$TenantSkus = Get-MgSubscribedSku -All
$SkuLookup = @{}

foreach ($Sku in $TenantSkus) {
    $SkuLookup[$Sku.SkuId.ToString()] = $Sku.SkuPartNumber
}

Write-Host "Retrieving Entra ID users..." -ForegroundColor Cyan

$AllUsers = Get-MgUser -All -Property `
    Id,
    DisplayName,
    UserPrincipalName,
    AccountEnabled,
    CreatedDateTime,
    SignInActivity,
    AssignedLicenses,
    EmployeeType,
    JobTitle,
    Department

$Users = @(
    $AllUsers | Where-Object {
        $_.UserPrincipalName -like "*$Domain"
    }
)

$Total = $Users.Count

Write-Host "Found $Total accounts ending in $Domain" -ForegroundColor Green

$Counter = 0

$Results = foreach ($User in $Users) {

    $Counter++

    if ($Total -gt 0) {

        Write-Progress `
            -Activity "Creating $Domain Account Audit" `
            -Status "Processing $Counter of $Total - $($User.UserPrincipalName)" `
            -PercentComplete (($Counter / $Total) * 100)
    }

    # ACCOUNT AGE

    $CreatedDate = $null
    $AccountAgeDays = $null

    if ($User.CreatedDateTime) {

        $CreatedDate = [datetimeoffset]$User.CreatedDateTime

        $AccountAgeDays = [math]::Floor(
            ((Get-Date).ToUniversalTime() - $CreatedDate.UtcDateTime).TotalDays
        )
    }

    # INTERACTIVE LOGIN

    $LastInteractive = $null
    $DaysSinceInteractive = $null

    if ($User.SignInActivity.LastSignInDateTime) {

        $LastInteractive =
            [datetimeoffset]$User.SignInActivity.LastSignInDateTime

        $DaysSinceInteractive = [math]::Floor(
            (
                (Get-Date).ToUniversalTime() -
                $LastInteractive.UtcDateTime
            ).TotalDays
        )
    }

    # SUCCESSFUL LOGIN

    $LastSuccessful = $null
    $DaysSinceSuccessful = $null

    if ($User.SignInActivity.LastSuccessfulSignInDateTime) {

        $LastSuccessful =
            [datetimeoffset]$User.SignInActivity.LastSuccessfulSignInDateTime

        $DaysSinceSuccessful = [math]::Floor(
            (
                (Get-Date).ToUniversalTime() -
                $LastSuccessful.UtcDateTime
            ).TotalDays
        )
    }

    # INACTIVITY

    if ($null -eq $DaysSinceSuccessful) {
        $InactivityStatus = "Never / No Successful Login Data"
    }
    elseif ($DaysSinceSuccessful -ge 180) {
        $InactivityStatus = "180+ Days"
    }
    elseif ($DaysSinceSuccessful -ge 90) {
        $InactivityStatus = "90-179 Days"
    }
    elseif ($DaysSinceSuccessful -ge 60) {
        $InactivityStatus = "60-89 Days"
    }
    elseif ($DaysSinceSuccessful -ge 30) {
        $InactivityStatus = "30-59 Days"
    }
    else {
        $InactivityStatus = "Active - Under 30 Days"
    }

    # LICENSES

    $LicenseNames = @()

    foreach ($AssignedLicense in $User.AssignedLicenses) {

        $SkuId = $AssignedLicense.SkuId.ToString()

        if ($SkuLookup.ContainsKey($SkuId)) {
            $LicenseNames += $SkuLookup[$SkuId]
        }
        else {
            $LicenseNames += $SkuId
        }
    }

    if ($LicenseNames.Count -eq 0) {
        $Licenses = "UNLICENSED"
    }
    else {
        $Licenses = ($LicenseNames | Sort-Object -Unique) -join "; "
    }

    # MANAGER

    $ManagerName = ""
    $ManagerUPN = ""

    try {

        $Manager = Get-MgUserManager `
            -UserId $User.Id `
            -ErrorAction Stop

        if ($Manager.AdditionalProperties) {

            $ManagerName =
                $Manager.AdditionalProperties["displayName"]

            $ManagerUPN =
                $Manager.AdditionalProperties["userPrincipalName"]
        }
    }
    catch {
        $ManagerName = ""
        $ManagerUPN = ""
    }

    # RESULT

    [PSCustomObject]@{

        DisplayName = $User.DisplayName

        UserPrincipalName = $User.UserPrincipalName

        AccountStatus =
            if ($User.AccountEnabled) {
                "Enabled"
            }
            else {
                "Disabled"
            }

        EmployeeType = $User.EmployeeType
        JobTitle = $User.JobTitle
        Department = $User.Department

        Manager = $ManagerName
        ManagerUPN = $ManagerUPN

        AccountCreated =
            if ($CreatedDate) {
                $CreatedDate.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
            }
            else {
                ""
            }

        AccountAgeDays =
            if ($null -ne $AccountAgeDays) {
                $AccountAgeDays
            }
            else {
                "N/A"
            }

        LastInteractiveLogin =
            if ($LastInteractive) {
                $LastInteractive.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
            }
            else {
                "Never / No Data"
            }

        DaysSinceInteractiveLogin =
            if ($null -ne $DaysSinceInteractive) {
                $DaysSinceInteractive
            }
            else {
                "N/A"
            }

        LastSuccessfulLogin =
            if ($LastSuccessful) {
                $LastSuccessful.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
            }
            else {
                "Never / No Data"
            }

        DaysSinceSuccessfulLogin =
            if ($null -ne $DaysSinceSuccessful) {
                $DaysSinceSuccessful
            }
            else {
                "N/A"
            }

        InactivityClassification = $InactivityStatus

        Microsoft365Licenses = $Licenses

        _SortDate =
            if ($LastInteractive) {
                $LastInteractive.UtcDateTime
            }
            else {
                [datetime]::MinValue
            }
    }
}

Write-Progress -Activity "Creating Account Audit" -Completed

$Results = $Results |
    Sort-Object _SortDate -Descending |
    Select-Object * -ExcludeProperty _SortDate

$Results | Export-Csv `
    -Path $OutputFile `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " AWM CHARIS ACCOUNT AUDIT V2 COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green

Write-Host ""
Write-Host "Accounts Found : $($Results.Count)"
Write-Host "Report Saved   : $OutputFile" -ForegroundColor Yellow
Write-Host ""

$Results | Format-Table `
    DisplayName,
    UserPrincipalName,
    AccountStatus,
    LastSuccessfulLogin,
    DaysSinceSuccessfulLogin,
    InactivityClassification,
    Microsoft365Licenses `
    -AutoSize