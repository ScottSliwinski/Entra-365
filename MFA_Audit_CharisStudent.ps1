# ============================================================
# CHARIS STUDENT - MFA AUDIT REPORT
# ============================================================

# Install required modules if needed
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Reports -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes `
    "User.Read.All",
    "AuditLog.Read.All"

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " CHARIS STUDENT MFA AUDIT" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Getting enabled @charisstudent.org users..." -ForegroundColor Yellow

# Get enabled Charis Student accounts
$Users = Get-MgUser -All `
    -Property Id,DisplayName,UserPrincipalName,AccountEnabled,CreatedDateTime,SignInActivity `
    | Where-Object {
        $_.AccountEnabled -eq $true -and
        $_.UserPrincipalName -like "*@charisstudent.org"
    }

Write-Host "Found $($Users.Count) enabled accounts." -ForegroundColor Green
Write-Host ""

Write-Host "Getting MFA registration information..." -ForegroundColor Yellow

# Authentication Methods Registration Details
$MFAReport = Get-MgReportAuthenticationMethodUserRegistrationDetail -All

Write-Host "Building MFA audit report..." -ForegroundColor Yellow
Write-Host ""

$Counter = 0

$Report = foreach ($User in $Users) {

    $Counter++

    $PercentComplete = [math]::Round(
        ($Counter / $Users.Count) * 100,
        0
    )

    Write-Progress `
        -Activity "Building MFA Audit Report" `
        -Status "$Counter of $($Users.Count) - $($User.UserPrincipalName)" `
        -PercentComplete $PercentComplete

    # Match MFA registration record to Entra user
    $MFA = $MFAReport | Where-Object {
        $_.Id -eq $User.Id
    }

    # Determine readable MFA status
    if ($MFA.IsMfaRegistered -eq $true) {
        $MFAStatus = "REGISTERED"
    }
    else {
        $MFAStatus = "NOT REGISTERED"
    }

    # Determine MFA capable status
    if ($MFA.IsMfaCapable -eq $true) {
        $MFACapable = "YES"
    }
    else {
        $MFACapable = "NO"
    }

    # Registered authentication methods
    if ($MFA.MethodsRegistered) {
        $Methods = ($MFA.MethodsRegistered -join ", ")
    }
    else {
        $Methods = "None"
    }

    # Preferred MFA method
    if ($MFA.UserPreferredMethodForSecondaryAuthentication) {
        $PreferredMFA = $MFA.UserPreferredMethodForSecondaryAuthentication
    }
    else {
        $PreferredMFA = "None"
    }

    # System preferred methods
    if ($MFA.SystemPreferredAuthenticationMethods) {
        $SystemPreferred = (
            $MFA.SystemPreferredAuthenticationMethods -join ", "
        )
    }
    else {
        $SystemPreferred = "None"
    }

    # Last successful sign-in
    if ($User.SignInActivity.LastSuccessfulSignInDateTime) {
        $LastSuccessfulSignIn = `
            $User.SignInActivity.LastSuccessfulSignInDateTime.ToLocalTime()
    }
    else {
        $LastSuccessfulSignIn = "Never / No Data"
    }

    # Last interactive sign-in attempt
    if ($User.SignInActivity.LastSignInDateTime) {
        $LastInteractiveSignIn = `
            $User.SignInActivity.LastSignInDateTime.ToLocalTime()
    }
    else {
        $LastInteractiveSignIn = "Never / No Data"
    }

    [PSCustomObject]@{

        DisplayName = $User.DisplayName

        UserPrincipalName = $User.UserPrincipalName

        AccountEnabled = $User.AccountEnabled

        MFAStatus = $MFAStatus

        MFACapable = $MFACapable

        MethodsRegistered = $Methods

        PreferredMFAMethod = $PreferredMFA

        SystemPreferredMethods = $SystemPreferred

        PasswordlessCapable = $MFA.IsPasswordlessCapable

        SSPRRegistered = $MFA.IsSsprRegistered

        LastSuccessfulSignIn = $LastSuccessfulSignIn

        LastInteractiveSignIn = $LastInteractiveSignIn

        AccountCreated = $User.CreatedDateTime

        MFAReportLastUpdated = $MFA.LastUpdatedDateTime
    }
}

Write-Progress -Activity "Building MFA Audit Report" -Completed

# ============================================================
# SORT REPORT
#
# NOT REGISTERED users first
# then newest successful login
# ============================================================

$Report = $Report |
    Sort-Object `
        @{Expression = {
            if ($_.MFAStatus -eq "NOT REGISTERED") { 0 } else { 1 }
        }},
        @{Expression = {
            if ($_.LastSuccessfulSignIn -is [DateTime]) {
                $_.LastSuccessfulSignIn
            }
            else {
                [DateTime]::MinValue
            }
        }; Descending = $true},
        DisplayName


# ============================================================
# DISPLAY REPORT
# ============================================================

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " MFA AUDIT RESULTS" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

$Report |
    Format-Table `
        DisplayName,
        UserPrincipalName,
        MFAStatus,
        MFACapable,
        PreferredMFAMethod,
        LastSuccessfulSignIn,
        AccountCreated `
        -AutoSize


# ============================================================
# SUMMARY
# ============================================================

$TotalUsers = $Report.Count

$MFARegistered = (
    $Report |
    Where-Object {
        $_.MFAStatus -eq "REGISTERED"
    }
).Count

$MFANotRegistered = (
    $Report |
    Where-Object {
        $_.MFAStatus -eq "NOT REGISTERED"
    }
).Count

if ($TotalUsers -gt 0) {

    $MFAPercentage = [math]::Round(
        ($MFARegistered / $TotalUsers) * 100,
        2
    )

}
else {

    $MFAPercentage = 0

}


Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " MFA SUMMARY" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan

Write-Host "Total Enabled Users:      $TotalUsers"
Write-Host "MFA Registered:           $MFARegistered" -ForegroundColor Green
Write-Host "MFA Not Registered:       $MFANotRegistered" -ForegroundColor Red
Write-Host "MFA Registration Rate:    $MFAPercentage%" -ForegroundColor Yellow


# ============================================================
# EXPORT CSV
# ============================================================

$DateStamp = Get-Date -Format "yyyy-MM-dd_HHmm"

$ExportPath = `
    "$env:USERPROFILE\OneDrive - Andrew Wommack Ministries\Documents\PowerShell\Scripts\CharisStudent_MFA_Audit_$DateStamp.csv"

$Report |
    Export-Csv `
        -Path $ExportPath `
        -NoTypeInformation `
        -Encoding UTF8


Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
Write-Host " REPORT COMPLETE" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Green

Write-Host ""
Write-Host "CSV exported to:" -ForegroundColor Yellow
Write-Host $ExportPath -ForegroundColor White
Write-Host ""