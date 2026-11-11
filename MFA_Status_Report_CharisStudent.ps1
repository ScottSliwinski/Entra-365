# ============================================================
# CHARIS STUDENT - COMPLETE MFA AUDIT REPORT
#
# Reports:
#   - Per-User MFA Status: Disabled / Enabled / Enforced
#   - MFA Registration Status
#   - MFA Capable
#   - Registered Authentication Methods
#   - Preferred MFA Method
#   - Passwordless Capability
#   - SSPR Registration
#   - Last Successful Sign-In
#   - Last Interactive Sign-In
#   - Days Since Last Login
#   - Account Creation Date
#
# Filters:
#   - Account must be ENABLED
#   - UPN must end with @charisstudent.org
# ============================================================


# ============================================================
# INSTALL REQUIRED MODULES
# ============================================================

$RequiredModules = @(
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Reports"
)

foreach ($Module in $RequiredModules) {

    if (-not (Get-Module -ListAvailable -Name $Module)) {

        Write-Host "Installing $Module..." -ForegroundColor Yellow

        Install-Module $Module `
            -Scope CurrentUser `
            -Force
    }
}


# ============================================================
# CONNECT TO MICROSOFT GRAPH
# ============================================================

Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes `
    "User.Read.All",
    "AuditLog.Read.All",
    "UserAuthenticationMethod.Read.All",
    "Policy.Read.All"

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " CHARIS STUDENT - COMPLETE MFA AUDIT" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""


# ============================================================
# GET ENABLED CHARIS STUDENT USERS
# ============================================================

Write-Host "Getting enabled @charisstudent.org users..." `
    -ForegroundColor Yellow

$Users = Get-MgUser `
    -All `
    -Property `
        Id,
        DisplayName,
        UserPrincipalName,
        AccountEnabled,
        CreatedDateTime,
        SignInActivity |
    Where-Object {

        $_.AccountEnabled -eq $true -and
        $_.UserPrincipalName -like "*@charisstudent.org"

    }

Write-Host ""
Write-Host "Found $($Users.Count) enabled accounts." `
    -ForegroundColor Green
Write-Host ""


# ============================================================
# GET MFA REGISTRATION REPORT
# ============================================================

Write-Host "Getting MFA registration information..." `
    -ForegroundColor Yellow

$MFAReport = `
    Get-MgReportAuthenticationMethodUserRegistrationDetail -All


# Create lookup table for faster matching
$MFALookup = @{}

foreach ($MFAUser in $MFAReport) {

    $MFALookup[$MFAUser.Id] = $MFAUser

}

Write-Host "MFA registration information loaded." `
    -ForegroundColor Green
Write-Host ""


# ============================================================
# BUILD REPORT
# ============================================================

Write-Host "Building complete MFA audit..." `
    -ForegroundColor Yellow
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


    # ========================================================
    # MFA REGISTRATION INFORMATION
    # ========================================================

    $MFA = $MFALookup[$User.Id]


    # MFA Registered
    if ($MFA.IsMfaRegistered -eq $true) {

        $MFARegistered = "YES"

    }
    else {

        $MFARegistered = "NO"

    }


    # MFA Capable
    if ($MFA.IsMfaCapable -eq $true) {

        $MFACapable = "YES"

    }
    else {

        $MFACapable = "NO"

    }


    # Registered Methods
    if ($MFA.MethodsRegistered) {

        $Methods = (
            $MFA.MethodsRegistered -join ", "
        )

    }
    else {

        $Methods = "None"

    }


    # Preferred MFA Method
    if ($MFA.UserPreferredMethodForSecondaryAuthentication) {

        $PreferredMFA = `
            $MFA.UserPreferredMethodForSecondaryAuthentication

    }
    else {

        $PreferredMFA = "None"

    }


    # System Preferred Methods
    if ($MFA.SystemPreferredAuthenticationMethods) {

        $SystemPreferred = (
            $MFA.SystemPreferredAuthenticationMethods -join ", "
        )

    }
    else {

        $SystemPreferred = "None"

    }


    # ========================================================
    # GET PER-USER MFA STATUS
    #
    # Disabled
    # Enabled
    # Enforced
    #
    # This uses the Microsoft Graph Beta endpoint:
    #
    # /users/{id}/authentication/requirements
    # ========================================================

    try {

        $MFARequirement = Invoke-MgGraphRequest `
            -Method GET `
            -Uri "https://graph.microsoft.com/beta/users/$($User.Id)/authentication/requirements" `
            -ErrorAction Stop

        $PerUserMFAStatus = `
            $MFARequirement.perUserMfaState

        if ([string]::IsNullOrWhiteSpace($PerUserMFAStatus)) {

            $PerUserMFAStatus = "Disabled"

        }
        else {

            # Make output pretty:
            # enforced -> Enforced

            $PerUserMFAStatus = (
                Get-Culture
            ).TextInfo.ToTitleCase(
                $PerUserMFAStatus.ToLower()
            )

        }

    }
    catch {

        $PerUserMFAStatus = "Unknown"

    }


    # ========================================================
    # LAST SUCCESSFUL SIGN-IN
    # ========================================================

    if ($User.SignInActivity.LastSuccessfulSignInDateTime) {

        $LastSuccessfulSignIn = `
            $User.SignInActivity.LastSuccessfulSignInDateTime.ToLocalTime()

        $DaysSinceLastLogin = (
            New-TimeSpan `
                -Start $LastSuccessfulSignIn `
                -End (Get-Date)
        ).Days

    }
    else {

        $LastSuccessfulSignIn = $null
        $DaysSinceLastLogin = $null

    }


    # ========================================================
    # LAST INTERACTIVE SIGN-IN
    # ========================================================

    if ($User.SignInActivity.LastSignInDateTime) {

        $LastInteractiveSignIn = `
            $User.SignInActivity.LastSignInDateTime.ToLocalTime()

    }
    else {

        $LastInteractiveSignIn = $null

    }


    # ========================================================
    # CREATE REPORT OBJECT
    # ========================================================

    [PSCustomObject]@{

        DisplayName =
            $User.DisplayName

        UserPrincipalName =
            $User.UserPrincipalName

        AccountEnabled =
            $User.AccountEnabled

        PerUserMFAStatus =
            $PerUserMFAStatus

        MFARegistered =
            $MFARegistered

        MFACapable =
            $MFACapable

        MethodsRegistered =
            $Methods

        PreferredMFAMethod =
            $PreferredMFA

        SystemPreferredMethods =
            $SystemPreferred

        PasswordlessCapable =
            $MFA.IsPasswordlessCapable

        SSPRRegistered =
            $MFA.IsSsprRegistered

        LastSuccessfulSignIn =
            $LastSuccessfulSignIn

        DaysSinceLastLogin =
            $DaysSinceLastLogin

        LastInteractiveSignIn =
            $LastInteractiveSignIn

        AccountCreated =
            $User.CreatedDateTime

        MFAReportLastUpdated =
            $MFA.LastUpdatedDateTime

    }

}

Write-Progress `
    -Activity "Building MFA Audit Report" `
    -Completed


# ============================================================
# SORT REPORT
#
# Disabled first
# Enabled second
# Enforced third
#
# Within each group:
# newest successful login first
# ============================================================

$Report = $Report |
    Sort-Object `
        @{
            Expression = {

                switch ($_.PerUserMFAStatus) {

                    "Disabled" { 1 }
                    "Enabled"  { 2 }
                    "Enforced" { 3 }
                    default    { 4 }

                }

            }
        },
        @{
            Expression = {

                if ($_.LastSuccessfulSignIn) {

                    $_.LastSuccessfulSignIn

                }
                else {

                    [DateTime]::MinValue

                }

            }

            Descending = $true
        },
        DisplayName


# ============================================================
# DISPLAY REPORT
# ============================================================

Write-Host ""
Write-Host "==================================================" `
    -ForegroundColor Cyan

Write-Host " MFA AUDIT RESULTS" `
    -ForegroundColor Cyan

Write-Host "==================================================" `
    -ForegroundColor Cyan

Write-Host ""

$Report |
    Format-Table `
        DisplayName,
        UserPrincipalName,
        PerUserMFAStatus,
        MFARegistered,
        MFACapable,
        PreferredMFAMethod,
        LastSuccessfulSignIn,
        DaysSinceLastLogin `
        -AutoSize


# ============================================================
# SUMMARY
# ============================================================

$TotalUsers = $Report.Count

$DisabledCount = (
    $Report |
    Where-Object {
        $_.PerUserMFAStatus -eq "Disabled"
    }
).Count

$EnabledCount = (
    $Report |
    Where-Object {
        $_.PerUserMFAStatus -eq "Enabled"
    }
).Count

$EnforcedCount = (
    $Report |
    Where-Object {
        $_.PerUserMFAStatus -eq "Enforced"
    }
).Count

$RegisteredCount = (
    $Report |
    Where-Object {
        $_.MFARegistered -eq "YES"
    }
).Count

$NotRegisteredCount = (
    $Report |
    Where-Object {
        $_.MFARegistered -eq "NO"
    }
).Count


# ============================================================
# DISPLAY SUMMARY
# ============================================================

Write-Host ""
Write-Host "==================================================" `
    -ForegroundColor Cyan

Write-Host " MFA SUMMARY" `
    -ForegroundColor Cyan

Write-Host "==================================================" `
    -ForegroundColor Cyan

Write-Host ""

Write-Host "Total Enabled Student Accounts:  $TotalUsers" `
    -ForegroundColor White

Write-Host ""

Write-Host "PER-USER MFA STATUS" `
    -ForegroundColor Yellow

Write-Host "----------------------------------------"

Write-Host "Disabled:                       $DisabledCount" `
    -ForegroundColor Red

Write-Host "Enabled:                        $EnabledCount" `
    -ForegroundColor Yellow

Write-Host "Enforced:                       $EnforcedCount" `
    -ForegroundColor Green

Write-Host ""

Write-Host "MFA REGISTRATION" `
    -ForegroundColor Yellow

Write-Host "----------------------------------------"

Write-Host "MFA Registered:                 $RegisteredCount" `
    -ForegroundColor Green

Write-Host "MFA NOT Registered:             $NotRegisteredCount" `
    -ForegroundColor Red


# ============================================================
# EXPORT COMPLETE CSV
# ============================================================

$DateStamp = Get-Date -Format "yyyy-MM-dd_HHmm"

$ExportPath = `
    "$env:USERPROFILE\OneDrive - Andrew Wommack Ministries\Documents\PowerShell\Scripts\CharisStudent_MFA_DISABLED_$DateStamp.csv"

$Report |
    Export-Csv `
        -Path $ExportPath `
        -NoTypeInformation `
        -Encoding UTF8


# ============================================================
# EXPORT MFA DISABLED USERS ONLY
# ============================================================

$DisabledExportPath = `
    "$env:USERPROFILE\OneDrive - Andrew Wommack Ministries\Documents\PowerShell\Scripts\CharisStudent_MFA_DISABLED_$DateStamp.csv"

$Report |
    Where-Object {
        $_.PerUserMFAStatus -eq "Disabled"
    } |
    Export-Csv `
        -Path $DisabledExportPath `
        -NoTypeInformation `
        -Encoding UTF8


# ============================================================
# COMPLETE
# ============================================================

Write-Host ""
Write-Host "==================================================" `
    -ForegroundColor Green

Write-Host " REPORT COMPLETE" `
    -ForegroundColor Green

Write-Host "==================================================" `
    -ForegroundColor Green

Write-Host ""

Write-Host "Complete MFA Audit:" `
    -ForegroundColor Yellow

Write-Host $ExportPath

Write-Host ""

Write-Host "MFA Disabled Users Only:" `
    -ForegroundColor Yellow

Write-Host $DisabledExportPath

Write-Host ""