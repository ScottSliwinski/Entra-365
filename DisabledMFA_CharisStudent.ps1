# Install modules if needed
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Reports -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All"

Write-Host "Getting enabled @charisstudent.org accounts..." -ForegroundColor Cyan

# Get enabled Charis Student accounts
$Users = Get-MgUser -All `
    -Property Id,DisplayName,UserPrincipalName,AccountEnabled,CreatedDateTime `
    | Where-Object {
        $_.AccountEnabled -eq $true -and
        $_.UserPrincipalName -like "*@charisstudent.org"
    }

Write-Host "Getting MFA registration information..." -ForegroundColor Cyan

# Get MFA registration details
$MFAReport = Get-MgReportAuthenticationMethodUserRegistrationDetail -All

# Build report
$Report = foreach ($User in $Users) {

    $MFA = $MFAReport | Where-Object {
        $_.Id -eq $User.Id
    }

    # Only report users who do NOT have MFA registered
    if ($MFA.IsMfaRegistered -ne $true) {

        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            AccountEnabled    = $User.AccountEnabled
            MFAStatus         = "disabled"
            MFACapable        = $MFA.IsMfaCapable
            MethodsRegistered = ($MFA.MethodsRegistered -join ", ")
            AccountCreated    = $User.CreatedDateTime
        }
    }
}

# Sort alphabetically
$Report = $Report | Sort-Object DisplayName

# Display report
$Report | Format-Table -AutoSize

# Export to CSV
$ExportPath = "$env:USERPROFILE\OneDrive - Andrew Wommack Ministries\Documents\PowerShell\Scripts\CharisStudent_MFA_Disabled.csv"

$Report | Export-Csv `
    -Path $ExportPath `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host ""
Write-Host "Report complete." -ForegroundColor Green
Write-Host "Exported to: $ExportPath" -ForegroundColor Yellow