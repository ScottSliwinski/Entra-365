Import-Module Microsoft.Graph.Users

$months = 6
$cutoff = (Get-Date).AddMonths(-$months)

Connect-MgGraph -Scopes "User.Read.All"

$users = Get-MgUser -All `
    -Filter "accountEnabled eq true and userType eq 'Guest'" `
    -Property "id,displayName,userPrincipalName,userType,accountEnabled,createdDateTime,lastPasswordChangeDateTime" `
    -ConsistencyLevel eventual

$stale = foreach ($u in $users) {

    if ($u.LastPasswordChangeDateTime) {
        $pwdChanged = [datetime]$u.LastPasswordChangeDateTime

        if ($pwdChanged -lt $cutoff) {
            [pscustomobject]@{
                DisplayName    = $u.DisplayName
                UPN            = $u.UserPrincipalName
                UserType       = $u.UserType
                AccountEnabled = $u.AccountEnabled
                Created        = $u.CreatedDateTime
                LastPwdChange  = $pwdChanged
            }
        }
    }
}

$path = ".\EnabledGuestUsers_PwdNotChanged_6Months.csv"
$stale | Sort-Object LastPwdChange, DisplayName | Export-Csv $path -NoTypeInformation
$path