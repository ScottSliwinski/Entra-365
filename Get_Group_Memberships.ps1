# Replace "user@domain.com" with the UPN
Connect-MgGraph -Scopes "User.Read.All","GroupMember.Read.All"

Get-MgUserMemberOf -UserId "brianmccloud@awmi.net" |
    Where-Object {$_.AdditionalProperties['displayName']} |
    Select-Object -ExpandProperty AdditionalProperties |
    ForEach-Object { $_['displayName'] } |
    Out-File "C:\Users\scottsliwinski\OneDrive - Andrew Wommack Ministries\Desktop\UserGroups.csv"
