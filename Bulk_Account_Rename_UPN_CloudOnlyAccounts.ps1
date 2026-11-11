#Requires -Modules Microsoft.Graph.Users

$CsvPath    = "C:\Temp\UPNChanges.csv"
$ReportPath = "C:\Temp\UPNChangeResults.csv"

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All"

$Changes = Import-Csv -Path $CsvPath
$Results = foreach ($Entry in $Changes) {

    $OldUPN = $Entry.OldUPN.Trim()
    $NewUPN = $Entry.NewUPN.Trim()

    try {
        # Find the existing account
        $User = Get-MgUser `
            -UserId $OldUPN `
            -Property Id,DisplayName,UserPrincipalName,OnPremisesSyncEnabled `
            -ErrorAction Stop

        # Synced accounts must normally be changed in on-premises AD
        if ($User.OnPremisesSyncEnabled -eq $true) {
            Write-Warning "$OldUPN is synchronized from on-premises AD. Skipping."

            [PSCustomObject]@{
                DisplayName = $User.DisplayName
                OldUPN      = $OldUPN
                NewUPN      = $NewUPN
                Status      = "Skipped"
                Details     = "Account is synchronized from on-premises AD"
            }

            continue
        }

        # Confirm the new UPN isn't already assigned
        $ExistingUser = Get-MgUser `
            -Filter "userPrincipalName eq '$NewUPN'" `
            -ConsistencyLevel eventual `
            -ErrorAction Stop

        if ($ExistingUser) {
            throw "The new UPN is already assigned to another account."
        }

        # Use the object ID because the UPN changes during the update
        Update-MgUser `
            -UserId $User.Id `
            -UserPrincipalName $NewUPN `
            -ErrorAction Stop

        Write-Host "Changed: $OldUPN -> $NewUPN" -ForegroundColor Green

        [PSCustomObject]@{
            DisplayName = $User.DisplayName
            OldUPN      = $OldUPN
            NewUPN      = $NewUPN
            Status      = "Successful"
            Details     = "UPN updated"
        }
    }
    catch {
        Write-Warning "Failed: $OldUPN - $($_.Exception.Message)"

        [PSCustomObject]@{
            DisplayName = $User.DisplayName
            OldUPN      = $OldUPN
            NewUPN      = $NewUPN
            Status      = "Failed"
            Details     = $_.Exception.Message
        }
    }
}

$Results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
$Results | Format-Table -AutoSize

Write-Host "`nResults exported to: $ReportPath" -ForegroundColor Cyan

Disconnect-MgGraph