#requires -Version 5.1
<#
Run in a fresh Windows PowerShell session:
  powershell.exe -NoProfile -STA -File .\Export-EntraUserGroups-GUI.ps1
Install once: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Read-only. Requires consent for User.Read.All and Group.Read.All.
Optional hidden membership requires Member.Read.Hidden and an appropriate Entra role.
Only groups visible in Entra are reported; AD-only groups are outside this report.
Location identifies current directory source, not a specific domain controller.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This GUI requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Launch with powershell.exe -NoProfile -STA -File .\Export-EntraUserGroups-GUI.ps1' }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    [Windows.Forms.MessageBox]::Show('Install the module first in PowerShell: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser','Module required') | Out-Null
    return
}
Import-Module Microsoft.Graph.Authentication
$script:rows = @()
$script:busy = $false
$script:lastPath = ''
function Get-GraphCollection([string]$Uri) {
    do {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable
        foreach ($item in $page['value']) { $item }
        $Uri = $page['@odata.nextLink']
    } while ($Uri)
}
function Set-Status([string]$Text) {
    $status.Text = $Text
    $form.Refresh()
    [Windows.Forms.Application]::DoEvents()
}
function Export-Report {
    $dialog = New-Object Windows.Forms.SaveFileDialog
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.DefaultExt = 'csv'
    $dialog.FileName = $script:defaultFile
    $dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    try {
        if ($dialog.ShowDialog($form) -eq 'OK') {
            if ($script:rows.Count -gt 0) {
                $script:rows | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
            } else {
                # An empty membership report still has column headers.
                $header = '"UserDisplayName","UserPrincipalName","GroupName","GroupLocation","SyncStatus","Membership","GroupType","MembershipRuleType","GroupEmail","GroupId","OnPremisesSAMAccountName","OnPremisesSecurityIdentifier","LastSyncUTC","ReportUTC"'
                Set-Content -LiteralPath $dialog.FileName -Value $header -Encoding UTF8
            }
            $script:lastPath = $dialog.FileName
            Set-Status "Exported $($script:rows.Count) groups to $($script:lastPath)"
        } else { Set-Status 'CSV export cancelled. Results remain available; click Export CSV to save.' }
    } finally { $dialog.Dispose() }
}
$form = New-Object Windows.Forms.Form
$form.Text = 'Entra User Group Report - Hybrid'
$form.Size = New-Object Drawing.Size(1220,760)
$form.MinimumSize = New-Object Drawing.Size(980,600)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object Drawing.Font('Segoe UI',10)
$label = New-Object Windows.Forms.Label
$label.Text = 'User UPN:'; $label.SetBounds(15,18,85,25)
$upn = New-Object Windows.Forms.TextBox
$upn.SetBounds(105,15,365,28)
$nested = New-Object Windows.Forms.CheckBox
$nested.Text = 'Include nested groups'; $nested.SetBounds(485,15,205,28)
$device = New-Object Windows.Forms.CheckBox
$device.Text = 'Device code sign-in'; $device.SetBounds(705,15,205,28)
$hidden = New-Object Windows.Forms.CheckBox
$hidden.Text = 'Read hidden memberships (extra permission)'; $hidden.SetBounds(15,55,385,28)
$tenantLabel = New-Object Windows.Forms.Label
$tenantLabel.Text = 'Tenant ID (optional):'; $tenantLabel.SetBounds(420,58,155,25)
$tenant = New-Object Windows.Forms.TextBox
$tenant.SetBounds(580,55,330,28)
$run = New-Object Windows.Forms.Button
$run.Text = 'Load Groups'; $run.SetBounds(930,14,125,32)
$export = New-Object Windows.Forms.Button
$export.Text = 'Export CSV'; $export.SetBounds(1065,14,125,32); $export.Enabled = $false
$note = New-Object Windows.Forms.Label
$note.Text = 'Cloud = cloud-managed. Windows AD Server = synced from AD. AD-only groups are not visible in Entra.'
$note.SetBounds(15,95,1170,28); $note.Anchor = 'Top,Left,Right'
$grid = New-Object Windows.Forms.DataGridView
$grid.SetBounds(15,132,1175,500); $grid.Anchor = 'Top,Bottom,Left,Right'
$grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = 'DisplayedCells'; $grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'; $grid.BackgroundColor = [Drawing.Color]::White
$progress = New-Object Windows.Forms.ProgressBar
$progress.SetBounds(15,645,1175,18); $progress.Anchor = 'Bottom,Left,Right'
$status = New-Object Windows.Forms.Label
$status.SetBounds(15,675,1175,40); $status.Anchor = 'Bottom,Left,Right'
$status.Text = 'Enter a UPN and click Load Groups. Sign in with your admin account when prompted.'
$form.Controls.AddRange(@($label,$upn,$nested,$device,$hidden,$tenantLabel,$tenant,$run,$export,$note,$grid,$progress,$status))
$form.AcceptButton = $run
$form.Add_FormClosing({ param($sender,$eventArgs) if ($script:busy) { $eventArgs.Cancel = $true } })
$run.Add_Click({
    if ($script:busy) { return }
    if ([string]::IsNullOrWhiteSpace($upn.Text) -or $upn.Text.Trim() -notmatch '^[^\s@]+@[^\s@]+$') {
        [Windows.Forms.MessageBox]::Show($form,'Enter a valid user UPN.','UPN required') | Out-Null; return
    }
    $script:busy = $true; $run.Enabled = $false; $export.Enabled = $false
    $upn.Enabled = $false; $nested.Enabled = $false; $device.Enabled = $false; $hidden.Enabled = $false; $tenant.Enabled = $false
    $script:rows = @(); $grid.DataSource = $null; $progress.Value = 0
    try {
        Set-Status 'Signing in to Microsoft Graph. For device code sign-in, see the PowerShell console for the code.'
        $scopes = @('User.Read.All','Group.Read.All')
        if ($hidden.Checked) { $scopes += 'Member.Read.Hidden' }
        $connect = @{Scopes=$scopes; ContextScope='Process'; NoWelcome=$true; ErrorAction='Stop'}
        if ($device.Checked) { $connect.UseDeviceAuthentication = $true }
        if ($tenant.Text.Trim()) { $connect.TenantId = $tenant.Text.Trim() }
        Connect-MgGraph @connect
        $context = Get-MgContext
        Set-Status "Connected to tenant $($context.TenantId). Looking up user..."
        $encoded = [Uri]::EscapeDataString($upn.Text.Trim())
        $user = Invoke-MgGraphRequest -Method GET -Uri ('https://graph.microsoft.com/v1.0/users/' + $encoded + '?$select=id,displayName,userPrincipalName') -OutputType Hashtable
        $base = 'https://graph.microsoft.com/v1.0/users/' + $user['id']
        Set-Status 'Reading direct group memberships...'
        $direct = @(Get-GraphCollection ($base + '/memberOf'))
        $directIds = @{}
        foreach ($entry in $direct) { if ($entry['@odata.type'] -eq '#microsoft.graph.group') { $directIds[$entry['id']] = $true } }
        $memberships = $direct
        if ($nested.Checked) {
            Set-Status 'Reading direct and nested memberships...'
            $memberships = @(Get-GraphCollection ($base + '/transitiveMemberOf'))
        }
        $groups = @($memberships | Where-Object { $_['@odata.type'] -eq '#microsoft.graph.group' } | Sort-Object { $_['id'] } -Unique)
        $reportTime = [DateTime]::UtcNow.ToString('o')
        $collected = New-Object 'System.Collections.Generic.List[object]'
        $index = 0
        foreach ($entry in $groups) {
            $index++
            Set-Status "Reading group $index of $($groups.Count)..."
            $select = 'id,displayName,mail,mailEnabled,securityEnabled,groupTypes,onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesLastSyncDateTime'
            $g = Invoke-MgGraphRequest -Method GET -Uri ('https://graph.microsoft.com/v1.0/groups/' + $entry['id'] + '?$select=' + $select) -OutputType Hashtable
            if ([string]::IsNullOrWhiteSpace($g['displayName'])) { throw "Group $($entry['id']) has no readable name. Check Group.Read.All consent." }
            $location = 'Cloud'; $syncStatus = 'Never synced from Windows AD'
            if ($g['onPremisesSyncEnabled'] -eq $true) { $location = 'Windows AD Server'; $syncStatus = 'Currently synced from Windows AD' }
            elseif ($g['onPremisesSyncEnabled'] -eq $false) { $syncStatus = 'Previously synced; no longer synced' }
            $type = 'Other'
            if ($g['groupTypes'] -contains 'Unified') { $type = 'Microsoft 365' }
            elseif ($g['securityEnabled'] -eq $true -and $g['mailEnabled'] -eq $true) { $type = 'Mail-enabled security' }
            elseif ($g['securityEnabled'] -eq $true) { $type = 'Security' }
            elseif ($g['mailEnabled'] -eq $true) { $type = 'Distribution' }
            $membership = 'Nested'; if ($directIds.ContainsKey($g['id'])) { $membership = 'Direct' }
            $ruleType = 'Assigned'; if ($g['groupTypes'] -contains 'DynamicMembership') { $ruleType = 'Dynamic' }
            $collected.Add([pscustomobject][ordered]@{
                UserDisplayName=$user['displayName']; UserPrincipalName=$user['userPrincipalName']
                GroupName=$g['displayName']; GroupLocation=$location; SyncStatus=$syncStatus
                Membership=$membership; GroupType=$type; MembershipRuleType=$ruleType
                GroupEmail=$g['mail']; GroupId=$g['id']; OnPremisesSAMAccountName=$g['onPremisesSamAccountName']
                OnPremisesSecurityIdentifier=$g['onPremisesSecurityIdentifier']; LastSyncUTC=$g['onPremisesLastSyncDateTime']; ReportUTC=$reportTime
            })
            $progress.Value = [int](100 * $index / [Math]::Max(1,$groups.Count))
        }
        $script:rows = @($collected.ToArray() | Sort-Object GroupLocation,GroupName)
        $table = New-Object System.Data.DataTable
        $columns = @('UserDisplayName','UserPrincipalName','GroupName','GroupLocation','SyncStatus','Membership','GroupType','MembershipRuleType','GroupEmail','GroupId','OnPremisesSAMAccountName','OnPremisesSecurityIdentifier','LastSyncUTC','ReportUTC')
        foreach ($column in $columns) { [void]$table.Columns.Add($column,[string]) }
        foreach ($row in $script:rows) {
            $dr = $table.NewRow()
            foreach ($column in $columns) { $dr[$column] = [string]$row.$column }
            $table.Rows.Add($dr)
        }
        $grid.DataSource = $table
        $progress.Value = 100; $export.Enabled = $true
        $safeUPN = $user['userPrincipalName'] -replace '[^a-zA-Z0-9@._-]','_'
        $script:defaultFile = 'EntraGroups_' + $safeUPN + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv'
        Set-Status "Found $($script:rows.Count) groups for $($user['displayName']). Choose where to save the CSV."
        try { Export-Report } catch {
            Set-Status 'CSV export failed. Results remain available; click Export CSV to try again.'
            [Windows.Forms.MessageBox]::Show($form,$_.Exception.Message,'Export failed') | Out-Null
        }
    } catch {
        $script:rows = @(); $grid.DataSource = $null; $export.Enabled = $false
        Set-Status 'Report failed. No partial report was exported.'
        [Windows.Forms.MessageBox]::Show($form,($_.Exception.Message + "`r`n`r`nIf authentication fails, use a fresh PowerShell window or select Device code sign-in. Check the tenant and admin consent."),'Report error') | Out-Null
    } finally {
        $script:busy = $false; $run.Enabled = $true
        $upn.Enabled = $true; $nested.Enabled = $true; $device.Enabled = $true; $hidden.Enabled = $true; $tenant.Enabled = $true
    }
})
$export.Add_Click({ try { Export-Report } catch { [Windows.Forms.MessageBox]::Show($form,$_.Exception.Message,'Export failed') | Out-Null } })
[void]$form.ShowDialog()
$form.Dispose()
