#Requires -Version 5.1
<#
Windows PowerShell 5.1, launched with -STA. Install once in a fresh console:
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Run: powershell.exe -NoProfile -STA -File .\SharePoint-StaleUsers-GUI.ps1
Requires SharePoint Administrator, site collection administrator on scanned sites,
and consent for Graph User.Read.All, Sites.Read.All and Group.Read.All. No AD or Entra writes.
Scope: all internal domains; site collection user information list; not an effective-access audit.
Missing means no current UPN/mail/SMTP alias match, NOT proof of departure/deletion.
Guests, groups, system identities, current operator and site admins are protected.
Removal affects only the displayed site, not Entra/M365 group memberships.
app@sharepoint is always skipped and blocked from removal (case-insensitive).
#>
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This GUI requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Launch with powershell.exe -NoProfile -STA -File <script path>.' }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
foreach ($module in @('Microsoft.Online.SharePoint.PowerShell','Microsoft.Graph.Authentication')) {
    if (-not (Get-Module -ListAvailable $module)) { [Windows.Forms.MessageBox]::Show("Install-Module $module -Scope CurrentUser`r`nThen restart PowerShell and launch this script.",'Missing module') | Out-Null; exit }
    Import-Module $module
}
$script:connected = $false
$script:busy = $false
$script:tenant = ''
$script:hostName = ''
$script:operator = ''
$script:index = @{}
$script:allSites = @()
$script:scanSite = ''
$script:table = New-Object System.Data.DataTable
[void]$script:table.Columns.Add('Select',[bool])
foreach ($name in @('Site','DisplayName','LoginName','EntraUPN','Status','Reason','ObjectId','Action','Error')) { [void]$script:table.Columns.Add($name,[string]) }
[void]$script:table.Columns.Add('CanRemove',[bool])
$form = New-Object Windows.Forms.Form
$form.Text = 'SharePoint - Review Stale Site Users'
$form.Size = New-Object Drawing.Size(1250,780)
$form.MinimumSize = New-Object Drawing.Size(1100,720)
$form.StartPosition = 'CenterScreen'
function Label($text,$x,$y,$w) { $c = New-Object Windows.Forms.Label; $c.Text=$text; $c.SetBounds($x,$y,$w,23); $form.Controls.Add($c); return $c }
function Button($text,$x,$y,$w) { $c = New-Object Windows.Forms.Button; $c.Text=$text; $c.SetBounds($x,$y,$w,30); $form.Controls.Add($c); return $c }
function Check($text,$x,$y,$w,$value) { $c=New-Object Windows.Forms.CheckBox; $c.Text=$text; $c.Checked=$value; $c.SetBounds($x,$y,$w,25); $form.Controls.Add($c); return $c }
$null=Label 'SharePoint admin URL:' 15 18 170
$admin=New-Object Windows.Forms.TextBox; $admin.SetBounds(190,15,620,26); $admin.Text='https://YOURTENANT-admin.sharepoint.com'; $form.Controls.Add($admin)
$connect=Button 'Connect / Load sites' 830 12 180
$null=Label 'Site collection:' 15 60 170
$site=New-Object Windows.Forms.ComboBox; $site.SetBounds(190,57,820,28); $site.DropDownStyle='DropDownList'; $site.DisplayMember='Label'; $form.Controls.Add($site)
$scan=Button 'Scan selected site' 1025 55 185
$null=Label 'Scan scope: All internal domains' 190 97 340
$flagged=Check 'Show flagged users only' 550 95 230 $true
$excludePrivate=Check 'Exclude private group sites' 800 95 300 $true
$null=Label 'Missing accounts need review. Disabled does not prove departure. Guests, groups, admins and your account are protected.' 15 130 1200
$grid=New-Object Windows.Forms.DataGridView; $grid.SetBounds(15,165,1195,435); $grid.Anchor='Top,Bottom,Left,Right'; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.AutoGenerateColumns=$false; $grid.AutoSizeColumnsMode='DisplayedCells';  $grid.SelectionMode='FullRowSelect'; $form.Controls.Add($grid)
$selectAll=Button 'Select eligible displayed' 15 615 190
$clear=Button 'Clear selection' 215 615 150
$remove=Button 'Remove selected...' 385 615 180
$export=Button 'Export full report CSV' 585 615 190
foreach ($c in @($selectAll,$clear,$remove,$export)) { $c.Anchor='Bottom,Left' }
$progress=New-Object Windows.Forms.ProgressBar; $progress.SetBounds(15,660,1195,20); $progress.Anchor='Bottom,Left,Right'; $form.Controls.Add($progress)
$status=Label 'Connect, select a site, then scan. No users are removed during scanning.' 15 690 1195; $status.Anchor='Bottom,Left,Right'
# Create columns explicitly; automatic binding may defer creation until the form opens.
foreach ($dataColumn in $script:table.Columns) {
    if ($dataColumn.ColumnName -eq 'Select') {
        $column = New-Object Windows.Forms.DataGridViewCheckBoxColumn
    } else {
        $column = New-Object Windows.Forms.DataGridViewTextBoxColumn
    }
    $column.Name = $dataColumn.ColumnName
    $column.DataPropertyName = $dataColumn.ColumnName
    $column.HeaderText = $dataColumn.ColumnName
    $column.ReadOnly = ($dataColumn.ColumnName -ne 'Select')
    $column.Visible = ($dataColumn.ColumnName -notin @('CanRemove','ObjectId'))
    [void]$grid.Columns.Add($column)
}
$grid.DataSource = $script:table.DefaultView
function Update-View {
    if ($flagged.Checked) { $script:table.DefaultView.RowFilter="Status <> 'Enabled' AND Status <> 'Excluded'" } else { $script:table.DefaultView.RowFilter='' }
}
function Update-SitePicker {
    $script:table.Clear(); $script:scanSite=''; $site.Items.Clear()
    $hidden=0
    foreach ($s in $script:allSites) {
        # Unknown group visibility is excluded rather than assumed public.
        if ($excludePrivate.Checked -and $s.Visibility -in @('Private','HiddenMembership','Unknown')) { $hidden++; continue }
        [void]$site.Items.Add($s)
    }
    if ($site.Items.Count -gt 0) { $site.SelectedIndex=0 }
    $status.Text="$($site.Items.Count) sites available; $hidden private/unknown group sites excluded. Select a site and scan."
}
function Load-SiteCatalog($sites) {
    $groups=@{}
    $status.Text='Reading Microsoft 365 group privacy...'; [Windows.Forms.Application]::DoEvents()
    $uri='https://graph.microsoft.com/v1.0/groups?$select=id,visibility&$top=999'
    do {
        $page=Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
        foreach ($g in $page.value) { $groups[[string]$g.id]=[string]$g.visibility }
        $uri=$page.'@odata.nextLink'
        [Windows.Forms.Application]::DoEvents()
    } while ($uri)
    $catalog=@(foreach ($s in $sites) {
        $gid=[string]$s.GroupId; $visibility='Not group-connected'
        if ($gid -and $gid -ne [guid]::Empty.ToString()) {
            $visibility='Unknown'
            if ($groups.ContainsKey($gid) -and $groups[$gid] -in @('Public','Private','HiddenMembership')) { $visibility=$groups[$gid] }
        } elseif ([string]$s.Template -eq 'GROUP#0') { $visibility='Unknown' }
        # Teams private/shared channel sites have separate membership scopes.
        if ([string]$s.Template -like 'TEAMCHANNEL#*') { $visibility='Private' }
        [pscustomobject]@{Label="$($s.Title) - $($s.Url) [$visibility]"; Url=[string]$s.Url; Visibility=$visibility}
    })
    $script:allSites=$catalog
    Update-SitePicker
}
function Set-Busy($value) {
    $script:busy=$value
    foreach ($c in @($connect,$scan,$site,$admin,$flagged,$excludePrivate,$selectAll,$clear,$remove,$export)) { $c.Enabled = -not $value }
    $form.UseWaitCursor=$value
    [Windows.Forms.Application]::DoEvents()
}
function Notice($text) { [Windows.Forms.MessageBox]::Show($form,$text,'SharePoint user review') | Out-Null }
function Confirm-UnmatchedEntry($text) {
    $dialog = New-Object Windows.Forms.Form
    $dialog.Text = 'Verify unmatched entry'
    $dialog.ClientSize = New-Object Drawing.Size(700,230)
    $dialog.StartPosition = 'CenterParent'
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $message = New-Object Windows.Forms.Label
    $message.Text = $text + "`r`n`r`nYes to All approves the remaining selected unmatched entries in this batch only."
    $message.SetBounds(15,15,670,155)
    $dialog.Controls.Add($message)
    $definitions = @(
        @{Text='Yes'; Result=[Windows.Forms.DialogResult]::Yes},
        @{Text='Yes to All'; Result=[Windows.Forms.DialogResult]::OK},
        @{Text='No / Skip'; Result=[Windows.Forms.DialogResult]::No},
        @{Text='Cancel batch'; Result=[Windows.Forms.DialogResult]::Cancel}
    )
    $x = 15
    foreach ($definition in $definitions) {
        $choice = New-Object Windows.Forms.Button
        $choice.Text = $definition.Text
        $choice.DialogResult = $definition.Result
        $choice.SetBounds($x,185,155,30)
        $dialog.Controls.Add($choice)
        if ($definition.Result -eq [Windows.Forms.DialogResult]::Cancel) { $dialog.CancelButton = $choice }
        $x += 170
    }
    try { return $dialog.ShowDialog($form) }
    finally { $dialog.Dispose() }
}
function Build-Index {
    $status.Text='Reading all current Entra users and SMTP aliases...'; [Windows.Forms.Application]::DoEvents()
    $newIndex=@{}
    $uri='https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,proxyAddresses,accountEnabled,userType&$top=999'
    do {
        $page=Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
        foreach ($u in $page.value) {
            $keys=@($u.userPrincipalName,$u.mail)
            foreach ($p in $u.proxyAddresses) { if ($p -match '^smtp:(.+)$') { $keys += $Matches[1] } }
            foreach ($key in ($keys | Where-Object { $_ } | Select-Object -Unique)) {
                $k=$key.ToLowerInvariant()
                if (-not $newIndex.ContainsKey($k)) { $newIndex[$k]=@{} }
                $newIndex[$k][$u.id]=$u
            }
        }
        $uri=$page.'@odata.nextLink'
        [Windows.Forms.Application]::DoEvents()
    } while ($uri)
    # Publish only after every page succeeded. No failed lookup is treated as missing.
    $script:index=$newIndex
}
function Test-BypassedIdentity([string]$login) {
    # Accept both the plain identity and SharePoint claims-prefixed forms.
    $identity = ($login.Trim() -split '\|')[-1]
    return ($identity -ieq 'app@sharepoint')
}
function Match-User($login) {
    $key=($login -replace '^i:0#\.f\|membership\|','').ToLowerInvariant()
    if ($script:index.ContainsKey($key)) { return @($script:index[$key].Values) }
    return @()
}
function Check-Tenant {
    $ctx=Get-MgContext
    if (-not $ctx -or $ctx.TenantId -ne $script:tenant) { throw 'Graph tenant changed. Reconnect and rescan.' }
    $root=Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl' -OutputType PSObject
    if (([uri]$root.webUrl).Host -ne $script:hostName) { throw 'Graph and SharePoint tenants do not match. Reconnect using the same tenant.' }
}
$connect.Add_Click({
    Set-Busy $true
    try {
        $script:connected=$false; $script:allSites=@(); $script:table.Clear(); $site.Items.Clear(); $script:scanSite=''
        $url=$admin.Text.Trim().TrimEnd('/')
        if ($url -notmatch '^https://[a-z0-9-]+-admin\.sharepoint\.com$') { throw 'Enter your actual SharePoint admin URL, such as https://contoso-admin.sharepoint.com.' }
        $script:hostName=([uri]$url).Host -replace '-admin\.sharepoint\.com$','.sharepoint.com'
        $status.Text='Sign in to Microsoft Graph in the browser...'; [Windows.Forms.Application]::DoEvents()
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-MgGraph -Scopes 'User.Read.All','Sites.Read.All','Group.Read.All' -ContextScope Process -NoWelcome
        $script:tenant=(Get-MgContext).TenantId; $script:operator=(Get-MgContext).Account
        Check-Tenant
        $status.Text='Sign in to SharePoint, then loading site collections...'; [Windows.Forms.Application]::DoEvents()
        $connectionArgs=@{Url=$url; ErrorAction='Stop'}
        if ((Get-Command Connect-SPOService).Parameters.ContainsKey('UseSystemBrowser')) { $connectionArgs.UseSystemBrowser=$true }
        Connect-SPOService @connectionArgs
        $sites=@(Get-SPOSite -Limit All | Sort-Object Title,Url)
        Load-SiteCatalog $sites
        $script:connected=$true
    } catch { Notice $_.Exception.Message; $status.Text='Connection failed. No cleanup performed.' }
    finally { Set-Busy $false }
})
$scan.Add_Click({
    Set-Busy $true
    try {
        if (-not $script:connected -or -not $site.SelectedItem) { throw 'Connect and select a site first.' }
        $script:table.Clear(); $script:scanSite=''; Check-Tenant; Build-Index
        $target=$site.SelectedItem.Url
        $status.Text='Reading site users...'; [Windows.Forms.Application]::DoEvents()
        $users=@(Get-SPOUser -Site $target -Limit All)
        $progress.Value=0; $progress.Maximum=[Math]::Max(1,$users.Count)
        foreach ($u in $users) {
            $login=[string]$u.LoginName
            if (Test-BypassedIdentity $login) {
                $progress.Value++; [Windows.Forms.Application]::DoEvents()
                continue
            }
            $identity=$login -replace '^i:0#\.f\|membership\|',''
            $matchesFound=@(Match-User $login)
            $state='Excluded'; $reason='Group, guest, or system identity.'; $eligible=$false; $id=''; $upn=''
            $human=($login -match '^i:0#\.f\|membership\|[^|]+@[^|]+$' -or $login -match '^[^|@]+@[^|@]+$') -and $identity -notmatch '#ext#'
            if ($human) {
                if ($matchesFound.Count -gt 1) { $state='Needs review'; $reason='Ambiguous UPN/mail/alias match. Removal blocked.' }
                elseif ($matchesFound.Count -eq 0) { $state='Missing - needs review'; $reason='No current UPN, mail or SMTP alias match. Renamed/deleted accounts need verification.'; $eligible=$true }
                else {
                    $m=$matchesFound[0]; $id=$m.id; $upn=$m.userPrincipalName
                    if ($m.userType -eq 'Guest') { $state='Excluded'; $reason='Guest account.' }
                    elseif ($m.accountEnabled -eq $false) { $state='Disabled'; $reason='Entra sign-in disabled. Verify departure before removal.'; $eligible=$true }
                    elseif ($m.accountEnabled -eq $true) { $state='Enabled'; $reason='Current enabled Entra account.' }
                    else { $state='Needs review'; $reason='Account status unavailable. Removal blocked.' }
                }
                if ($u.IsSiteAdmin -or $identity -eq $script:operator -or ($upn -and $upn -eq $script:operator)) { $eligible=$false; $reason += ' Protected: site administrator or current operator.' }
            }
            $r=$script:table.NewRow(); $r.Select=$false; $r.Site=$target; $r.DisplayName=[string]$u.DisplayName; $r.LoginName=$login; $r.EntraUPN=$upn; $r.Status=$state; $r.Reason=$reason; $r.ObjectId=$id; $r.CanRemove=$eligible; $r.Action='Not removed'; $r.Error=''; $script:table.Rows.Add($r)
            $progress.Value++; [Windows.Forms.Application]::DoEvents()
        }
        $script:scanSite=$target; Update-View
        $flagCount=@($script:table.Rows | Where-Object { $_.Status -notin @('Enabled','Excluded') }).Count
        $status.Text="Scan complete: $($users.Count) site identities; $flagCount flagged. No changes made."
    } catch { $script:table.Clear(); $script:scanSite=''; Notice ("Scan failed: " + $_.Exception.Message + "`r`nYour account must be a site collection administrator on this site."); $status.Text='Scan incomplete. No results available for removal.' }
    finally { Set-Busy $false }
})
$excludePrivate.Add_CheckedChanged({
    if (-not $script:busy -and $script:connected) { Update-SitePicker }
})
$flagged.Add_CheckedChanged({ Update-View })
$site.Add_SelectedIndexChanged({ if (-not $script:busy) { $script:table.Clear(); $script:scanSite=''; $status.Text='Site changed. Scan this site before cleanup.' } })
$grid.Add_CellValueChanged({ param($sender,$eventArgs) if ($eventArgs.RowIndex -ge 0 -and $eventArgs.ColumnIndex -eq 0) { $rv=$grid.Rows[$eventArgs.RowIndex].DataBoundItem; if ($rv -and -not $rv.Row.CanRemove) { $rv.Row.Select=$false } } })
$grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { [void]$grid.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })
$selectAll.Add_Click({ foreach ($v in $script:table.DefaultView) { if ($v.Row.CanRemove -and $v.Row.Action -ne 'Removed') { $v.Row.Select=$true } } })
$clear.Add_Click({ foreach ($r in $script:table.Rows) { $r.Select=$false } })
$export.Add_Click({
    try {
        if ($script:table.Rows.Count -eq 0) { throw 'Scan a site first.' }
        $dialog=New-Object Windows.Forms.SaveFileDialog; $dialog.Filter='CSV files (*.csv)|*.csv'; $dialog.FileName="SharePoint-UserReview-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
        if ($dialog.ShowDialog() -eq 'OK') {
            $script:table.Rows | Select-Object Site,DisplayName,LoginName,EntraUPN,Status,Reason,Action,Error | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
            $status.Text='Full report exported, including hidden rows and cleanup outcomes.'
        }
        $dialog.Dispose()
    } catch { Notice $_.Exception.Message }
})
$remove.Add_Click({
    [void]$grid.EndEdit()
    $chosen=@($script:table.Rows | Where-Object { $_.Select -and $_.CanRemove -and $_.Action -ne 'Removed' })
    if ($chosen.Count -eq 0) { Notice 'Select eligible entries first.'; return }
    $prompt="Remove $($chosen.Count) selected entries from this site?`r`n$script:scanSite`r`n`r`nThis removes site user entries/access. It does not remove Entra or Microsoft 365 group membership.`r`nMissing is not proof of deletion. Confirm you have verified these people have left or should lose access."
    if ([Windows.Forms.MessageBox]::Show($form,$prompt,'Confirm selected site removals','YesNo','Warning') -ne 'Yes') { return }
    Set-Busy $true
    try {
        Check-Tenant; Build-Index
        $approveAllUnmatched = $false
        $batchCancelled = $false
        foreach ($r in $chosen) {
            try {
                if (Test-BypassedIdentity ([string]$r.LoginName)) { throw 'Excluded identity app@sharepoint. Removal blocked.' }
                if ($r.Site -ne $script:scanSite -or $site.SelectedItem.Url -ne $script:scanSite) { throw 'Site selection changed. Rescan.' }
                $current=Get-SPOUser -Site $r.Site -LoginName $r.LoginName
                if (Test-BypassedIdentity ([string]$current.LoginName)) { throw 'Excluded identity app@sharepoint. Removal blocked.' }
                if ($current.IsSiteAdmin) { throw 'Now a site administrator. Removal blocked.' }
                $now=@(Match-User $r.LoginName)
                if ($r.Status -eq 'Disabled') {
                    if ($now.Count -ne 1 -or $now[0].id -ne $r.ObjectId) { throw 'Identity changed. Rescan.' }
                    # A live read immediately before removal guards against re-enabled accounts.
                    $live=Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/" + $r.ObjectId + '?$select=id,accountEnabled,userType,userPrincipalName') -OutputType PSObject
                    if ($live.accountEnabled -ne $false -or $live.userType -eq 'Guest' -or $live.userPrincipalName -eq $script:operator) { throw 'Account no longer eligible. Rescan.' }
                } elseif ($r.Status -eq 'Missing - needs review') {
                    if ($now.Count -ne 0) { throw 'Account now resolves in Entra. Rescan.' }
                    if (-not $approveAllUnmatched) {
                        $confirm="Still unmatched: $($r.LoginName)`r`nSite: $($r.Site)`r`n`r`nHave you independently verified this entry should be removed?"
                        $answer = Confirm-UnmatchedEntry $confirm
                        if ($answer -eq [Windows.Forms.DialogResult]::OK) {
                            $approveAllUnmatched = $true
                        } elseif ($answer -eq [Windows.Forms.DialogResult]::No) {
                            $r.Action='Skipped'; $r.Select=$false; continue
                        } elseif ($answer -ne [Windows.Forms.DialogResult]::Yes) {
                            $batchCancelled = $true
                            break
                        }
                    }
                } else { throw 'This status cannot be removed.' }
                Remove-SPOUser -Site $r.Site -LoginName $r.LoginName
                $r.Action='Removed'; $r.Error=''; $r.CanRemove=$false; $r.Select=$false
            } catch { $r.Action='Blocked or failed'; $r.Error=$_.Exception.Message }
            [Windows.Forms.Application]::DoEvents()
        }
        if ($batchCancelled) { $status.Text='Batch cancelled. Previously completed removals remain recorded; remaining entries were not processed.' }
        else { $status.Text='Cleanup finished. Review Action/Error columns and export the report.' }
    } catch { Notice $_.Exception.Message; $status.Text='Cleanup stopped. Review report for any completed actions.' }
    finally { Set-Busy $false }
})
$form.Add_FormClosing({ param($sender,$e) if ($script:busy) { $e.Cancel=$true } })
Update-View
[void]$form.ShowDialog()
$form.Dispose()
