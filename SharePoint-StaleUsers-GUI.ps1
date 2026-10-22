#Requires -Version 5.1
<#
Windows PowerShell 5.1, launched with -STA. Install once in a fresh console:
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Run: powershell.exe -NoProfile -STA -File .\SharePoint-StaleUsers-GUI-v2-AllSites.ps1
Scan selected site or all currently listed sites; private-site filter applies to both.
All-sites means site collections returned by Get-SPOSite, not every subsite or OneDrive.
Failed sites are logged and skipped; no permissions are automatically granted.
Requires SharePoint Administrator, site collection administrator on scanned sites,
and consent for Graph User.Read.All, Sites.Read.All and Group.Read.All. No AD or Entra writes.
Scope: all internal domains; site collection user information list; not an effective-access audit.
Missing means no current UPN/mail/SMTP alias match, NOT proof of departure/deletion.
Guests, groups, system identities, current operator and site admins are protected.
Removal affects only each displayed site, not Entra/M365 group memberships.
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
$script:scannedSites = @{}
$script:scanLog = @()
$script:cancelScan = $false
$script:table = New-Object System.Data.DataTable
[void]$script:table.Columns.Add('Select',[bool])
foreach ($name in @('SiteName','Site','DisplayName','LoginName','EntraUPN','Status','Reason','ObjectId','Action','Error')) { [void]$script:table.Columns.Add($name,[string]) }
[void]$script:table.Columns.Add('CanRemove',[bool])
$form = New-Object Windows.Forms.Form
$form.Text = 'SharePoint - Stale Users v2 (Selected or All Sites)'
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
$scanAll=Button 'Scan all available sites' 1025 90 185
$stopScan=Button 'Stop scan' 1025 12 185; $stopScan.Enabled=$false
$null=Label 'Scan scope: All internal domains' 190 97 340
$flagged=Check 'Show flagged users only' 510 95 245 $true
$excludePrivate=Check 'Exclude private group sites' 765 95 245 $true
$null=Label 'Missing accounts need review. Disabled does not prove departure. Guests, groups, admins and your account are protected.' 15 130 1200
$grid=New-Object Windows.Forms.DataGridView; $grid.SetBounds(15,165,1195,435); $grid.Anchor='Top,Bottom,Left,Right'; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.AutoGenerateColumns=$false; $grid.AutoSizeColumnsMode='DisplayedCells';  $grid.SelectionMode='FullRowSelect'; $form.Controls.Add($grid)
$selectAll=Button 'Select eligible displayed' 15 615 190
$clear=Button 'Clear selection' 215 615 150
$remove=Button 'Remove selected...' 385 615 180
$export=Button 'Export full report CSV' 585 615 190
$exportLog=Button 'Export site scan log' 795 615 185
foreach ($c in @($selectAll,$clear,$remove,$export,$exportLog)) { $c.Anchor='Bottom,Left' }
$progress=New-Object Windows.Forms.ProgressBar; $progress.SetBounds(15,660,1195,20); $progress.Anchor='Bottom,Left,Right'; $form.Controls.Add($progress)
$status=Label 'Connect, then scan one site or all available sites. No users are removed during scanning.' 15 690 1195; $status.Anchor='Bottom,Left,Right'
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
    $script:table.Clear(); $script:scannedSites=@{}; $script:scanLog=@(); $site.Items.Clear()
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
        [pscustomobject]@{Label="$($s.Title) - $($s.Url) [$visibility]"; Url=[string]$s.Url; Visibility=$visibility; Title=[string]$s.Title}
    })
    $script:allSites=$catalog
    Update-SitePicker
}
function Set-Busy($value) {
    $script:busy=$value
    foreach ($c in @($connect,$scan,$scanAll,$site,$admin,$flagged,$excludePrivate,$selectAll,$clear,$remove,$export,$exportLog)) { $c.Enabled = -not $value }
    $grid.Enabled= -not $value
    $stopScan.Enabled=$false
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
        $script:connected=$false; $script:allSites=@(); $script:table.Clear(); $site.Items.Clear(); $script:scannedSites=@{}; $script:scanLog=@()
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
        $sites=@(Get-SPOSite -Limit All | Sort-Object -Property @('Title','Url'))
        Load-SiteCatalog $sites
        $script:connected=$true
    } catch { Notice $_.Exception.Message; $status.Text='Connection failed. No cleanup performed.' }
    finally { Set-Busy $false }
})
function Scan-OneSite($entry,$siteNumber,$siteTotal) {
    $target=[string]$entry.Url
    $status.Text="Site $siteNumber/$siteTotal - reading $target"
    [Windows.Forms.Application]::DoEvents()
    if ($script:cancelScan) { throw [OperationCanceledException]::new('Scan cancelled.') }
    # Get-SPOUser must finish successfully before any identity is classified.
    $users=@(Get-SPOUser -Site $target -Limit All -ErrorAction Stop)
    $userNumber=0
        foreach ($u in $users) {
            if ($script:cancelScan) { throw [OperationCanceledException]::new('Scan cancelled.') }
            $login=[string]$u.LoginName
            if (Test-BypassedIdentity $login) {
                [Windows.Forms.Application]::DoEvents()
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
            $r=$script:table.NewRow(); $r.Select=$false; $r.SiteName=[string]$entry.Title; $r.Site=$target; $r.DisplayName=[string]$u.DisplayName; $r.LoginName=$login; $r.EntraUPN=$upn; $r.Status=$state; $r.Reason=$reason; $r.ObjectId=$id; $r.CanRemove=$eligible; $r.Action='Not removed'; $r.Error=''; [void]$script:table.Rows.Add($r)
            $userNumber++
            $status.Text="Site $siteNumber/$siteTotal - $($entry.Title): user $userNumber/$($users.Count)"
            [Windows.Forms.Application]::DoEvents()
            if ($script:cancelScan) { throw [OperationCanceledException]::new('Scan cancelled.') }
        }

    return $users.Count
}
function Start-SiteScan($targets) {
    if (-not $script:connected -or @($targets).Count -eq 0) { Notice 'Connect and load at least one available site first.'; return }
    Set-Busy $true
    $script:table.Clear(); $script:scannedSites=@{}; $script:scanLog=@(); $script:cancelScan=$false
    $done=0; $failed=0; $total=@($targets).Count
    $progress.Value=0; $progress.Maximum=[Math]::Max(1,$total)
    try {
        Check-Tenant
        # One complete directory snapshot for the entire run, not once per site.
        Build-Index
        $stopScan.Enabled=$true
        foreach ($entry in $targets) {
            if ($script:cancelScan) { break }
            $before=$script:table.Rows.Count
            $siteState='Complete'; $siteError=''; $identityCount=0
            try {
                $identityCount=Scan-OneSite $entry ($done+1) $total
                $script:scannedSites[[string]$entry.Url]=$true
            } catch {
                # Discard partial results for this site; earlier successful sites remain.
                while ($script:table.Rows.Count -gt $before) { $script:table.Rows.RemoveAt($script:table.Rows.Count-1) }
                if ($script:cancelScan) { $siteState='Cancelled' }
                else { $siteState='Skipped / failed'; $failed++ }
                $siteError=$_.Exception.Message
            }
            $script:scanLog += [pscustomobject]@{SiteName=[string]$entry.Title; Site=[string]$entry.Url; Status=$siteState; Identities=$identityCount; Error=$siteError}
            $done++; $progress.Value=$done
            Update-View
            [Windows.Forms.Application]::DoEvents()
        }
        $flagCount=@($script:table.Rows | Where-Object { $_.Status -notin @('Enabled','Excluded') }).Count
        if ($script:cancelScan) {
            for ($i=$done; $i -lt $total; $i++) {
                $pending=$targets[$i]
                $script:scanLog += [pscustomobject]@{SiteName=[string]$pending.Title; Site=[string]$pending.Url; Status='Not scanned - cancelled'; Identities=0; Error=''}
            }
        }
        $state=if ($script:cancelScan) { 'CANCELLED / PARTIAL' } elseif ($failed) { 'FINISHED WITH SKIPPED SITES' } else { 'COMPLETE' }
        $status.Text="$state - $($script:scannedSites.Count)/$total sites scanned; $failed failed; $flagCount flagged. No changes made. Export site scan log for details."
    } catch {
        # Authentication/directory failures never become missing-user results.
        $script:table.Clear(); $script:scannedSites=@{}
        $script:scanLog += [pscustomobject]@{SiteName=''; Site=''; Status='Run failed'; Identities=0; Error=$_.Exception.Message}
        Notice ("Scan failed: " + $_.Exception.Message)
        $status.Text='Scan failed. No results available for removal. Export site scan log for details.'
    } finally { Update-View; Set-Busy $false }
}
$scan.Add_Click({
    if (-not $site.SelectedItem) { Notice 'Connect and select a site first.'; return }
    Start-SiteScan @($site.SelectedItem)
})
$scanAll.Add_Click({
    # Snapshot the filtered picker so hidden/private sites cannot leak into this run.
    $targets=@($site.Items | ForEach-Object { $_ })
    Start-SiteScan $targets
})
$stopScan.Add_Click({ $script:cancelScan=$true; $stopScan.Enabled=$false; $status.Text='Stopping scan after the current request returns...' })
$excludePrivate.Add_CheckedChanged({
    if (-not $script:busy -and $script:connected) { Update-SitePicker }
})
$flagged.Add_CheckedChanged({ Update-View })
$site.Add_SelectedIndexChanged({ if (-not $script:busy) { $script:table.Clear(); $script:scannedSites=@{}; $script:scanLog=@(); $status.Text='Site changed. Scan this site before cleanup.' } })
$grid.Add_CellValueChanged({ param($sender,$eventArgs) if ($eventArgs.RowIndex -ge 0 -and $eventArgs.ColumnIndex -eq 0) { $rv=$grid.Rows[$eventArgs.RowIndex].DataBoundItem; if ($rv -and -not $rv.Row.CanRemove) { $rv.Row.Select=$false } } })
$grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { [void]$grid.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })
$selectAll.Add_Click({ foreach ($v in $script:table.DefaultView) { if ($v.Row.CanRemove -and $v.Row.Action -ne 'Removed') { $v.Row.Select=$true } } })
$clear.Add_Click({ foreach ($r in $script:table.Rows) { $r.Select=$false } })
$export.Add_Click({
    try {
        if ($script:table.Rows.Count -eq 0) { throw 'Scan a site first.' }
        $dialog=New-Object Windows.Forms.SaveFileDialog; $dialog.Filter='CSV files (*.csv)|*.csv'; $dialog.FileName="SharePoint-UserReview-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
        if ($dialog.ShowDialog() -eq 'OK') {
            $script:table.Rows | Select-Object -Property @('SiteName','Site','DisplayName','LoginName','EntraUPN','Status','Reason','Action','Error') | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
            $status.Text='Full report exported, including hidden rows and cleanup outcomes.'
        }
        $dialog.Dispose()
    } catch { Notice $_.Exception.Message }
})
$exportLog.Add_Click({
    if ($script:scanLog.Count -eq 0) { Notice 'Run a scan first.'; return }
    $dialog=New-Object Windows.Forms.SaveFileDialog
    $dialog.Filter='CSV files (*.csv)|*.csv'; $dialog.FileName="SharePoint-SiteScanLog-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
    try {
        if ($dialog.ShowDialog() -eq 'OK') {
            $script:scanLog | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8
            $status.Text='Site scan log exported.'
        }
    } catch { Notice $_.Exception.Message }
    finally { $dialog.Dispose() }
})
$remove.Add_Click({
    [void]$grid.EndEdit()
    $chosen=@($script:table.Rows | Where-Object { $_.Select -and $_.CanRemove -and $_.Action -ne 'Removed' })
    if ($chosen.Count -eq 0) { Notice 'Select eligible entries first.'; return }
    $siteCount=@($chosen | Select-Object -ExpandProperty Site -Unique).Count
    $prompt="Remove $($chosen.Count) selected entries across $siteCount site(s)?`r`n`r`nEach row will be removed ONLY from the Site URL shown on that row. This removes site user entries/access. It does not remove Entra or Microsoft 365 group membership.`r`nMissing is not proof of deletion. Confirm you have verified these people have left or should lose access."
    if ([Windows.Forms.MessageBox]::Show($form,$prompt,'Confirm selected site removals','YesNo','Warning') -ne 'Yes') { return }
    Set-Busy $true
    try {
        Check-Tenant; Build-Index
        $approveAllUnmatched = $false
        $batchCancelled = $false
        foreach ($r in $chosen) {
            try {
                if (Test-BypassedIdentity ([string]$r.LoginName)) { throw 'Excluded identity app@sharepoint. Removal blocked.' }
                if (-not $script:scannedSites.ContainsKey([string]$r.Site)) { throw 'Site was not successfully scanned. Rescan.' }
                $current=Get-SPOUser -Site $r.Site -LoginName $r.LoginName
                if (Test-BypassedIdentity ([string]$current.LoginName)) { throw 'Excluded identity app@sharepoint. Removal blocked.' }
                if ($current.IsSiteAdmin) { throw 'Site administrator. Removal blocked.' }
                if (($r.LoginName -replace '^i:0#\.f\|membership\|','') -eq $script:operator) { throw 'Current operator is protected.' }
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
