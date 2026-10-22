#Requires -Version 5.1
<#
Hybrid group membership copier. Run Windows PowerShell 5.1 in STA on a domain
joined PC with RSAT ActiveDirectory. No modules are installed automatically.
Optional cloud support requires Microsoft.Graph.Authentication.
Both AD users must be in the selected AD domain. Different UPN suffixes are OK.
Direct memberships only; preserves primary group. Does not copy ownership,
mailbox permissions, directory roles, individual licenses, or user attributes.
Additions run before removals. If any addition fails, all removals are withheld.
No automatic rollback. CSV snapshots and results are written before/after changes.
Cloud dynamic, role-assignable, synced and mail-enabled non-M365 groups are skipped.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This GUI requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Start with: powershell.exe -NoProfile -STA -File .\Copy-HybridUserGroups-GUI.ps1'
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# All AD calls below use an explicit server; do not initialize a default AD: drive.
$previousADDriveSetting = $env:ADPS_LoadDefaultDrive
try {
    $env:ADPS_LoadDefaultDrive = '0'
    Import-Module ActiveDirectory -ErrorAction Stop
} finally {
    if ($null -eq $previousADDriveSetting) { Remove-Item Env:ADPS_LoadDefaultDrive -ErrorAction SilentlyContinue }
    else { $env:ADPS_LoadDefaultDrive = $previousADDriveSetting }
}
[Windows.Forms.Application]::EnableVisualStyles()
$script:preview = $null
$script:busy = $false
$script:rows = New-Object System.Data.DataTable
[void]$script:rows.Columns.Add('Select',[bool])
foreach ($name in 'Action','Location','Group','GroupId','Reason','Result') { [void]$script:rows.Columns.Add($name,[string]) }
$script:items = @{}
$form = New-Object Windows.Forms.Form
$form.Text = 'Hybrid AD / Entra - Copy User Groups'
$form.Size = New-Object Drawing.Size(1120,960)
$form.MinimumSize = New-Object Drawing.Size(1080,820)
$form.StartPosition = 'CenterScreen'
$top = New-Object Windows.Forms.Panel
$top.Dock = 'Top'; $top.Height = 240; $form.Controls.Add($top)
function Label($text,$x,$y,$width=170) {
    $c = New-Object Windows.Forms.Label
    $c.Text=$text; $c.Location=New-Object Drawing.Point($x,$y); $c.Size=New-Object Drawing.Size($width,22)
    $top.Controls.Add($c)
}
function TextBox($x,$y,$width=320) {
    $c=New-Object Windows.Forms.TextBox
    $c.Location=New-Object Drawing.Point($x,$y); $c.Width=$width; $top.Controls.Add($c); return $c
}
function Check($text,$x,$y,$checked=$false,$width=450) {
    $c=New-Object Windows.Forms.CheckBox
    $c.Text=$text; $c.Location=New-Object Drawing.Point($x,$y); $c.Size=New-Object Drawing.Size($width,25)
    $c.Checked=$checked; $top.Controls.Add($c); return $c
}
Label 'Source user UPN:' 15 18
$source=TextBox 185 15
Label 'Target user UPN:' 550 18
$target=TextBox 720 15
Label 'AD domain / DC:' 15 53
$server=TextBox 185 50
# Suggest the logon AD DNS domain without querying a DC at startup.
if ($env:USERDNSDOMAIN) { $server.Text=$env:USERDNSDOMAIN }
Label 'Tenant ID (cloud option):' 550 53
$tenant=TextBox 720 50
$clear=Check 'Remove target-only groups when copying' 15 88
$only=Check 'Clear TARGET groups only (no source required)' 550 88
$cloud=Check 'Include supported cloud-only Entra groups' 15 118
$test=Check 'Test mode - preview actions without changes' 550 118 $true
Label 'Direct memberships only. Both AD users must be in the selected AD domain; UPN suffixes can differ.' 15 155 1040
Label 'Protected AD groups are unchecked by default. Primary, dynamic and unsupported cloud groups are preserved.' 15 179 1040
Label 'All Add / Remove actions affect ONLY the TARGET UPN. The SOURCE account is read-only.' 15 207 1040
$bottom=New-Object Windows.Forms.Panel
$bottom.Dock='Bottom'; $bottom.Height=100; $form.Controls.Add($bottom)
$status=New-Object Windows.Forms.Label
$status.Location=New-Object Drawing.Point(15,48); $status.Size=New-Object Drawing.Size(1060,43)
$status.Text='Enter users, then click Preview. Test mode is enabled.'; $bottom.Controls.Add($status)
function Button($text,$x,$handler) {
    $b=New-Object Windows.Forms.Button
    $b.Text=$text; $b.Location=New-Object Drawing.Point($x,8); $b.Size=New-Object Drawing.Size(170,32)
    $b.Add_Click($handler); $bottom.Controls.Add($b); return $b
}
$grid=New-Object Windows.Forms.DataGridView
$grid.Dock='Fill'; $grid.AutoGenerateColumns=$false
$grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false
$grid.AutoSizeColumnsMode='Fill'; $grid.RowHeadersVisible=$false
# Create columns before binding: auto-generated columns may not exist until shown.
foreach ($name in 'Select','Action','Location','Group','GroupId','Reason','Result') {
    if ($name -eq 'Select') { $column=New-Object Windows.Forms.DataGridViewCheckBoxColumn }
    else { $column=New-Object Windows.Forms.DataGridViewTextBoxColumn }
    $column.Name=$name; $column.HeaderText=$name; $column.DataPropertyName=$name
    $column.ReadOnly=($name -ne 'Select')
    if ($name -eq 'Select') { $column.FillWeight=35 }
    if ($name -eq 'GroupId') { $column.Visible=$false }
    [void]$grid.Columns.Add($column)
}
$grid.DataSource=$script:rows.DefaultView
# Side-by-side memberships above the selectable action plan.
$body=New-Object Windows.Forms.TableLayoutPanel
$body.Dock='Fill'; $body.ColumnCount=2; $body.RowCount=2
[void]$body.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle -ArgumentList ([Windows.Forms.SizeType]::Percent),50))
[void]$body.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle -ArgumentList ([Windows.Forms.SizeType]::Percent),50))
[void]$body.RowStyles.Add((New-Object Windows.Forms.RowStyle -ArgumentList ([Windows.Forms.SizeType]::Percent),55))
[void]$body.RowStyles.Add((New-Object Windows.Forms.RowStyle -ArgumentList ([Windows.Forms.SizeType]::Percent),45))
$sourceBox=New-Object Windows.Forms.GroupBox
$sourceBox.Text='SOURCE - read-only membership preview'; $sourceBox.Dock='Fill'
$targetBox=New-Object Windows.Forms.GroupBox
$targetBox.Text='TARGET - membership preview'; $targetBox.Dock='Fill'
$actionsBox=New-Object Windows.Forms.GroupBox
$actionsBox.Text='Proposed actions - TARGET ONLY (select changes below)'; $actionsBox.Dock='Fill'
function New-MembershipGrid {
    $view=New-Object Windows.Forms.DataGridView
    $view.Dock='Fill'; $view.AutoGenerateColumns=$false; $view.ReadOnly=$true
    $view.AllowUserToAddRows=$false; $view.AllowUserToDeleteRows=$false
    $view.RowHeadersVisible=$false; $view.AutoSizeColumnsMode='Fill'
    foreach ($name in 'Group','Location','Notes') {
        $col=New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name=$name; $col.HeaderText=$name; $col.DataPropertyName=$name
        [void]$view.Columns.Add($col)
    }
    return $view
}
$sourceGrid=New-MembershipGrid
$targetGrid=New-MembershipGrid
$sourceBox.Controls.Add($sourceGrid); $targetBox.Controls.Add($targetGrid)
$actionsBox.Controls.Add($grid)
$body.Controls.Add($sourceBox,0,0); $body.Controls.Add($targetBox,1,0)
$body.Controls.Add($actionsBox,0,1); $body.SetColumnSpan($actionsBox,2)
$form.Controls.Add($body); $body.BringToFront()
function Fill-MembershipGrid($view,$groups) {
    $table=New-Object System.Data.DataTable
    foreach ($name in 'Group','Location','Notes') { [void]$table.Columns.Add($name,[string]) }
    foreach ($g in @($groups | Sort-Object Location,Name)) {
        $notes=$g.Reason
        if (-not $notes -and $g.Protected) { $notes='Protected AD group (adminCount=1)' }
        [void]$table.Rows.Add([string]$g.Name,[string]$g.Location,[string]$notes)
    }
    $view.DataSource=$table.DefaultView
}
function Invalidate-Preview {
    $script:preview=$null; $script:rows.Clear(); $script:items=@{}
    $sourceGrid.DataSource=$null; $targetGrid.DataSource=$null
    $sourceBox.Text='SOURCE - read-only (click Preview to load)'
    $targetBox.Text='TARGET - click Preview to load'
    $status.Text='Inputs changed. Click Preview to reload memberships and proposed target actions.'
}
function Note($text) { $status.Text=$text; [Windows.Forms.Application]::DoEvents() }
function Busy($value) { $script:busy=$value; $top.Enabled= -not $value; $bottom.Enabled= -not $value; $grid.Enabled= -not $value }
function Escape-Ldap($s) { $s.Replace('\','\5c').Replace('*','\2a').Replace('(','\28').Replace(')','\29').Replace([string][char]0,'\00') }
function Resolve-AD($upn,$dc) {
    if ([string]::IsNullOrWhiteSpace($upn)) { throw 'Enter the required UPN.' }
    $users=@(Get-ADUser -LDAPFilter "(userPrincipalName=$(Escape-Ldap $upn))" -Server $dc -Properties memberOf,primaryGroupID,Enabled)
    if ($users.Count -ne 1) { throw "UPN '$upn' did not resolve uniquely in AD domain/DC '$dc'." }
    return $users[0]
}
function Get-DirectAD($u,$dc) {
    foreach ($dn in @($u.memberOf)) {
        $g=Get-ADGroup -Identity $dn -Server $dc -Properties adminCount
        [pscustomobject]@{ Key="AD:$($g.ObjectGUID)"; Id=$g.DistinguishedName; SID=$g.SID.Value; Name=$g.Name; Location='Windows AD'; Server=$dc; Protected=($g.adminCount -eq 1); Reason='' }
    }
}
function Graph-Pages($uri) {
    do {
        $response=Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
        foreach ($v in @($response.value)) { $v }
        $uri=$response.'@odata.nextLink'
    } while ($uri)
}
function Cloud-User($ad) {
    $id=[uri]::EscapeDataString($ad.UserPrincipalName)
    $u=Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/${id}?`$select=id,userPrincipalName,onPremisesSecurityIdentifier,onPremisesSyncEnabled" -OutputType PSObject
    if (-not $u.onPremisesSyncEnabled -or $u.onPremisesSecurityIdentifier -ne $ad.SID.Value) {
        throw "Cloud identity does not match the synced AD SID for $($ad.UserPrincipalName)."
    }
    return $u
}
function Get-DirectCloud($u) {
    $uri="https://graph.microsoft.com/v1.0/users/$($u.id)/memberOf/microsoft.graph.group?`$select=id,displayName,groupTypes,onPremisesSyncEnabled,onPremisesSecurityIdentifier,mailEnabled,securityEnabled,isAssignableToRole&`$top=999"
    foreach ($g in @(Graph-Pages $uri)) {
        $reason=''
        if ($g.onPremisesSyncEnabled -or $g.onPremisesSecurityIdentifier) { $reason='Synced / previously synced: manage in AD' }
        elseif (@($g.groupTypes) -contains 'DynamicMembership') { $reason='Dynamic membership: managed by rules' }
        elseif ($g.isAssignableToRole) { $reason='Role-assignable group: excluded' }
        elseif ($g.mailEnabled -and @($g.groupTypes) -notcontains 'Unified') { $reason='Cloud mail-enabled security / distribution: use Exchange' }
        elseif (-not $g.securityEnabled -and @($g.groupTypes) -notcontains 'Unified') { $reason='Unsupported cloud group type' }
        [pscustomobject]@{ Key="Cloud:$($g.id)"; Id=$g.id; SID=[string]$g.onPremisesSecurityIdentifier; Name=$g.displayName; Location='Entra cloud'; Server=''; Protected=$false; Reason=$reason }
    }
}
function Options-Key { "$($source.Text.Trim())|$($target.Text.Trim())|$($server.Text.Trim())|$($tenant.Text.Trim())|$($cloud.Checked)|$($clear.Checked)|$($only.Checked)" }
function Load-State {
    $dc=$server.Text.Trim()
    if (-not $dc) { throw 'Enter an AD domain or domain controller.' }
    try { $dc=(Get-ADRootDSE -Server $dc -ErrorAction Stop).dnsHostName }
    catch {
        throw "Cannot connect to AD server '$dc'. Enter a reachable domain controller DNS name in AD domain / DC. Verify your domain network/VPN, permissions and Active Directory Web Services. Details: $($_.Exception.Message)"
    }
    $t=Resolve-AD $target.Text.Trim() $dc
    $s=$null; $sp=$null; $sg=@(); $tg=@(Get-DirectAD $t $dc); $tc=$null
    $primarySid=$t.SID.Value.Substring(0,$t.SID.Value.LastIndexOf('-'))+'-'+$t.primaryGroupID
    $primary=Get-ADGroup -Identity $primarySid -Server $dc
    $tg+= [pscustomobject]@{ Key="AD:$($primary.ObjectGUID)"; Id=$primary.DistinguishedName; SID=$primary.SID.Value; Name=$primary.Name; Location='Windows AD'; Server=$dc; Protected=$true; Reason='Primary group: always preserved' }
    if (-not $only.Checked) {
        $s=Resolve-AD $source.Text.Trim() $dc
        if ($s.ObjectGUID -eq $t.ObjectGUID) { throw 'Source and target must be different users.' }
        $sg=@(Get-DirectAD $s $dc)
        $sourcePrimarySid=$s.SID.Value.Substring(0,$s.SID.Value.LastIndexOf('-'))+'-'+$s.primaryGroupID
        $sourcePrimary=Get-ADGroup -Identity $sourcePrimarySid -Server $dc
        $sp=[pscustomobject]@{ Name=$sourcePrimary.Name; SID=$sourcePrimary.SID.Value; Location='Windows AD'; Protected=$true; Reason='Primary group: shown for reference; not copied' }
    }
    if ($cloud.Checked) {
        $tid=[guid]::Empty
        if (-not [guid]::TryParse($tenant.Text.Trim(),[ref]$tid)) { throw 'Enter the expected Entra tenant GUID.' }
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $ctx=Get-MgContext
        if (-not $ctx -or $ctx.TenantId -ne $tid.ToString() -or @($ctx.Scopes) -notcontains 'GroupMember.ReadWrite.All' -or @($ctx.Scopes) -notcontains 'Group.Read.All' -or @($ctx.Scopes) -notcontains 'User.Read.All') {
            if ($ctx) { Disconnect-MgGraph | Out-Null }
            Note 'Opening Microsoft sign-in. Complete authentication in the browser or Microsoft account window (check behind this GUI).'
            $form.Refresh()
            try {
                # Interactive authentication opens the browser/account prompt.
                # Device-code authentication does not open a sign-in window.
                Connect-MgGraph -TenantId $tid.ToString() -Scopes 'User.Read.All','Group.Read.All','GroupMember.ReadWrite.All' -ContextScope Process -NoWelcome -ErrorAction Stop
            } catch {
                throw "Microsoft Graph sign-in failed: $($_.Exception.Message)"
            }
        }
        $ctx=Get-MgContext
        if (-not $ctx) { throw 'Microsoft Graph sign-in did not establish a session. Preview stopped.' }
        if ($ctx.TenantId -ne $tid.ToString()) { throw 'Connected Graph tenant differs from the expected tenant.' }
        Note "Connected to Entra as $($ctx.Account). Loading cloud group memberships..."
        $tc=Cloud-User $t
        $tg+=@(Get-DirectCloud $tc)
        if ($s) { $sc=Cloud-User $s; $sg+=@(Get-DirectCloud $sc) }
    }
    $fingerprint=(@($sg | ForEach-Object { "S:$($_.Key):$($_.Reason):$($_.Protected)" }) + @($tg | ForEach-Object { "T:$($_.Key):$($_.Reason):$($_.Protected)" }) | Sort-Object) -join '|'
    [pscustomobject]@{ Source=$s; SourcePrimary=$sp; Target=$t; SourceGroups=$sg; TargetGroups=$tg; CloudTarget=$tc; DC=$dc; Options=(Options-Key); Fingerprint="$($s.ObjectGUID)|$($t.ObjectGUID)|$($t.primaryGroupID)|$fingerprint" }
}
function Add-Row($g,$action) {
    $reason=$g.Reason; $selected=$action -in 'Add','Remove'
    if ($reason) { $action='Skip'; $selected=$false }
    elseif ($g.Protected) { $reason='AD adminCount=1: review before selecting'; $selected=$false }
    $row=$script:rows.NewRow()
    $row.Select=$selected; $row.Action=$action; $row.Location=$g.Location; $row.Group=$g.Name
    $row.GroupId=$g.Key; $row.Reason=$reason; $row.Result='Pending'
    $script:rows.Rows.Add($row); $script:items[$g.Key]=$g
}
function Get-DisplayMemberships($groups) {
    # Match the immutable on-premises SID, never the display name: distinct
    # AD and cloud-only groups can legitimately have identical names.
    $adSids=@{}
    foreach ($g in $groups) {
        if ($g.Location -eq 'Windows AD' -and $g.PSObject.Properties['SID'] -and $g.SID) { $adSids[[string]$g.SID]=$true }
    }
    foreach ($g in $groups) {
        if ($g.Location -eq 'Entra cloud' -and $g.PSObject.Properties['SID'] -and $g.SID -and $adSids.ContainsKey([string]$g.SID)) { continue }
        $g
    }
}
function Make-Preview($state) {
    $script:rows.Clear(); $script:items=@{}
    $sourceKeys=@($state.SourceGroups | ForEach-Object { $_.Key })
    $targetKeys=@($state.TargetGroups | ForEach-Object { $_.Key })
    foreach ($g in $state.SourceGroups) {
        # Shared memberships need no action. Synced cloud entries are already
        # represented by their authoritative Windows AD memberships.
        if ($targetKeys -contains $g.Key) { continue }
        if ($g.Reason -eq 'Synced / previously synced: manage in AD') { continue }
        Add-Row $g 'Add'
    }
    foreach ($g in $state.TargetGroups) {
        if ($sourceKeys -contains $g.Key) { continue }
        if ($g.Reason -in @('Primary group: always preserved','Synced / previously synced: manage in AD')) { continue }
        if ($only.Checked -or $clear.Checked) { Add-Row $g 'Remove' } else { Add-Row $g 'Keep' }
    }
    $displaySource=@($state.SourceGroups)
    if ($state.SourcePrimary) { $displaySource+= $state.SourcePrimary }
    $displaySource=@(Get-DisplayMemberships $displaySource)
    $displayTarget=@(Get-DisplayMemberships $state.TargetGroups)
    Fill-MembershipGrid $sourceGrid $displaySource
    Fill-MembershipGrid $targetGrid $displayTarget
    if ($state.Source) { $sourceBox.Text="SOURCE (read-only): $($state.Source.UserPrincipalName) - $($displaySource.Count) groups" }
    else { $sourceBox.Text='SOURCE - not used in Clear TARGET groups only mode' }
    $targetBox.Text="TARGET: $($state.Target.UserPrincipalName) - $($displayTarget.Count) groups"
    $actionsBox.Text="Membership differences - ONLY for $($state.Target.UserPrincipalName) (shared groups hidden)"
    $script:preview=$state
}
function Save-Table($path) {
    @($script:rows.Rows | Select-Object Select,Action,Location,Group,GroupId,Reason,Result) | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
}
$previewButton=Button 'Preview groups' 15 {
    if ($script:busy) { return }
    Busy $true
    try { Invalidate-Preview; Note 'Reading direct AD / Entra memberships...'; Make-Preview (Load-State); Note "Preview ready: $($script:rows.Rows.Count) groups. Review selections, then run." }
    catch { Note $_.Exception.Message; [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Preview failed') }
    finally { Busy $false }
}
$applyButton=Button 'Run selected actions' 195 {
    if ($script:busy) { return }
    $grid.EndEdit(); Busy $true
    try {
        if (-not $script:preview -or (Options-Key) -ne $script:preview.Options) { throw 'Inputs changed or no preview exists. Click Preview again.' }
        $actions=@($script:rows.Rows | Where-Object { $_.Select -and $_.Action -in 'Add','Remove' })
        if (-not $actions.Count) { throw 'Select at least one Add or Remove action.' }
        Note 'Rechecking identities and membership before running...'
        $current=Load-State
        if ($current.Source -and $current.Source.ObjectGUID -eq $current.Target.ObjectGUID) { throw 'Source and target must be different accounts.' }
        if ($current.Fingerprint -ne $script:preview.Fingerprint) { $script:preview=$null; throw 'Membership or identity changed since preview. Click Preview again.' }
        $mode=if ($test.Checked) { 'TEST - no changes' } else { 'LIVE - changes will be made' }
        $adds=@($actions | Where-Object Action -eq 'Add').Count
        $removes=@($actions | Where-Object Action -eq 'Remove').Count
        $answer=[Windows.Forms.MessageBox]::Show("$mode`r`nOnly this TARGET account will be changed: $($current.Target.UserPrincipalName)`r`nAdd: $adds    Remove: $removes`r`nThe source account will NOT be modified.`r`n`r`nGroup-based licensing and resource access can change.`r`nProceed?",'Confirm selected group actions','YesNo','Warning')
        if ($answer -ne 'Yes') { Note 'Cancelled. No changes made.'; return }
        $root=Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'HybridGroupReports'
        $folder=Join-Path $root ((Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,6))
        [void](New-Item -ItemType Directory -Path $folder -Force)
        @($current.TargetGroups | Select-Object Name,Location,Id,Server,Protected,Reason) | Export-Csv (Join-Path $folder 'Target-Before.csv') -NoTypeInformation -Encoding UTF8
        @($current.SourceGroups | Select-Object Name,Location,Id,Server,Protected,Reason) | Export-Csv (Join-Path $folder 'Source-Before.csv') -NoTypeInformation -Encoding UTF8
        $report=Join-Path $folder 'Actions.csv'
        Save-Table $report
        $failedAdd=$false; $i=0
        foreach ($row in @($actions | Sort-Object @{Expression={if ($_.Action -eq 'Add') {0} else {1}}})) {
            $i++; Note "Processing $i of $($actions.Count): $($row.Action) $($row.Group)"
            $g=$script:items[$row.GroupId]
            if ($test.Checked) { $row.Result='Test only - no change'; Save-Table $report; continue }
            if ($row.Action -eq 'Remove' -and $failedAdd) { $row.Result='Withheld: an addition failed'; Save-Table $report; continue }
            try {
                if ($g.Location -eq 'Windows AD') {
                    if ($row.Action -eq 'Add') { Add-ADGroupMember -Identity $g.Id -Members $current.Target.DistinguishedName -Server $g.Server -ErrorAction Stop }
                    else { Remove-ADGroupMember -Identity $g.Id -Members $current.Target.DistinguishedName -Server $g.Server -Confirm:$false -ErrorAction Stop }
                    $check=Get-ADGroup -Identity $g.Id -Server $g.Server -Properties member
                    $present=@($check.member) -contains $current.Target.DistinguishedName
                } else {
                    $uid=$current.CloudTarget.id
                    if ($row.Action -eq 'Add') {
                        $body=@{'@odata.id'="https://graph.microsoft.com/v1.0/directoryObjects/$uid"} | ConvertTo-Json
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups/$($g.Id)/members/`$ref" -Body $body -ContentType 'application/json' | Out-Null
                    } else {
                        # Always use /$ref: removing the reference must never delete the user.
                        Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/groups/$($g.Id)/members/${uid}/`$ref" | Out-Null
                    }
                    $row.Result='API succeeded - cloud propagation may be pending'
                    Save-Table $report; continue
                }
                if ($present -ne ($row.Action -eq 'Add')) { throw 'Write returned but membership verification did not match; review the group.' }
                $row.Result='Succeeded and verified in AD'
            } catch {
                $row.Result="Failed / review: $($_.Exception.Message)"
                if ($row.Action -eq 'Add') { $failedAdd=$true }
            }
            Save-Table $report
        }
        if (-not $test.Checked) { $script:preview=$null }
        Note "Finished. CSV reports: $folder. AD changes appear in Entra after normal synchronization."
        [void][Windows.Forms.MessageBox]::Show("Finished. Review Actions.csv for successes, skips or failures.`r`nReports: $folder",'Group processing complete')
    } catch { Note $_.Exception.Message; [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Operation stopped') }
    finally { Busy $false }
}
$exportButton=Button 'Export displayed CSV' 375 {
    if ($script:busy) { return }
    $grid.EndEdit()
    $dialog=New-Object Windows.Forms.SaveFileDialog
    $dialog.Filter='CSV files (*.csv)|*.csv'; $dialog.FileName='Hybrid-Group-Preview.csv'
    if ($dialog.ShowDialog() -eq 'OK') {
        try { Save-Table $dialog.FileName; Note "Exported: $($dialog.FileName)" }
        catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Export failed') }
    }
    $dialog.Dispose()
}
$selectButton=Button 'Select eligible actions' 555 {
    foreach ($r in $script:rows.Rows) { $g=$script:items[$r.GroupId]; $r.Select=($r.Action -in 'Add','Remove' -and -not $g.Protected -and -not $g.Reason) }
}
$noneButton=Button 'Uncheck all' 735 { foreach ($r in $script:rows.Rows) { $r.Select=$false } }
$only.Add_CheckedChanged({ $source.Enabled= -not $only.Checked; $clear.Enabled= -not $only.Checked; Invalidate-Preview })
foreach ($inputBox in @($source,$target,$server,$tenant)) { $inputBox.Add_TextChanged({ Invalidate-Preview }) }
foreach ($option in @($cloud,$clear)) { $option.Add_CheckedChanged({ Invalidate-Preview }) }
$form.Add_FormClosing({ param($sender,$e) if ($script:busy) { $e.Cancel=$true } })
[void]$form.ShowDialog()
$form.Dispose()
