#requires -Version 7.0
<#
.SYNOPSIS
 Bulk Active Directory User Creator - PowerShell 7 / Windows Forms
.DESCRIPTION
 Creates on-prem AD users from a CSV, setting mail and uppercase SMTP: primary proxyAddresses. Assign a template per row in the GUI, preview
 selectable group memberships, validate, create selected accounts and export reports.
 Requires Windows, RSAT ActiveDirectory module, and AD create/group modification rights.
 No Entra or cloud verification is performed.
#>
[CmdletBinding()]
param()

if (-not $IsWindows) { throw 'This GUI requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-STA','-File',('"' + $PSCommandPath + '"'))
    return
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
try { Import-Module ActiveDirectory -ErrorAction Stop } catch {
    [Windows.Forms.MessageBox]::Show("Install RSAT Active Directory PowerShell tools first.`r`n$($_.Exception.Message)",'Missing module','OK','Error') | Out-Null
    return
}

$script:Rows = [Collections.Generic.List[object]]::new()
$script:Results = [Collections.Generic.List[object]]::new()
$script:LogLines = [Collections.Generic.List[string]]::new()
$script:CsvPath = ''
$script:Validated = $false
$script:DefaultOU = ''
$script:DefaultDomain = ''
$script:PasswordColumn = 'Password'
$script:AdServer = ''
$script:AdCredential = $null
$script:Connected = $false
$script:SelectedRow = $null
$script:AvailableOUs = @()

function Log([string]$message) {
    $line = ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$message)
    $script:LogLines.Add($line)
    $txtLog.AppendText($line + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
    [Windows.Forms.Application]::DoEvents()
}
function ADParams {
    $params = @{}
    if (-not [string]::IsNullOrWhiteSpace($script:AdServer)) { $params.Server = $script:AdServer }
    if ($null -ne $script:AdCredential) { $params.Credential = $script:AdCredential }
    return $params
}
function AssertADConnection {
    if (-not $script:Connected) { throw 'Connect to Active Directory and test credentials before continuing.' }
    if ($txtServer.Text.Trim() -ne $script:AdServer) { $script:Connected=$false; throw 'AD server changed. Connect again.' }
}

function GetADUser([string]$identity) {
    if ([string]::IsNullOrWhiteSpace($identity)) { return $null }
    $p = ADParams
    try { return Get-ADUser -Identity $identity -Properties DistinguishedName,DisplayName,Department,Title,Office,Company,Description,Manager,MemberOf,UserPrincipalName -ErrorAction Stop @p } catch {}
    $escaped = $identity.Replace("'","''")
    try { return @(Get-ADUser -Filter "UserPrincipalName -eq '$escaped'" -Properties DistinguishedName,DisplayName,Department,Title,Office,Company,Description,Manager,MemberOf,UserPrincipalName -ErrorAction Stop @p | Select-Object -First 1)[0] } catch { return $null }
}
function FindExisting([string]$sam,[string]$upn) {
    $p = ADParams
    $a=$sam.Replace("'","''"); $b=$upn.Replace("'","''")
    $found = @(Get-ADUser -Filter "SamAccountName -eq '$a' -or UserPrincipalName -eq '$b'" -ErrorAction Stop @p)
    return ($found.Count -gt 0)
}
function GetPrimaryEmail($r) {
    $email=([string]$r.Email).Trim()
    if (-not $email -and $chkEmailUPN.Checked) { $email=ResolveRowUPN $r }
    return $email.ToLowerInvariant()
}
function FindExistingProxyAddress([string]$email) {
    $p=ADParams
    $escaped=("SMTP:$email").Replace("'","''")
    # AD matching is case-insensitive; catches primary AND secondary proxy collisions.
    return @((Get-ADObject -LDAPFilter "(proxyAddresses=$escaped)" -ErrorAction Stop @p)).Count -gt 0
}
function HasColumn($row,[string]$name) { return ($row.PSObject.Properties.Name -contains $name) }
function Field($row,[string]$name) {
    if (HasColumn $row $name) { return [string]$row.$name }
    return ''
}
function ResolveRowUPN($row) {
    $upn = $row.UPN.Trim()
    if (-not $upn -and $row.SamAccountName -and $script:DefaultDomain) { $upn = "$($row.SamAccountName)@$($script:DefaultDomain)" }
    return $upn
}
function GetRowOU($row,$template) {
    if ($row.OU.Trim()) { return $row.OU.Trim() }
    if ($script:DefaultOU.Trim()) { return $script:DefaultOU.Trim() }
    if ($template) { return ($template.DistinguishedName -replace '^[^,]+,','') }
    return ''
}
function SetStatus($row,[string]$status,[string]$detail) {
    $row.Status=$status; $row.Detail=$detail
    $script:Validated=$false
}
function RefreshOUChoices {
    AssertADConnection
    $p=ADParams
    $script:AvailableOUs=@(Get-ADOrganizationalUnit -Filter * -ErrorAction Stop @p | Sort-Object DistinguishedName | ForEach-Object DistinguishedName)
    $previous=[string]$cmbOU.SelectedItem
    $cmbOU.Items.Clear(); [void]$cmbOU.Items.Add('')
    foreach ($dn in $script:AvailableOUs) { [void]$cmbOU.Items.Add($dn) }
    if ($previous -and $script:AvailableOUs -contains $previous) { $cmbOU.SelectedItem=$previous } else { $cmbOU.SelectedIndex=0 }
    $col=[Windows.Forms.DataGridViewComboBoxColumn]$grid.Columns['OU']
    $col.Items.Clear(); [void]$col.Items.Add('')
    foreach ($dn in $script:AvailableOUs) { [void]$col.Items.Add($dn) }
    Log "Loaded $($script:AvailableOUs.Count) OUs from Active Directory."
}
function LoadGrid {
    $grid.Rows.Clear()
    foreach ($r in $script:Rows) {
        $index = $grid.Rows.Add($r.Selected,$r.FirstName,$r.LastName,$r.SamAccountName,(ResolveRowUPN $r),$r.Template,$(if ($script:AvailableOUs -contains $r.OU) { $r.OU } else { '' }),$r.Status,$r.Detail)
        $grid.Rows[$index].Tag=$r
        if ($r.Status -eq 'Invalid' -or $r.Status -eq 'Failed') { $grid.Rows[$index].DefaultCellStyle.BackColor=[Drawing.Color]::MistyRose }
        elseif ($r.Status -eq 'Created') { $grid.Rows[$index].DefaultCellStyle.BackColor=[Drawing.Color]::Honeydew }
    }
    $lblCount.Text = "Loaded: $($script:Rows.Count)    Created: $(@($script:Rows | Where-Object Status -eq 'Created').Count)"
}
function SyncGridEdits {
    $grid.EndEdit()
    foreach ($g in $grid.Rows) {
        if ($null -eq $g.Tag) { continue }
        $r=$g.Tag
        $r.Selected=($g.Cells['Selected'].Value -eq $true)
        $r.Template=[string]$g.Cells['Template'].Value
        if ($script:AvailableOUs.Count -gt 0) { $r.OU=[string]$g.Cells['OU'].Value }
    }
}
function CheckRow($r) {
    if ($r.Status -eq 'Created') { return }
    $errors=[Collections.Generic.List[string]]::new()
    $r.UPN=ResolveRowUPN $r
    if (-not $r.FirstName.Trim()) { $errors.Add('Missing FirstName') }
    if (-not $r.LastName.Trim()) { $errors.Add('Missing LastName') }
    if ($r.SamAccountName -notmatch '^[a-zA-Z0-9._-]{1,20}$') { $errors.Add('Invalid SamAccountName (1-20 characters)') }
    if ($r.UPN -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { $errors.Add('Invalid UPN') }
    if (-not $r.Password) { $errors.Add('Missing Password') }
    $email=GetPrimaryEmail $r
    if (-not $email -or $email -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { $errors.Add('Missing or invalid primary SMTP email') }
    elseif ($email -match '[,;]' -or $email -match '^(?i:smtp:)') { $errors.Add('Email must be one plain address without SMTP prefix') }
    else { try { if (FindExistingProxyAddress $email) { $errors.Add('SMTP address already assigned in AD') } } catch { $errors.Add("SMTP lookup failed: $($_.Exception.Message)") } }
    if (-not $r.Template.Trim()) { $errors.Add('Choose a template') }
    $t=$null
    if ($r.Template.Trim()) {
        $t=GetADUser $r.Template.Trim()
        if (-not $t) { $errors.Add('Template not found') }
    }
    $ou=GetRowOU $r $t
    if (-not $ou) { $errors.Add('No target OU') }
    else {
        try { $adp=ADParams; Get-ADOrganizationalUnit -Identity $ou -ErrorAction Stop @adp | Out-Null }
        catch { $errors.Add('OU not found or inaccessible') }
    }
    if ($r.SamAccountName -and $r.UPN) {
        try { if (FindExisting $r.SamAccountName $r.UPN) { $errors.Add('Existing AD account (SAM or UPN)') } }
        catch { $errors.Add("AD lookup failed: $($_.Exception.Message)") }
    }
    if ($errors.Count) { $r.Status='Invalid'; $r.Detail=($errors -join '; ') }
    else { $r.Status='Ready'; $r.Detail='Validated (password policy checked on creation)' }
}
function GetSelectedTemplateGroups($r) {
    $t=GetADUser $r.Template
    if (-not $t) { throw "Template '$($r.Template)' not found" }
    $p=ADParams
    $groups=@()
    foreach ($dn in @($t.MemberOf)) {
        try {
            $g=Get-ADGroup -Identity $dn -Properties GroupCategory,GroupScope,Name,DistinguishedName -ErrorAction Stop @p
            $groups += $g
        } catch { Log "WARNING: Could not resolve template group: $dn" }
    }
    return @($groups | Sort-Object Name)
}
function EditGroups($r) {
    if (-not $r.Template) { [Windows.Forms.MessageBox]::Show('Choose a template first.') | Out-Null; return }
    try { $groups=@(GetSelectedTemplateGroups $r) } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message) | Out-Null; return }
    $dlg=[Windows.Forms.Form]::new(); $dlg.Text="Groups for $($r.SamAccountName)"; $dlg.Size=[Drawing.Size]::new(760,530); $dlg.StartPosition='CenterParent'
    $note=[Windows.Forms.Label]::new(); $note.Dock='Top'; $note.Height=45; $note.Text='Choose direct groups to copy. Built-in primary group (Domain Users) is not in MemberOf. Group scope and membership permissions are enforced by AD.'; $dlg.Controls.Add($note)
    $checked=[Windows.Forms.CheckedListBox]::new(); $checked.Dock='Fill'; $checked.CheckOnClick=$true
    $saved=@($r.GroupDNs)
    for ($i=0;$i -lt $groups.Count;$i++) {
        $g=$groups[$i]
        $label="$($g.Name) [$($g.GroupCategory)/$($g.GroupScope)]"
        [void]$checked.Items.Add($label)
        if (-not $r.GroupsCustomized -or $saved -contains $g.DistinguishedName) { $checked.SetItemChecked($i,$true) }
    }
    $dlg.Controls.Add($checked); $checked.BringToFront()
    $bottom=[Windows.Forms.FlowLayoutPanel]::new(); $bottom.Dock='Bottom'; $bottom.Height=45; $bottom.FlowDirection='RightToLeft'
    $ok=[Windows.Forms.Button]::new(); $ok.Text='Save Selection'; $ok.Width=130; $ok.DialogResult='OK'; $bottom.Controls.Add($ok)
    $cancel=[Windows.Forms.Button]::new(); $cancel.Text='Cancel'; $cancel.DialogResult='Cancel'; $bottom.Controls.Add($cancel)
    $dlg.Controls.Add($bottom)
    if ($dlg.ShowDialog($form) -eq 'OK') {
        $r.GroupDNs=@(for ($i=0;$i -lt $groups.Count;$i++) { if ($checked.GetItemChecked($i)) { $groups[$i].DistinguishedName } })
        $r.GroupsCustomized=$true
        Log "Selected $($r.GroupDNs.Count) groups for $($r.SamAccountName)"
    }
    $dlg.Dispose()
}
function CreateOne($r) {
    $p=ADParams
    $t=GetADUser $r.Template
    if (-not $t) { throw 'Template missing at creation time' }
    if (FindExisting $r.SamAccountName $r.UPN) { throw 'Account already exists: refusing overwrite' }
    $ou=GetRowOU $r $t
    $null=Get-ADOrganizationalUnit -Identity $ou -ErrorAction Stop @p
    $name=if ($r.DisplayName.Trim()) { $r.DisplayName.Trim() } else { "$($r.FirstName) $($r.LastName)" }
    $new=@{ Name=$name; GivenName=$r.FirstName; Surname=$r.LastName; DisplayName=$name; SamAccountName=$r.SamAccountName; UserPrincipalName=$r.UPN; Path=$ou; AccountPassword=(ConvertTo-SecureString $r.Password -AsPlainText -Force); Enabled=$true; ChangePasswordAtLogon=$chkChange.Checked; ErrorAction='Stop' }
    foreach ($field in @('Department','Title','Office','Company','Description')) {
        $value=[string]$r.$field
        if (-not $value.Trim() -and $chkAttributes.Checked) { $value=[string]$t.$field }
        if ($value.Trim()) { $new[$field]=$value.Trim() }
    }
    $email=GetPrimaryEmail $r
    if (-not $email -or $email -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$' -or $email -match '[,;]' -or $email -match '^(?i:smtp:)') { throw 'Invalid primary SMTP email' }
    if (FindExistingProxyAddress $email) { throw 'Primary SMTP address already assigned in AD' }
    $new['EmailAddress']=$email
    $new['OtherAttributes']=@{ proxyAddresses=@("SMTP:$email") }
    if ($r.Manager.Trim()) {
        $manager=GetADUser $r.Manager.Trim()
        if (-not $manager) { throw "Manager not found: $($r.Manager)" }
        $new['Manager']=$manager.DistinguishedName
    }
    if ($r.EmployeeID.Trim()) { $new['EmployeeID']=$r.EmployeeID.Trim() }
    $created=$false
    try {
        New-ADUser @new @p
        $created=$true
        $r.Status='Created'; $r.Detail='Account created; group assignment pending'
        Log "CREATED: $($r.UPN)"
        $groupFailures=@()
        if ($chkGroups.Checked) {
            $dns=if ($r.GroupsCustomized) { @($r.GroupDNs) } else { @((GetSelectedTemplateGroups $r) | ForEach-Object DistinguishedName) }
            foreach ($dn in $dns) {
                try { Add-ADGroupMember -Identity $dn -Members $r.SamAccountName -ErrorAction Stop @p; Log "  GROUP OK: $dn" }
                catch { $groupFailures += "${dn}: $($_.Exception.Message)"; Log "  GROUP FAILED: $dn -- $($_.Exception.Message)" }
            }
        }
        if ($groupFailures.Count) { $r.Status='Created (partial)'; $r.Detail=($groupFailures -join ' | ') }
        else { $r.Status='Created'; $r.Detail='Created successfully' }
    } catch {
        if ($created) { $r.Status='Created (partial)'; $r.Detail=$_.Exception.Message }
        else { $r.Status='Failed'; $r.Detail=$_.Exception.Message }
        throw
    } finally {
        $r.Password=$null
    }
}

$form=[Windows.Forms.Form]::new(); $form.Text='Bulk AD User Creator | On-Premises Active Directory'; $form.Size=[Drawing.Size]::new(1320,850); $form.MinimumSize=[Drawing.Size]::new(1000,650); $form.StartPosition='CenterScreen'; $form.Font=[Drawing.Font]::new('Segoe UI',9)
$top=[Windows.Forms.TableLayoutPanel]::new(); $top.Dock='Top'; $top.Height=117; $top.ColumnCount=4; $top.RowCount=3
$top.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',110)); $top.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Percent',50)); $top.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Absolute',110)); $top.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Percent',50))
function TLabel($s,$col,$row) { $x=[Windows.Forms.Label]::new(); $x.Text=$s; $x.Dock='Fill'; $x.TextAlign='MiddleLeft'; $top.Controls.Add($x,$col,$row) }
function TText($col,$row) { $x=[Windows.Forms.TextBox]::new(); $x.Dock='Fill'; $top.Controls.Add($x,$col,$row); return $x }
TLabel 'AD Server:' 0 0; $txtServer=TText 1 0; TLabel 'Default OU:' 2 0
$cmbOU=[Windows.Forms.ComboBox]::new(); $cmbOU.Dock='Fill'; $cmbOU.DropDownStyle='DropDownList'; $top.Controls.Add($cmbOU,3,0)
[void]$cmbOU.Items.Add(''); $cmbOU.SelectedIndex=0
TLabel 'UPN suffix:' 0 1; $txtDomain=TText 1 1; TLabel 'Template:' 2 1; $txtTemplate=TText 3 1
TLabel 'CSV file:' 0 2; $txtCsv=TText 1 2; $txtCsv.ReadOnly=$true
$buttons=[Windows.Forms.FlowLayoutPanel]::new(); $buttons.Dock='Top'; $buttons.Height=43; $buttons.WrapContents=$false; $buttons.AutoScroll=$true
function Btn([string]$label,[int]$width=115) { $b=[Windows.Forms.Button]::new(); $b.Text=$label; $b.Width=$width; $b.Height=29; $b.Margin=[Windows.Forms.Padding]::new(4); $buttons.Controls.Add($b); return $b }
$btnConnect=Btn 'Connect AD' 110; $btnRefreshOU=Btn 'Refresh OUs' 105; $btnApplyOU=Btn 'Apply OU to All' 125; $btnImport=Btn 'Import CSV'; $btnTemplate=Btn 'Assign Template' 125; $btnGroups=Btn 'Select Groups' 120; $btnValidate=Btn 'Validate' 90; $btnCreate=Btn 'Create Selected' 125; $btnExport=Btn 'Export Report' 115; $btnLog=Btn 'Export Log' 100; $btnSample=Btn 'Sample CSV' 100
$options=[Windows.Forms.FlowLayoutPanel]::new(); $options.Dock='Top'; $options.Height=34; $options.WrapContents=$false
function Opt($s,$default) { $c=[Windows.Forms.CheckBox]::new(); $c.Text=$s; $c.AutoSize=$true; $c.Checked=$default; $c.Margin=[Windows.Forms.Padding]::new(8,6,10,0); $options.Controls.Add($c); return $c }
$chkGroups=Opt 'Copy selected direct groups' $true
$chkAttributes=Opt 'Fill blank attributes from template' $true
$chkChange=Opt 'Require password change at first login' $true
$chkEmailUPN=Opt 'Use UPN as primary SMTP when Email blank' $true
$split=[Windows.Forms.SplitContainer]::new(); $split.Dock='Fill'; $split.Orientation='Horizontal'; $split.SplitterDistance=420
$grid=[Windows.Forms.DataGridView]::new(); $grid.Dock='Fill'; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.MultiSelect=$true; $grid.SelectionMode='FullRowSelect'; $grid.AutoSizeColumnsMode='Fill'; $grid.RowHeadersVisible=$false
$col=[Windows.Forms.DataGridViewCheckBoxColumn]::new(); $col.Name='Selected'; $col.HeaderText='Use'; $col.FillWeight=35; [void]$grid.Columns.Add($col)
foreach ($spec in @(@('FirstName',75),@('LastName',75),@('SamAccountName',100),@('UPN',140),@('Template',135),@('OU',170),@('Status',85),@('Detail',220))) {
    $c=if ($spec[0] -eq 'OU') { [Windows.Forms.DataGridViewComboBoxColumn]::new() } else { [Windows.Forms.DataGridViewTextBoxColumn]::new() }; $c.Name=$spec[0]; $c.HeaderText=$spec[0]; $c.FillWeight=[int]$spec[1]; $c.ReadOnly=($spec[0] -notin @('Template','OU')); [void]$grid.Columns.Add($c)
}
$split.Panel1.Controls.Add($grid)
$bottom=[Windows.Forms.Panel]::new(); $bottom.Dock='Fill'
$txtLog=[Windows.Forms.TextBox]::new(); $txtLog.Dock='Fill'; $txtLog.Multiline=$true; $txtLog.ReadOnly=$true; $txtLog.ScrollBars='Both'; $txtLog.WordWrap=$false; $txtLog.Font=[Drawing.Font]::new('Consolas',9)
$progress=[Windows.Forms.ProgressBar]::new(); $progress.Dock='Bottom'; $progress.Height=20
$lblCount=[Windows.Forms.Label]::new(); $lblCount.Dock='Top'; $lblCount.Height=25; $lblCount.Text='Loaded: 0'; $bottom.Controls.Add($txtLog); $bottom.Controls.Add($progress); $bottom.Controls.Add($lblCount)
$split.Panel2.Controls.Add($bottom)
$connectionBar=[Windows.Forms.Panel]::new(); $connectionBar.Dock='Top'; $connectionBar.Height=25
$lblConnection=[Windows.Forms.Label]::new(); $lblConnection.Dock='Fill'; $lblConnection.Text='Not connected - enter DC hostname and click Connect AD'; $lblConnection.ForeColor=[Drawing.Color]::DarkRed
$connectionBar.Controls.Add($lblConnection)
$form.Controls.Add($split); $form.Controls.Add($options); $form.Controls.Add($buttons); $form.Controls.Add($connectionBar); $form.Controls.Add($top)

$btnConnect.Add_Click({
    try {
        $server = $txtServer.Text.Trim()
        if (-not $server) { throw 'Enter a domain controller FQDN (e.g. dc01.contoso.local) in AD Server first.' }
        $credential = Get-Credential -Message 'Enter domain credentials (DOMAIN\username or username@domain). A local machine administrator account is not sufficient.'
        if ($null -eq $credential) { return }
        $script:Connected=$false
        $script:AdCredential=$null
        # Test against the explicit DC using the supplied credentials. No password saved to disk.
        $domain = Get-ADDomain -Server $server -Credential $credential -ErrorAction Stop
        $null = Get-ADUser -Filter * -ResultSetSize 1 -Server $server -Credential $credential -ErrorAction Stop
        $script:AdServer=$server
        $script:AdCredential=$credential
        $script:Connected=$true
        RefreshOUChoices
        LoadGrid
        $lblConnection.Text="Connected: $($domain.DNSRoot) via $server as $($credential.UserName)"
        $lblConnection.ForeColor=[Drawing.Color]::DarkGreen
        Log "AD connection successful: $($domain.DNSRoot) via $server as $($credential.UserName)"
    } catch {
        $script:Connected=$false; $script:AdCredential=$null
        $lblConnection.Text='Not connected'
        $lblConnection.ForeColor=[Drawing.Color]::DarkRed
        [Windows.Forms.MessageBox]::Show("AD connection failed:`r`n$($_.Exception.Message)",'AD connection','OK','Error') | Out-Null
    }
})
$btnApplyOU.Add_Click({
    try {
        SyncGridEdits
        $chosenOU=([string]$cmbOU.SelectedItem).Trim()
        if (-not $chosenOU) { throw 'Choose an OU from the Default OU dropdown first.' }
        if ($script:AvailableOUs -notcontains $chosenOU) { throw 'The chosen OU is not in the refreshed AD OU list.' }
        $eligible=@($script:Rows | Where-Object { $_.Status -ne 'Created' })
        if ($eligible.Count -eq 0) { throw 'No uncreated imported users to update.' }
        $answer=[Windows.Forms.MessageBox]::Show("Apply this OU to all $($eligible.Count) uncreated imported users?`n`n$chosenOU",'Confirm bulk OU assignment',[Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
        foreach ($r in $eligible) { $r.OU=$chosenOU; if ($r.Status -eq 'Ready' -or $r.Status -eq 'Invalid') { $r.Status='Pending'; $r.Detail='' } }
        $script:DefaultOU=$chosenOU
        $script:Validated=$false
        LoadGrid
        Log "Applied OU to $($eligible.Count) uncreated imported users: $chosenOU"
    } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Apply OU to All','OK','Warning') | Out-Null }
})
$btnRefreshOU.Add_Click({
    try { SyncGridEdits; RefreshOUChoices; LoadGrid }
    catch { [Windows.Forms.MessageBox]::Show("OU lookup failed: $($_.Exception.Message)",'OU lookup','OK','Error') | Out-Null }
})
$cmbOU.Add_SelectedIndexChanged({ $script:Validated=$false; $script:DefaultOU=[string]$cmbOU.SelectedItem })
$grid.Add_DataError({ param($sender,$e) $e.ThrowException=$false })
$txtServer.Add_TextChanged({ $script:Connected=$false; if ($lblConnection) { $lblConnection.Text='Not connected - reconnect required'; $lblConnection.ForeColor=[Drawing.Color]::DarkRed } })
$btnSample.Add_Click({
    $save=[Windows.Forms.SaveFileDialog]::new(); $save.Filter='CSV files (*.csv)|*.csv'; $save.FileName='BulkADUsers-Sample.csv'
    if ($save.ShowDialog() -eq 'OK') {
        @([pscustomobject]@{FirstName='Jane';LastName='Doe';SamAccountName='jdoe';UPN='jdoe@example.com';Password='REPLACE_WITH_PASSWORD';DisplayName='Jane Doe';Email='jdoe@example.com';OU='';Template='';Department='';Title='';Office='';Company='';Description='';Manager='';EmployeeID=''}) | Export-Csv -LiteralPath $save.FileName -NoTypeInformation -Encoding utf8
        Log "Sample CSV saved: $($save.FileName)"
    }
})
$btnImport.Add_Click({
    $open=[Windows.Forms.OpenFileDialog]::new(); $open.Filter='CSV files (*.csv)|*.csv'; $open.Multiselect=$false
    if ($open.ShowDialog() -ne 'OK') { return }
    try {
        $items=@(Import-Csv -LiteralPath $open.FileName -ErrorAction Stop)
        if (-not $items.Count) { throw 'CSV is empty' }
        foreach ($required in @('FirstName','LastName','SamAccountName','Password')) {
            if (-not (HasColumn $items[0] $required)) { throw "Missing required CSV column: $required" }
        }
        if (-not (HasColumn $items[0] 'UPN') -and -not $txtDomain.Text.Trim()) { throw 'CSV must include UPN, or provide a default UPN suffix.' }
        $script:Rows.Clear(); $script:Results.Clear(); $script:CsvPath=$open.FileName
        $script:DefaultDomain=$txtDomain.Text.Trim().TrimStart('@')
        $script:DefaultOU=([string]$cmbOU.SelectedItem).Trim()
        foreach ($item in $items) {
            $r=[pscustomobject]@{ Selected=$true; FirstName=(Field $item 'FirstName').Trim(); LastName=(Field $item 'LastName').Trim(); SamAccountName=(Field $item 'SamAccountName').Trim(); UPN=(Field $item 'UPN').Trim(); Password=(Field $item 'Password'); DisplayName=(Field $item 'DisplayName').Trim(); Email=(Field $item 'Email').Trim(); OU=(Field $item 'OU').Trim(); Template=(Field $item 'Template').Trim(); Department=(Field $item 'Department').Trim(); Title=(Field $item 'Title').Trim(); Office=(Field $item 'Office').Trim(); Company=(Field $item 'Company').Trim(); Description=(Field $item 'Description').Trim(); Manager=(Field $item 'Manager').Trim(); EmployeeID=(Field $item 'EmployeeID').Trim(); Status='Pending'; Detail=''; GroupDNs=@(); GroupsCustomized=$false }
            $script:Rows.Add($r)
        }
        $txtCsv.Text=$open.FileName; LoadGrid; Log "Imported $($items.Count) users. Passwords are held in memory only during this session; the source CSV still contains plaintext passwords."
    } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Import failed','OK','Error') | Out-Null }
})
$btnTemplate.Add_Click({
    SyncGridEdits
    $id=$txtTemplate.Text.Trim()
    if (-not $id) { [Windows.Forms.MessageBox]::Show('Enter a template username or UPN in the Template field above.') | Out-Null; return }
    try { AssertADConnection } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message) | Out-Null; return }
    $t=GetADUser $id
    if (-not $t) { [Windows.Forms.MessageBox]::Show('Template account not found.') | Out-Null; return }
    $selected=@($grid.SelectedRows | ForEach-Object Tag)
    if (-not $selected.Count) { [Windows.Forms.MessageBox]::Show('Select one or more grid rows first.') | Out-Null; return }
    foreach ($r in $selected) { if ($r.Status -ne 'Created') { $r.Template=$t.DistinguishedName; $r.GroupsCustomized=$false; $r.GroupDNs=@(); $r.Status='Pending'; $r.Detail='' } }
    LoadGrid; Log "Assigned template $($t.SamAccountName) to $($selected.Count) selected rows"
})
$btnGroups.Add_Click({
    SyncGridEdits
    if ($grid.SelectedRows.Count -ne 1) { [Windows.Forms.MessageBox]::Show('Select exactly one user row to choose groups.') | Out-Null; return }
    try { AssertADConnection } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message) | Out-Null; return }
    EditGroups $grid.SelectedRows[0].Tag
})
$btnValidate.Add_Click({
    try { AssertADConnection } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message) | Out-Null; return }
    SyncGridEdits; $script:DefaultDomain=$txtDomain.Text.Trim().TrimStart('@'); $script:DefaultOU=([string]$cmbOU.SelectedItem).Trim()
    $seenSam=@{}; $seenUPN=@{}; $seenEmail=@{}; $count=0
    foreach ($r in $script:Rows) {
        if (-not $r.Selected -or $r.Status -eq 'Created') { continue }
        try { CheckRow $r } catch { $r.Status='Invalid'; $r.Detail=$_.Exception.Message }
        $samKey=$r.SamAccountName.ToLowerInvariant(); $upnKey=$r.UPN.ToLowerInvariant(); $emailKey=GetPrimaryEmail $r
        if ($seenSam.ContainsKey($samKey) -or ($upnKey -and $seenUPN.ContainsKey($upnKey))) { $r.Status='Invalid'; $r.Detail='Duplicate SAM or UPN within CSV' }
        if ($emailKey -and $seenEmail.ContainsKey($emailKey)) { $r.Status='Invalid'; $r.Detail='Duplicate primary SMTP address within CSV' }
        $seenSam[$samKey]=$true; if ($upnKey) { $seenUPN[$upnKey]=$true }; if ($emailKey) { $seenEmail[$emailKey]=$true }
        $count++
    }
    $script:Validated=$true; LoadGrid
    Log "Validated $count selected rows. Ready: $(@($script:Rows | Where-Object {$_.Selected -and $_.Status -eq 'Ready'}).Count); Invalid: $(@($script:Rows | Where-Object {$_.Selected -and $_.Status -eq 'Invalid'}).Count)."
})
$btnCreate.Add_Click({
    try { AssertADConnection } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message) | Out-Null; return }
    SyncGridEdits
    if (-not $script:Validated) { [Windows.Forms.MessageBox]::Show('Click Validate immediately before creating accounts.') | Out-Null; return }
    $ready=@($script:Rows | Where-Object {$_.Selected -and $_.Status -eq 'Ready'})
    if (-not $ready.Count) { [Windows.Forms.MessageBox]::Show('No selected Ready users.') | Out-Null; return }
    $confirm=[Windows.Forms.MessageBox]::Show("Create $($ready.Count) on-prem AD accounts? This cannot be automatically undone. Passwords will not be exported.",'Confirm bulk creation','YesNo','Warning')
    if ($confirm -ne 'Yes') { return }
    foreach ($control in @($btnConnect,$btnApplyOU,$btnImport,$btnTemplate,$btnGroups,$btnValidate,$btnCreate,$grid,$txtServer,$cmbOU,$btnRefreshOU,$txtDomain,$chkGroups,$chkAttributes,$chkChange,$chkEmailUPN)) { $control.Enabled=$false }
    $progress.Maximum=$ready.Count; $progress.Value=0
    try {
        foreach ($r in $ready) {
            try { CreateOne $r }
            catch { Log "ERROR $($r.SamAccountName): $($_.Exception.Message)" }
            $script:Results.Add([pscustomobject]@{ Timestamp=(Get-Date -Format 's'); FirstName=$r.FirstName; LastName=$r.LastName; SamAccountName=$r.SamAccountName; UPN=$r.UPN; Template=$r.Template; OU=(GetRowOU $r (GetADUser $r.Template)); Status=$r.Status; Detail=$r.Detail })
            $progress.Value++; [Windows.Forms.Application]::DoEvents()
        }
        Log 'Batch complete. No Entra sync or cloud verification attempted.'
    } finally {
        foreach ($control in @($btnConnect,$btnApplyOU,$btnImport,$btnTemplate,$btnGroups,$btnValidate,$btnCreate,$grid,$txtServer,$cmbOU,$btnRefreshOU,$txtDomain,$chkGroups,$chkAttributes,$chkChange,$chkEmailUPN)) { $control.Enabled=$true }
        $script:Validated=$false; LoadGrid
    }
})
$btnExport.Add_Click({
    $save=[Windows.Forms.SaveFileDialog]::new(); $save.Filter='CSV files (*.csv)|*.csv'; $save.FileName="AD-Creation-Report-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
    if ($save.ShowDialog() -eq 'OK') {
        @($script:Rows | ForEach-Object { [pscustomobject]@{ FirstName=$_.FirstName; LastName=$_.LastName; SamAccountName=$_.SamAccountName; UPN=$_.UPN; Template=$_.Template; OU=$_.OU; Status=$_.Status; Detail=$_.Detail; SelectedGroupCount=$(if ($_.GroupsCustomized) {$_.GroupDNs.Count} else {'Template defaults'}) } }) | Export-Csv -LiteralPath $save.FileName -NoTypeInformation -Encoding utf8
        Log "Report exported (without passwords): $($save.FileName)"
    }
})
$btnLog.Add_Click({
    $save=[Windows.Forms.SaveFileDialog]::new(); $save.Filter='Text files (*.txt)|*.txt'; $save.FileName="AD-Creation-Log-$(Get-Date -Format yyyyMMdd-HHmmss).txt"
    if ($save.ShowDialog() -eq 'OK') { [IO.File]::WriteAllLines($save.FileName,$script:LogLines); Log 'Log exported.' }
})
$grid.Add_CellValueChanged({ param($sender,$e)
    if ($e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0 -and $grid.Columns[$e.ColumnIndex].Name -in @('Selected','Template','OU')) { $script:Validated=$false }
})
$grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit) | Out-Null } })
$form.Add_Shown({ Log 'Ready. Import CSV, select rows, assign per-user templates, choose groups, validate, then create.' })
$form.Add_FormClosed({ foreach ($r in $script:Rows) { $r.Password=$null } })
[void]$form.ShowDialog()

$script:AdCredential = $null
