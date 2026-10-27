#Requires -Version 5.1
<#
Run on Windows with RSAT ActiveDirectory, as an operator delegated to create users.
Launch: powershell.exe -NoProfile -STA -File .\New-HybridUser-GUI.ps1
Install Graph once: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
MFA uses Graph beta per-user MFA; requires Authentication Policy Administrator
and admin consent for User.Read.All, Policy.ReadWrite.AuthenticationMethod.
No Exchange mailbox provisioning, license assignment, or sync configuration changes.
Choose an OU included in your existing Entra Connect/Cloud Sync scope.
#>
[CmdletBinding()]param()
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Launch with powershell.exe -NoProfile -STA -File .\New-HybridUser-GUI.ps1' }
Import-Module ActiveDirectory
[Windows.Forms.Application]::EnableVisualStyles()
$script:created = $null
$script:reportSaved = $false
$script:mimic = 'N/A'
$script:copyIdentity = ''
$script:defaultGroups = @{
    'awmi.net' = @('365 AWM Licensing (AWMI)','All AWM Staff','All Colorado Staff','All HQ Staff','All Charis WP Staff')
    'awmcharis.com' = @('365 AWM Licensing (AWMCharis)','All Charis Staff','All Colorado Staff','All Charis WP Staff')
}
$script:previousDefaults = @()
$script:groupServers = @{}

$script:mfa = 'Pending - not yet enabled'
$script:issues = New-Object 'System.Collections.Generic.List[string]'
$script:controls = @{}
$script:rng = [Security.Cryptography.RandomNumberGenerator]::Create()
function Random-Index([int]$Max) {
    $b = New-Object byte[] 4
    $limit = [uint64]4294967296 - ([uint64]4294967296 % [uint64]$Max)
    do { $script:rng.GetBytes($b); $n = [BitConverter]::ToUInt32($b,0) } while ([uint64]$n -ge $limit)
    return [int]($n % $Max)
}
function Password-Error([string]$p) {
    if ($p.Length -ne 16 -or $p -cnotmatch '[A-Z]' -or $p -cnotmatch '[a-z]' -or $p -notmatch '\d' -or $p -notmatch '[!#$%&*@]' -or $p -match '[^A-Za-z0-9!#$%&*@]') { return 'Use exactly 16 characters with uppercase, lowercase, numbers and at least one of !#$%&*@.' }
    if ($p -match '\d{2}') { return 'Do not use adjacent numbers.' }
    foreach ($key in @('First','Last','Logon')) {
        foreach ($part in ($script:controls[$key].Text -split '[^a-zA-Z0-9]+')) {
            if ($part.Length -ge 3 -and $p.IndexOf($part,[StringComparison]::OrdinalIgnoreCase) -ge 0) { return 'Password must not contain name or username components of three or more characters.' }
        }
    }
    return $null
}
function Generate-Password {
    # O, I, l, o, 0 and 1 are omitted to reduce visually ambiguous characters.
    $sets = @('ABCDEFGHJKLMNPQRSTUVWXYZ','abcdefghijkmnpqrstuvwxyz','23456789','!#$%&*@')
    $all = $sets -join ''
    do {
        $chars = New-Object 'System.Collections.Generic.List[char]'
        foreach ($s in $sets) { $chars.Add($s[(Random-Index $s.Length)]) }
        while ($chars.Count -lt 16) { $chars.Add($all[(Random-Index $all.Length)]) }
        for ($i=15; $i -gt 0; $i--) { $j=Random-Index ($i+1); $v=$chars[$i]; $chars[$i]=$chars[$j]; $chars[$j]=$v }
        $p = -join $chars.ToArray()
    } while (Password-Error $p)
    $script:controls.Password.Text = $p
}
$f = New-Object Windows.Forms.Form
$f.Text = 'AWM - Hybrid Account Creation v7'
$f.Size = New-Object Drawing.Size(1000,900)
$f.MinimumSize = New-Object Drawing.Size(850,650)
$f.StartPosition = 'CenterScreen'
$panel = New-Object Windows.Forms.FlowLayoutPanel
$panel.Dock='Fill'; $panel.AutoScroll=$true; $panel.FlowDirection='TopDown'; $panel.WrapContents=$false
$f.Controls.Add($panel)
function Field([string]$key,[string]$label,[string]$kind='Text') {
    $row=New-Object Windows.Forms.Panel; $row.Size=New-Object Drawing.Size(940,34)
    $l=New-Object Windows.Forms.Label; $l.Text=$label; $l.Location=New-Object Drawing.Point(8,8); $l.Size=New-Object Drawing.Size(275,22)
    if ($kind -eq 'Combo') { $c=New-Object Windows.Forms.ComboBox; $c.DropDownStyle='DropDownList' }
    elseif ($kind -eq 'Check') { $c=New-Object Windows.Forms.CheckBox }
    else { $c=New-Object Windows.Forms.TextBox }
    $c.Location=New-Object Drawing.Point(285,4); $c.Size=New-Object Drawing.Size(630,26)
    $row.Controls.AddRange(@($l,$c)); $panel.Controls.Add($row); $script:controls[$key]=$c
}
Field Domain 'AD domain (account location)' Combo
Field Suffix 'User logon UPN suffix' Combo
Field NoGroups 'Intern / World Outreach - no groups' Check
$script:controls.NoGroups.Text='AWMCharis.com account; skip all optional AD groups'
$script:groupsBeforeNoGroups=@()
$script:copyGroupsBeforeNoGroups=$true

Field OU 'Destination OU / container' Combo
$script:controls.OU.Width=470
$script:controls.OU.DropDownWidth=1100
$script:controls.OU.MaxDropDownItems=20
$browseOU=New-Object Windows.Forms.Button
$browseOU.Text='Browse / Refresh OUs'
$browseOU.Location=New-Object Drawing.Point(765,3)
$browseOU.Size=New-Object Drawing.Size(150,28)
$script:controls.OU.Parent.Controls.Add($browseOU)
Field Mimic 'Copy from user (UPN or logon name)'
Field CopyAttributes 'Copy work attributes / manager' Check
$script:controls.CopyAttributes.Checked=$true
Field CopyGroups 'Copy direct AD groups (merge defaults)' Check
$script:controls.CopyGroups.Checked=$true

Field First 'First name'
Field Last 'Last name'
Field Full 'Full name / Display name'
Field Logon 'User logon name (sAMAccountName)'
Field Password 'Password (16 characters)'
$script:controls.Password.UseSystemPasswordChar=$true
Field Show 'Show password' Check
$script:controls.Show.Add_CheckedChanged({$script:controls.Password.UseSystemPasswordChar=-not $script:controls.Show.Checked})
Field Change 'Change password at next logon' Check
$script:controls.Change.Checked=$true
foreach ($item in @(@('Description','Description'),@('Office','Office'),@('Email','Email / primary SMTP'),@('Title','Job title'),@('Department','Department'),@('Company','Company'),@('Manager','Manager UPN / logon / DN (optional)'),@('Proxy','proxyAddresses (one per line)'))) { Field $item[0] $item[1] }
# Normalize typed and pasted email text to lowercase before SMTP synchronization.
$script:controls.Email.CharacterCasing=[Windows.Forms.CharacterCasing]::Lower
$script:controls.Proxy.Multiline=$true; $script:controls.Proxy.Height=75; $script:controls.Proxy.Parent.Height=83
Field Groups 'AD groups (name / DN; one per line)'
$script:controls.Groups.Multiline=$true
$script:controls.Groups.ScrollBars='Vertical'
$script:controls.Groups.Height=100
$script:controls.Groups.Parent.Height=108

Field Tenant 'Entra tenant ID (GUID; required for MFA)'
$script:controls.Tenant.Text='db945de8-1bf0-4bd1-af88-9825a67e3325'
Field MFA 'Enable per-user MFA after sync' Check
$script:controls.MFA.Checked=$true
Field License 'Office license details (report only)'
$script:controls.License.Text='Microsoft 365 A1 for faculty (WEB ONLY ACCESS: https://www.office.com/) OR Mobile Apps / Microsoft 365 A3 for faculty'
Field Notes 'Notes (report only)'
$script:controls.Notes.Text='N/A'
$buttons=New-Object Windows.Forms.FlowLayoutPanel; $buttons.Size=New-Object Drawing.Size(940,82)
$panel.Controls.Add($buttons)
function Button([string]$text) { $b=New-Object Windows.Forms.Button; $b.Text=$text; $b.AutoSize=$true; $b.Height=32; $buttons.Controls.Add($b); return $b }
$copy=Button 'Load / Copy user'; $generate=Button 'Generate password'; $create=Button 'Create AD account'; $enable=Button 'Check sync / Enable MFA'; $export=Button 'Export text report'; $another=Button 'Create another user'; $another.Enabled=$false
$enable.Enabled=$false; $export.Enabled=$false
$status=New-Object Windows.Forms.TextBox; $status.Multiline=$true; $status.ReadOnly=$true; $status.ScrollBars='Vertical'; $status.Size=New-Object Drawing.Size(930,145); $panel.Controls.Add($status)
function Log([string]$s) { $status.AppendText("$(Get-Date -Format HH:mm:ss)  $s`r`n") }
function Load-OUs {
    $old=[string]$script:controls.OU.SelectedItem
    $script:controls.OU.BeginUpdate()
    try {
        $script:controls.OU.Items.Clear()
        $ous=@(Get-ADOrganizationalUnit -Filter * -SearchBase $script:domain.DistinguishedName -SearchScope Subtree -Server $script:dc | Sort-Object DistinguishedName)
        foreach ($ou in $ous) { [void]$script:controls.OU.Items.Add([string]$ou.DistinguishedName) }
        [void]$script:controls.OU.Items.Add([string]$script:domain.UsersContainer)
        if ($old -and $script:controls.OU.Items.Contains($old)) { $script:controls.OU.SelectedItem=$old }
        else { $script:controls.OU.SelectedIndex=-1 }
        Log "Loaded $($ous.Count) OUs from $($script:dc). Select a destination or use Browse / Refresh OUs."
    } finally { $script:controls.OU.EndUpdate() }
}
function Set-DefaultGroups {
    if ($script:controls.NoGroups.Checked) {
        $script:controls.Groups.Clear()
        $script:previousDefaults=@()
        return
    }
    $existing=@($script:controls.Groups.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notin $script:previousDefaults })
    $suffix=[string]$script:controls.Suffix.SelectedItem
    $new=@($script:defaultGroups[$suffix.ToLowerInvariant()] | Where-Object { $_ })
    $script:controls.Groups.Lines=[string[]]@(($new+$existing) | Select-Object -Unique)
    $script:previousDefaults=$new
    Log "Default groups loaded for $suffix. Review and edit the group list before creation."
}
function Refresh-Domain {
    $selected=[string]$script:controls.Domain.SelectedItem
    $found=Get-ADDomainController -Discover -DomainName $selected -Writable
    $script:dc=[string](@($found.HostName)[0])
    $script:domain=Get-ADDomain -Server $script:dc
    Load-OUs
    if (-not $script:controls.NoGroups.Checked -and $script:controls.Suffix.Items.Contains($selected)) { $script:controls.Suffix.SelectedItem=$selected }
    $create.Enabled=$true
}
$browseOU.Add_Click({
    try {
        Load-OUs
        $picker=New-Object Windows.Forms.Form
        $picker.Text="Select destination - $($script:domain.DNSRoot)"
        $picker.Size=New-Object Drawing.Size(950,550)
        $picker.StartPosition='CenterParent'
        $search=New-Object Windows.Forms.TextBox; $search.Dock='Top'
        $list=New-Object Windows.Forms.ListBox; $list.Dock='Fill'; $list.HorizontalScrollbar=$true
        $ok=New-Object Windows.Forms.Button; $ok.Text='Use selected OU'; $ok.Dock='Bottom'; $ok.Height=36
        $picker.Controls.Add($list); $picker.Controls.Add($search); $picker.Controls.Add($ok)
        $items=@($script:controls.OU.Items | ForEach-Object { [string]$_ })
        foreach ($item in $items) { [void]$list.Items.Add($item) }
        $search.Add_TextChanged({
            $list.BeginUpdate(); $list.Items.Clear()
            foreach ($item in $items) { if ($item.IndexOf($search.Text,[StringComparison]::OrdinalIgnoreCase) -ge 0) { [void]$list.Items.Add($item) } }
            $list.EndUpdate()
        })
        $choose={ if ($list.SelectedItem) { $script:controls.OU.SelectedItem=[string]$list.SelectedItem; $picker.DialogResult='OK'; $picker.Close() } }
        $ok.Add_Click($choose); $list.Add_DoubleClick($choose)
        [void]$picker.ShowDialog($f); $picker.Dispose()
    } catch { Log "OU lookup failed: $($_.Exception.Message)"; [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'OU lookup') }
})
$forest=Get-ADForest
foreach ($d in $forest.Domains) { [void]$script:controls.Domain.Items.Add($d) }
foreach ($s in (@($forest.Domains)+@($forest.UPNSuffixes) | Sort-Object -Unique)) { [void]$script:controls.Suffix.Items.Add($s) }
$script:controls.Domain.Add_SelectedIndexChanged({ try { Refresh-Domain } catch { Log "Domain discovery failed: $($_.Exception.Message)"; $create.Enabled=$false } })
$script:controls.Domain.SelectedItem=(Get-ADDomain).DNSRoot
if ($script:controls.Domain.SelectedIndex -lt 0) {$script:controls.Domain.SelectedIndex=0}
$script:controls.Suffix.Add_SelectedIndexChanged({Set-DefaultGroups})
$script:controls.Suffix.SelectedItem=$script:controls.Domain.SelectedItem
Set-DefaultGroups
$script:controls.NoGroups.Add_CheckedChanged({
    if ($script:controls.NoGroups.Checked) {
        $charisSuffix=@($script:controls.Suffix.Items | Where-Object { [string]$_ -ieq 'awmcharis.com' })
        if ($charisSuffix.Count -ne 1) {
            [void][Windows.Forms.MessageBox]::Show('AWMCharis.com is not an available UPN suffix in this AD forest. Configure the suffix in AD before using this option.','UPN suffix unavailable')
            $script:controls.NoGroups.Checked=$false
            return
        }
        $script:groupsBeforeNoGroups=[string[]]@($script:controls.Groups.Lines)
        $script:copyGroupsBeforeNoGroups=$script:controls.CopyGroups.Checked
        $script:controls.Suffix.SelectedItem=$charisSuffix[0]
        $script:controls.Suffix.Enabled=$false
        $script:controls.Groups.Clear()
        $script:controls.Groups.Enabled=$false
        $script:controls.CopyGroups.Checked=$false
        $script:controls.CopyGroups.Enabled=$false
        Log 'Intern / World Outreach selected: AWMCharis.com UPN; no optional AD groups will be assigned.'
    } else {
        $script:controls.Suffix.Enabled=$true
        $script:controls.Groups.Enabled=$true
        $script:controls.CopyGroups.Enabled=$true
        $script:controls.CopyGroups.Checked=$script:copyGroupsBeforeNoGroups
        # Resume the standard defaults for the current UPN suffix.
        $script:controls.Groups.Lines=[string[]]@($script:groupsBeforeNoGroups | Where-Object { $_ -notin $script:defaultGroups['awmi.net'] -and $_ -notin $script:defaultGroups['awmcharis.com'] })
        Set-DefaultGroups
        Log 'Standard group assignment restored. Review the group list.'
    }
})
$generate.Add_Click({Generate-Password})
function Escape-Ldap([string]$s) { $s.Replace('\','\5c').Replace('*','\2a').Replace('(','\28').Replace(')','\29').Replace([string][char]0,'\00') }
function Server-ForDN([string]$dn) {
    $owners=@($forest.Domains | Where-Object {
        $base=(@($_ -split '\.' | ForEach-Object {'DC='+$_}) -join ',')
        $dn.EndsWith(','+$base,[StringComparison]::OrdinalIgnoreCase)
    } | Sort-Object Length -Descending)
    if ($owners.Count -eq 0) {throw "Cannot locate the forest domain for $dn"}
    return [string]$owners[0]
}
function Resolve-Group([string]$identity) {
    if ($identity -match '^(?i)CN=') {
        $server=Server-ForDN $identity
        $g=Get-ADGroup -Identity $identity -Server $server -ErrorAction Stop
        $script:groupServers[$g.DistinguishedName]=$server
        return $g
    }

    $found=@(foreach ($d in $forest.Domains) {
        Get-ADGroup -LDAPFilter "(|(distinguishedName=$(Escape-Ldap $identity))(sAMAccountName=$(Escape-Ldap $identity))(name=$(Escape-Ldap $identity)))" -Server $d | ForEach-Object {
            $script:groupServers[$_.DistinguishedName]=[string]$d
            $_
        }
    })
    if ($found.Count -ne 1) { throw "Group '$identity' resolved to $($found.Count) matches. Use its exact distinguished name if ambiguous." }
    return $found[0]
}
$copy.Add_Click({
    try {
        $script:copyIdentity=''
        $identity=$script:controls.Mimic.Text.Trim()
        if (-not $identity) { throw 'Enter the source user UPN or logon name.' }
        $matches=@(foreach ($d in $forest.Domains) {
            Get-ADUser -LDAPFilter "(|(userPrincipalName=$(Escape-Ldap $identity))(sAMAccountName=$(Escape-Ldap $identity)))" -Server $d -Properties Description,Office,Title,Department,Company,Manager,memberOf
        })
        if ($matches.Count -ne 1) { throw 'Source user must resolve to exactly one AD user. Use the UPN.' }
        $source=$matches[0]
        $toCopy=@()
        if ($script:controls.CopyGroups.Checked -and -not $script:controls.NoGroups.Checked) {
            # memberOf gives direct memberships; nested memberships and primary group are not copied.
            foreach ($dn in @($source.memberOf)) { $g=Resolve-Group ([string]$dn); $toCopy += [string]$g.DistinguishedName }
        }
        $managerIdentity=''
        if ($script:controls.CopyAttributes.Checked -and $source.Manager) {
            $managerServer=Server-ForDN ([string]$source.Manager)
            $managerObject=Get-ADUser -Identity $source.Manager -Server $managerServer -ErrorAction Stop
            $managerIdentity=[string]$managerObject.DistinguishedName
        }
        $summary="Copy from $($source.UserPrincipalName)?`r`nWork attributes: $($script:controls.CopyAttributes.Checked)`r`nDirect groups: $($toCopy.Count)`r`n`r`n$($toCopy -join "`r`n")`r`n`r`nManager: $managerIdentity`r`nDepartment: $($source.Department)`r`nJob title: $($source.Title)`r`nEntra administrator roles are not copied.`r`nReview the resulting fields and group list before creating the account."
        if ([Windows.Forms.MessageBox]::Show($summary,'Review source account','YesNo','Question') -ne 'Yes') { return }
        if ($script:controls.CopyAttributes.Checked) {
            foreach ($key in @('Description','Office','Title','Department','Company')) { $script:controls[$key].Text=[string]$source.$key }
            $script:controls.Manager.Text=$managerIdentity
        }
        if ($script:controls.CopyGroups.Checked -and -not $script:controls.NoGroups.Checked) {
            $combined=@($script:controls.Groups.Lines | Where-Object { $_.Trim() }) + $toCopy
            $script:controls.Groups.Lines=[string[]]@($combined | Select-Object -Unique)
        }
        $script:copyIdentity=$identity
        $script:mimic=[string]$source.UserPrincipalName
        if (-not $script:mimic) { $script:mimic=[string]$source.SamAccountName }
        Log "Copied from $script:mimic. Review groups (including privileged groups) and attributes."
    } catch { Log "Copy failed: $($_.Exception.Message)"; [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Copy user') }
})
$create.Add_Click({
    $create.Enabled=$false
    try {
        if ($script:controls.Mimic.Text.Trim() -and $script:controls.Mimic.Text.Trim() -ine $script:copyIdentity) {
            throw 'Click Load / Copy user and confirm a successful copy before creating this account. Clear Copy from user to create without copying.'
        }
        foreach ($k in @('First','Last','Full','Logon','Email','Password')) { if ([string]::IsNullOrWhiteSpace($script:controls[$k].Text)) {throw "$k is required."} }
        if (-not $script:controls.OU.SelectedItem) { throw 'Select the destination OU using the dropdown or Browse / Refresh OUs.' }
        if (-not $script:controls.Suffix.SelectedItem) { throw 'Select a user logon domain.' }
        # Revalidate the selected location against the selected AD domain.
        $destination=Get-ADObject -Identity ([string]$script:controls.OU.SelectedItem) -Server $script:dc
        if ($destination.ObjectClass -notin @('organizationalUnit','container')) { throw 'Destination is not an OU or container.' }
        $sam=$script:controls.Logon.Text.Trim()
        if ($sam.Length -gt 20 -or $sam -match '["/\\\[\]:;|=,+*?<>@\s]' -or $sam -match '^\.|\.$') {throw 'Logon name must be 1-20 characters, with no spaces, @ or invalid AD characters.'}
        $p=$script:controls.Password.Text
        $problem=Password-Error $p; if ($problem) {throw $problem}
        $upn="$sam@$($script:controls.Suffix.SelectedItem)"
        $mail=$script:controls.Email.Text.Trim().ToLowerInvariant()
        if ($mail -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') {throw 'Enter a valid email address.'}
        if ($script:controls.MFA.Checked) { $tid=[guid]::Empty; if (-not [guid]::TryParse($script:controls.Tenant.Text.Trim(),[ref]$tid)) {throw 'Enter the Entra tenant GUID before creating an account with MFA selected.'} }
        $proxies=@($script:controls.Proxy.Lines | ForEach-Object {$_.Trim()} | Where-Object {$_})
        if ($proxies.Count -eq 0) {$proxies=@("SMTP:$mail")}
        if (@($proxies | Where-Object {$_ -cmatch '^SMTP:'}).Count -ne 1 -or -not ($proxies -ccontains "SMTP:$mail")) {throw 'Use exactly one uppercase SMTP: primary address matching Email. Use lowercase smtp: for aliases.'}
        if (@($proxies | Sort-Object -Unique).Count -ne $proxies.Count) {throw 'Duplicate proxy addresses.'}
        foreach ($a in $proxies) {if ($a -notmatch '^[A-Za-z][A-Za-z0-9]*:.+\S$') {throw "Invalid proxy address: $a"}}
        # Search every forest domain for conflicting UPN and mail/proxy addresses.
        foreach ($d in $forest.Domains) {
            $parts=@("(userPrincipalName=$(Escape-Ldap $upn))","(mail=$(Escape-Ldap $mail))")
            foreach ($a in $proxies) {$parts+="(proxyAddresses=$(Escape-Ldap $a))"}
            if (Get-ADObject -LDAPFilter "(|$($parts -join ''))" -Server $d) {throw "UPN/email/proxy conflict found in $d."}
        }
        if (Get-ADUser -LDAPFilter "(sAMAccountName=$(Escape-Ldap $sam))" -Server $script:dc) {throw 'That logon name already exists.'}
        $manager=$null
        if ($script:controls.Manager.Text.Trim()) {
            $m=$script:controls.Manager.Text.Trim()
            if ($m -match '^(?i)CN=') {
                $matches=@(Get-ADUser -Identity $m -Server (Server-ForDN $m) -ErrorAction Stop)
            } else {
                $matches=@(foreach ($d in $forest.Domains) {Get-ADUser -LDAPFilter "(|(userPrincipalName=$(Escape-Ldap $m))(sAMAccountName=$(Escape-Ldap $m)))" -Server $d})
            }
            if ($matches.Count -ne 1) {throw 'Manager must resolve to exactly one AD user. Use their UPN.'}
            $manager=$matches[0].DistinguishedName
        }
        if ($script:controls.NoGroups.Checked -and [string]$script:controls.Suffix.SelectedItem -ine 'awmcharis.com') { throw 'Intern / World Outreach accounts must use the AWMCharis.com UPN suffix.' }
        $groups=@()
        if (-not $script:controls.NoGroups.Checked) {
        $groups=@(foreach ($g in ($script:controls.Groups.Text -split '[\r\n]+' | ForEach-Object {$_.Trim()} | Where-Object {$_} | Select-Object -Unique)) {Resolve-Group $g})
        $groups=@($groups | Sort-Object DistinguishedName -Unique)
        }
        $argsAD=@{Name=$script:controls.Full.Text.Trim(); DisplayName=$script:controls.Full.Text.Trim(); GivenName=$script:controls.First.Text.Trim(); Surname=$script:controls.Last.Text.Trim(); SamAccountName=$sam; UserPrincipalName=$upn; Path=$script:controls.OU.SelectedItem.ToString(); Server=$script:dc; AccountPassword=(ConvertTo-SecureString $p -AsPlainText -Force); Enabled=$false; ChangePasswordAtLogon=$script:controls.Change.Checked; OtherAttributes=@{proxyAddresses=[string[]]$proxies}; EmailAddress=$mail; PassThru=$true}
        foreach ($pair in @(@('Description','Description'),@('Office','Office'),@('Title','Title'),@('Department','Department'),@('Company','Company'))) {if ($script:controls[$pair[0]].Text.Trim()) {$argsAD[$pair[1]]=$script:controls[$pair[0]].Text.Trim()}}
        if ($manager) {$argsAD.Manager=$manager}
        $answer=[Windows.Forms.MessageBox]::Show("Create $upn in $($argsAD.Path)?`r`nGroups: $($groups.Name -join '; ')`r`nCopied from: $script:mimic`r`nMFA will require the account to sync first.",'Confirm account creation','YesNo','Question')
        if ($answer -ne 'Yes') {return}
        # Create disabled first. A failure during enablement is retained and reported.
        $u=New-ADUser @argsAD
        $script:created=@{Guid=$u.ObjectGUID; SID=$u.SID.Value; UPN=$upn; Sam=$sam; Password=$p; DC=$script:dc; Tenant=$script:controls.Tenant.Text.Trim(); Change=$script:controls.Change.Checked; MFARequested=$script:controls.MFA.Checked; License=$script:controls.License.Text; Notes=$script:controls.Notes.Text; Mimic=$script:mimic; NoGroups=$script:controls.NoGroups.Checked}
        $export.Enabled=$true; $enable.Enabled=$true
        foreach ($c in $script:controls.Values) {$c.Enabled=$false}; $generate.Enabled=$false; $copy.Enabled=$false; $browseOU.Enabled=$false
        # Tenant ID can be supplied or corrected for the separate cloud step.
        $script:controls.Tenant.Enabled=$true
        Log "AD object created: $upn"
        # Defense in depth: never assign optional groups for Intern / World Outreach.
        if ($script:created.NoGroups) { $groups=@() }
        foreach ($g in $groups) {try {Add-ADGroupMember -Identity $g.DistinguishedName -Members $u.DistinguishedName -Server $script:groupServers[$g.DistinguishedName] -ErrorAction Stop
            $checkGroup=Get-ADGroup -Identity $g.DistinguishedName -Server $script:groupServers[$g.DistinguishedName] -Properties member -ErrorAction Stop
            if (@($checkGroup.member) -notcontains $u.DistinguishedName) {throw 'Membership was not confirmed by readback.'}
            Log "Added group: $($g.Name)"} catch {$script:issues.Add("Group $($g.Name): $($_.Exception.Message)"); Log $script:issues[$script:issues.Count-1]}}
        $saved=Get-ADUser -Identity $u.ObjectGUID -Server $script:dc -Properties Description,Office,Title,Department,Company,Manager
        foreach ($attribute in @('Description','Office','Title','Department','Company','Manager')) {
            if ($argsAD.ContainsKey($attribute) -and [string]$saved.$attribute -ine [string]$argsAD[$attribute]) {
                $script:issues.Add("Attribute $attribute did not match the requested value after creation.")
            }
        }
        if ($script:issues.Count -gt 0) {
            [void][Windows.Forms.MessageBox]::Show(($script:issues -join "`r`n"),'Account created with incomplete operations','OK','Warning')
        }
        Enable-ADAccount -Identity $u.ObjectGUID -Server $script:dc
        Log 'AD account enabled. Attributes saved. Check sync / Enable MFA is now available. This tool does not wait automatically: allow directory synchronization, then click the button; retry if the cloud account is not yet available.'
        if (-not $script:created.MFARequested) {$script:mfa='Not requested / not verified'}
    } catch {if ($script:created) {$script:issues.Add($_.Exception.Message)}; Log "FAILED: $($_.Exception.Message)"; [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'Account creation')}
    finally {
        $create.Enabled=($null -eq $script:created)
        $enable.Enabled=($null -ne $script:created)
        if ($script:created) {$export.Enabled=$true; $another.Enabled=$true; $script:controls.Tenant.Enabled=$true}
    }
})
$enable.Add_Click({
    $enable.Enabled=$false; $another.Enabled=$false; $export.Enabled=$false; $f.UseWaitCursor=$true
    $mfaCompleted=$false
    try {
        if (-not $script:created) {throw 'Create the AD account first.'}
        $tid=[guid]::Empty
        if (-not [guid]::TryParse($script:controls.Tenant.Text.Trim(),[ref]$tid)) {throw 'Enter a valid Entra tenant ID (GUID), then click Check sync / Enable MFA again.'}
        $script:created.Tenant=$tid.ToString()
        $script:created.MFARequested=$true
        Log "Checking Entra for $($script:created.UPN). Sign in when prompted. If synchronization is pending, retry this button later."
        Import-Module Microsoft.Graph.Authentication
        $ctx=Get-MgContext
        $needed=@('User.Read.All','Policy.ReadWrite.AuthenticationMethod')
        if (-not $ctx -or $ctx.TenantId -ne $script:created.Tenant -or @($needed | Where-Object {$_ -notin $ctx.Scopes}).Count -gt 0) {
            if ($ctx) {Disconnect-MgGraph | Out-Null}
            Connect-MgGraph -TenantId $script:created.Tenant -Scopes $needed -ContextScope Process -NoWelcome
        }
        if ((Get-MgContext).TenantId -ne $script:created.Tenant) {throw 'Graph tenant mismatch. No MFA change was made.'}
        $id=[uri]::EscapeDataString($script:created.UPN)
        $cloud=Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/${id}?`$select=id,userPrincipalName,onPremisesSyncEnabled,onPremisesSecurityIdentifier"
        if (-not $cloud.onPremisesSyncEnabled -or $cloud.onPremisesSecurityIdentifier -ne $script:created.SID) {throw 'Cloud identity is not the matching synced AD account. No MFA change was made.'}
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/beta/users/$($cloud.id)/authentication/requirements" -Body '{"perUserMfaState":"enabled"}' -ContentType 'application/json' | Out-Null
        $verify=Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($cloud.id)/authentication/requirements"
        if ($verify.perUserMfaState -notin @('enabled','enforced')) {throw 'MFA write was not confirmed by readback.'}
        $script:mfa="Enabled (verified per-user MFA state: $($verify.perUserMfaState); will be prompted to configure on sign-in, https://aka.ms/MFASetup/)"
        Log $script:mfa
        $mfaCompleted=$true
    } catch {$script:mfa="Pending / failed - $($_.Exception.Message)"; Log 'MFA incomplete. If the user is not found, wait for sync and retry.'; Log $_.Exception.Message}
    finally {$enable.Enabled=($null -ne $script:created); $another.Enabled=($null -ne $script:created); $export.Enabled=($null -ne $script:created); $f.UseWaitCursor=$false}
    if ($mfaCompleted) { Export-AccountReport }
})
function Export-AccountReport {
    try {
        if (-not $script:created) {throw 'Create an account before exporting a report.'}
        $u=Get-ADUser -Identity $script:created.Guid -Server $script:created.DC -Properties DisplayName,Description,Office,mail,Title,Department,Company,Manager,proxyAddresses,pwdLastSet,Enabled
        $groupNames=@(Get-ADPrincipalGroupMembership -Identity $u -Server $script:created.DC | Sort-Object Name | Select-Object -ExpandProperty Name)
        $complete=if ($u.Enabled -and $script:mfa -like 'Enabled*' -and $script:issues.Count -eq 0) {'The following Account creation is complete, details are below'} else {'AD account creation details are below. Review pending actions and status before handing off.'}
        $mfaReport=if ($script:mfa -like 'Enabled*') {'Enabled (Will be prompted to configure on first login, https://aka.ms/MFASetup/)'} else {$script:mfa}
        $txt=@"
$complete

Do not pass/forward this information to the Employee/Contractor/Intern, if Hardware is being configured an Onsite Tech will reach out to set up Delivery. Although your account creation is complete, other members of the IT Team may still working on additional aspects of your request. We’ll provide an update once everything has been completed.

• Computer/Network: $(($u.UserPrincipalName -split '@',2)[0])
• Display Name: $($u.DisplayName)
• Email: $($u.mail)
• Password: $($script:created.Password) (Will be prompted to change this on first login)
• 2FA: $mfaReport
• Office License(s): $($script:created.License)
• Groups: $($groupNames -join '; ')
• 365 Role: N/A
• Account Mimic'd: $($script:created.Mimic)
• Additional Mailbox/Distribution Group Access: N/A
• Additional Azure Applications: N/A
• Additional SharePoint/OneDrive Access: N/A
• Notes: $($script:created.Notes)

Please make sure to have users sign into their accounts at least once every 30 days. Inactive accounts are monitored and if an extended period of inactivity is seen, the account(s) may be disabled/removed/deleted from the AWM Tenant.

Per AWM Policy, Passwords need to be 16 Characters with a combination of UPPERCASE, lowercase, Numbers (1234567890) and at least one special character (@#$%^&*!). Passwords cannot include any part of the account/username or consecutive numbers. Use a different password than other accounts (Professional or Personal) you use. Do NOT re-use Passwords from other accounts. Try to stay away from letters that resemble numbers and vice versa.

AD ATTRIBUTE DETAILS
User principal name: $($u.UserPrincipalName)
Account enabled: $($u.Enabled)
Distinguished name: $($u.DistinguishedName)
First name: $($u.GivenName)
Last name: $($u.Surname)
Description: $($u.Description)
Office: $($u.Office)
Job title: $($u.Title)
Department: $($u.Department)
Company: $($u.Company)
Manager: $($u.Manager)
proxyAddresses: $($u.proxyAddresses -join '; ')
Intern / World Outreach (no optional groups): $($script:created.NoGroups)
Created object GUID: $($u.ObjectGUID)
Report generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')
Pending / operation errors: $($script:issues -join '; ')
License details are operator-entered; this tool does not assign or verify licenses.
"@
        if ($complete -notlike 'The following*') {$txt=$txt.Replace('Although your account creation is complete,','The AD account has been created; review pending actions below. Also,') }
        $dlg=New-Object Windows.Forms.SaveFileDialog; $dlg.Filter='Text files (*.txt)|*.txt'; $dlg.FileName="Account-Creation-$($script:created.Sam)-$(Get-Date -Format yyyyMMdd-HHmmss).txt"
        if ($dlg.ShowDialog() -eq 'OK') {[IO.File]::WriteAllText($dlg.FileName,$txt,(New-Object Text.UTF8Encoding($true))); $script:reportSaved=$true; Log "Report exported: $($dlg.FileName)"}
    } catch {Log "Report failed: $($_.Exception.Message)"}
}
$export.Add_Click({ Export-AccountReport })
$another.Add_Click({
    if (-not $script:created) {return}
    if (-not $script:reportSaved) {
        $answer=[Windows.Forms.MessageBox]::Show('The current report has not been saved. Start another user and discard the current account details and password from this window? The AD account will remain unchanged.','Unsaved account report','YesNo','Warning')
        if ($answer -ne 'Yes') {return}
    }
    # Retain domain, OU and tenant selections; clear all user-specific state.
    $script:created=$null
    $script:reportSaved=$false
    $script:mfa='Pending - not yet enabled'
    $script:mimic='N/A'
    $script:copyIdentity=''
    $script:issues.Clear()
    $script:groupServers.Clear()
    foreach ($c in $script:controls.Values) {$c.Enabled=$true}
    $script:controls.NoGroups.Checked=$false
    foreach ($key in @('Mimic','First','Last','Full','Logon','Password','Description','Office','Email','Title','Department','Company','Manager','Proxy','Groups')) {$script:controls[$key].Clear()}
    $script:controls.Proxy.Text='SMTP:'
    $script:previousDefaults=@()
    $script:groupsBeforeNoGroups=@()
    $script:copyGroupsBeforeNoGroups=$true
    $script:controls.CopyAttributes.Checked=$true
    $script:controls.CopyGroups.Checked=$true
    $script:controls.Change.Checked=$true
    $script:controls.MFA.Checked=$true
    $script:controls.Show.Checked=$false
    $script:controls.License.Text='Microsoft 365 A1 for faculty (WEB ONLY ACCESS: https://www.office.com/) OR Mobile Apps / Microsoft 365 A3 for faculty'
    $script:controls.Notes.Text='N/A'
    Set-DefaultGroups
    Generate-Password
    $generate.Enabled=$true; $copy.Enabled=$true; $browseOU.Enabled=$true
    $create.Enabled=$true; $enable.Enabled=$false; $export.Enabled=$false; $another.Enabled=$false
    Log 'Ready for another user. Domain, OU and tenant selections retained; review them before creation.'
    $panel.ScrollControlIntoView($script:controls.First.Parent)
    [void]$script:controls.First.Focus()
})
# Keep the uppercase primary SMTP address synchronized, retaining aliases.
$script:controls.Email.Add_TextChanged({
    if ($script:created) {return}
    $aliases=@($script:controls.Proxy.Lines | Where-Object {$_ -and $_ -cnotmatch '^SMTP:'})
    $script:controls.Proxy.Lines=[string[]](@('SMTP:'+$script:controls.Email.Text.Trim())+$aliases)
})
$script:controls.Proxy.Text='SMTP:'

# Help with common defaults while allowing all fields to be edited.
$fillDefaults = {
    if (-not $script:controls.Full.Text.Trim()) { $script:controls.Full.Text = ("$($script:controls.First.Text.Trim()) $($script:controls.Last.Text.Trim())").Trim() }
    if (-not $script:controls.Email.Text.Trim() -and $script:controls.Logon.Text.Trim()) { $script:controls.Email.Text = "$($script:controls.Logon.Text.Trim())@$($script:controls.Suffix.SelectedItem)" }
}
$script:controls.Logon.Add_Leave($fillDefaults)
$f.Add_FormClosing({param($sender,$e) if ($script:created) {if ([Windows.Forms.MessageBox]::Show('Have you exported the report? Closing discards the password held in this window.','Close','YesNo','Question') -ne 'Yes') {$e.Cancel=$true}}})
Log 'Default groups follow the user logon UPN suffix. Use Browse / Refresh OUs to choose a location in sync scope. Enter details, then generate a password. Licenses are report-only.'
Generate-Password
[void]$f.ShowDialog()
$script:rng.Dispose()
