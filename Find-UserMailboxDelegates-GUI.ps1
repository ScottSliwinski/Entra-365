#Requires -Version 5.1
<#
Read-only Exchange Online mailbox delegation report for hybrid organizations.
Run in Windows PowerShell 5.1 or PowerShell 7 on Windows with -STA.
Prerequisite: Install-Module ExchangeOnlineManagement -Scope CurrentUser
Domain filters use MAILBOX PRIMARY SMTP ADDRESS, not delegate UPN.
Includes user/shared/room/equipment mailboxes hosted in Exchange Online.
Does not scan on-premises mailbox databases, folder/calendar permissions,
Microsoft 365 group membership, or administrative role-based access.
Group results describe matching assignments, not a calculated effective-access verdict.
Nested mail-enabled security groups are resolved using Exchange's synced view.
Unsynced/non-mail-enabled groups and unresolved trustees are reported as warnings.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne 'Win32NT') { throw 'This GUI requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Start with powershell.exe -STA -File .\Find-UserMailboxDelegates-GUI.ps1 (or pwsh -STA -File ...).'
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$script:rows = New-Object 'System.Collections.Generic.List[object]'
$script:issues = New-Object 'System.Collections.Generic.List[object]'
$script:cache = @{}
$script:members = @{}
$script:busy = $false
$script:cancel = $false
$script:connected = $false

function Issue([string]$scope, [string]$message) {
    $script:issues.Add([pscustomobject]@{ Scope=$scope; Message=$message })
    $log.AppendText("[$scope] $message`r`n")
}
function Resolve-Trustee([string]$identity) {
    if ($script:cache.ContainsKey($identity)) { return $script:cache[$identity] }
    try {
        $r = Get-Recipient -Identity $identity -ErrorAction Stop
        $script:cache[$identity] = $r
        return $r
    } catch {
        Issue $identity "Cannot resolve trustee: $($_.Exception.Message)"
        $script:cache[$identity] = $null
        return $null
    }
}
function Same-Recipient($a, $b) {
    foreach ($p in @('ExternalDirectoryObjectId','Guid','Sid','DistinguishedName','PrimarySmtpAddress')) {
        $av = [string]$a.$p; $bv = [string]$b.$p
        if ($av -and $bv -and $av -eq $bv) { return $true }
    }
    return $false
}
function Group-ContainsUser($group, $user, [hashtable]$visited) {
    $key = [string]$group.DistinguishedName
    if (-not $key) { $key = [string]$group.Identity }
    if ($visited.ContainsKey($key)) { return $false }
    $visited[$key] = $true
    if (-not $script:members.ContainsKey($key)) {
        try {
            $script:members[$key] = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop)
        } catch {
            Issue ([string]$group.Identity) "Group membership could not be checked: $($_.Exception.Message)"
            $script:members[$key] = @()
        }
    }
    foreach ($member in $script:members[$key]) {
        if (Same-Recipient $member $user) { return $true }
        if ([string]$member.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {
            if (Group-ContainsUser $member $user $visited) { return $true }
        }
    }
    return $false
}
function Match-Trustee([string]$identity, $user) {
    if ($identity -match '^(NT AUTHORITY\\SELF|S-1-5-10)$') { return $null }
    foreach ($p in @('UserPrincipalName','PrimarySmtpAddress','Sid','DistinguishedName','Identity','Guid')) {
        if ([string]$user.$p -and $identity -eq [string]$user.$p) {
            return [pscustomobject]@{ Via='Direct'; Trustee=$identity }
        }
    }
    $r = Resolve-Trustee $identity
    if (-not $r) { return $null }
    if (Same-Recipient $r $user) {
        return [pscustomobject]@{ Via='Direct'; Trustee=[string]$r.PrimarySmtpAddress }
    }
    if ($groups.Checked -and [string]$r.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {
        if (Group-ContainsUser $r $user @{}) {
            return [pscustomobject]@{ Via='Group (direct or nested membership)'; Trustee=[string]$r.PrimarySmtpAddress }
        }
    }
    return $null
}
function Add-Result($mailbox, [string]$permission, $match, [bool]$deny, [bool]$inherited) {
    $script:rows.Add([pscustomobject]@{
        UserUPN=$script:scannedUPN
        MailboxName=[string]$mailbox.DisplayName
        MailboxEmail=[string]$mailbox.PrimarySmtpAddress
        MailboxType=[string]$mailbox.RecipientTypeDetails
        Permission=$permission
        Assignment=$(if ($deny) { 'Deny' } else { 'Allow' })
        Via=$match.Via
        Trustee=$match.Trustee
        Inherited=$inherited
        ScanTime=$script:scanTime
    })
}

$form = New-Object Windows.Forms.Form
$form.Text = 'Mailbox Delegate Access - Hybrid / Exchange Online'
$form.Size = New-Object Drawing.Size(1150,790)
$form.MinimumSize = New-Object Drawing.Size(1000,700)
$form.StartPosition = 'CenterScreen'
function Label([string]$text, [int]$x, [int]$y, [int]$width) {
    $c = New-Object Windows.Forms.Label
    $c.Text=$text; $c.Location=New-Object Drawing.Point($x,$y); $c.Size=New-Object Drawing.Size($width,24)
    $form.Controls.Add($c)
}
function Button([string]$text, [int]$x, [int]$y, [int]$width) {
    $c=New-Object Windows.Forms.Button
    $c.Text=$text; $c.Location=New-Object Drawing.Point($x,$y); $c.Size=New-Object Drawing.Size($width,30)
    $form.Controls.Add($c); return $c
}
function Check([string]$text, [int]$x, [int]$y, [int]$width, [bool]$checked) {
    $c=New-Object Windows.Forms.CheckBox
    $c.Text=$text; $c.Location=New-Object Drawing.Point($x,$y); $c.Size=New-Object Drawing.Size($width,26); $c.Checked=$checked
    $form.Controls.Add($c); return $c
}
Label 'User UPN:' 15 18 90
$upn=New-Object Windows.Forms.TextBox
$upn.Location=New-Object Drawing.Point(110,15); $upn.Size=New-Object Drawing.Size(380,26); $form.Controls.Add($upn)
$connect=Button 'Connect Exchange Online' 510 12 210
$scan=Button 'Scan Mailboxes' 735 12 150
$stop=Button 'Cancel Scan' 900 12 125; $stop.Enabled=$false
Label 'Mailbox email domains to search:' 15 55 250
$awmi=Check 'AWMI.net' 270 51 130 $true
$charis=Check 'AWMCharis.com' 420 51 170 $true
$groups=Check 'Include access through security groups' 610 51 370 $true
Label 'Read-only. Cloud-hosted mailboxes. Domains match primary email address; the user UPN is not filtered.' 15 88 1090
Label 'Reports Full Access, Send As and Send on Behalf assignments. Calendar/folder permissions are outside this scan.' 15 112 1090
$progress=New-Object Windows.Forms.ProgressBar
$progress.Location=New-Object Drawing.Point(15,145); $progress.Size=New-Object Drawing.Size(1095,20)
$progress.Anchor='Top,Left,Right'; $form.Controls.Add($progress)
$status=New-Object Windows.Forms.Label
$status.Text='Connect to Exchange Online to begin.'; $status.Location=New-Object Drawing.Point(15,173); $status.Size=New-Object Drawing.Size(1095,30)
$status.Anchor='Top,Left,Right'; $form.Controls.Add($status)
$grid=New-Object Windows.Forms.DataGridView
$grid.Location=New-Object Drawing.Point(15,210); $grid.Size=New-Object Drawing.Size(1095,350)
$grid.Anchor='Top,Bottom,Left,Right'; $grid.ReadOnly=$true; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false
$grid.AutoSizeColumnsMode='DisplayedCells'; $grid.SelectionMode='FullRowSelect'; $grid.RowHeadersVisible=$false
$form.Controls.Add($grid)
$export=Button 'Export Results CSV' 15 570 180; $export.Anchor='Bottom,Left'; $export.Enabled=$false
$exportIssues=Button 'Export Scan Issues CSV' 210 570 200; $exportIssues.Anchor='Bottom,Left'; $exportIssues.Enabled=$false
$log=New-Object Windows.Forms.TextBox
$log.Multiline=$true; $log.ReadOnly=$true; $log.ScrollBars='Vertical'
$log.Location=New-Object Drawing.Point(15,610); $log.Size=New-Object Drawing.Size(1095,125); $log.Anchor='Bottom,Left,Right'
$form.Controls.Add($log)

$connect.Add_Click({
    try {
        $connect.Enabled=$false
        if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
            throw 'Install the module first: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
        }
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        $argsConnect=@{ ShowBanner=$false; ErrorAction='Stop' }
        if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) { $argsConnect.DisableWAM=$true }
        Connect-ExchangeOnline @argsConnect
        $script:connected=$true
        $status.Text='Connected. Enter a user UPN and select the mailbox domains.'
    } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Connection failed') | Out-Null }
    finally { $connect.Enabled=$true }
})
$stop.Add_Click({ $script:cancel=$true; $status.Text='Cancelling after the current mailbox request finishes...' })
$form.Add_FormClosing({ param($sender,$eventArgs) if ($script:busy) { $eventArgs.Cancel=$true; $script:cancel=$true } })
$scan.Add_Click({
    if (-not $script:connected) { [Windows.Forms.MessageBox]::Show('Connect to Exchange Online first.') | Out-Null; return }
    if ($upn.Text.Trim() -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { [Windows.Forms.MessageBox]::Show('Enter a valid user UPN.') | Out-Null; return }
    if (-not ($awmi.Checked -or $charis.Checked)) { [Windows.Forms.MessageBox]::Show('Select at least one mailbox domain.') | Out-Null; return }
    $script:busy=$true; $script:cancel=$false
    foreach ($c in @($connect,$scan,$upn,$awmi,$charis,$groups,$export,$exportIssues)) { $c.Enabled=$false }
    $stop.Enabled=$true
    $script:rows.Clear(); $script:issues.Clear(); $script:cache=@{}; $script:members=@{}
    $grid.DataSource=$null; $log.Clear(); $progress.Value=0
    $script:scannedUPN=$upn.Text.Trim(); $script:scanTime=(Get-Date).ToString('o')
    $completed=0; $total=0; $fatal=$false
    try {
        $status.Text='Resolving user and loading mailbox list...'; [Windows.Forms.Application]::DoEvents()
        # Resolve by the dedicated UPN selector. Get-User's returned UPN can be
        # absent/different in some hybrid recipient views; it is not our validator.
        $resolved=@(Get-EXORecipient -UserPrincipalName $script:scannedUPN -ErrorAction Stop)
        if ($resolved.Count -ne 1) { throw 'The UPN lookup did not return exactly one mail-enabled user in this Exchange Online tenant.' }
        $cloudUser=$resolved[0]
        $recipients=@(Get-Recipient -Identity $script:scannedUPN -ErrorAction Stop)
        if ($recipients.Count -ne 1) { throw 'The recipient lookup did not return exactly one user.' }
        $user=$recipients[0]
        if ([string]$user.RecipientTypeDetails -match 'Group|Contact') { throw 'The entered UPN resolves to a group/contact, not a user.' }
        $cloudId=[string]$cloudUser.ExternalDirectoryObjectId
        $recipientId=[string]$user.ExternalDirectoryObjectId
        if (-not $cloudId -or -not $recipientId -or $cloudId -ne $recipientId) {
            throw 'Exchange could not verify that the UPN and recipient lookups refer to the same account. Check the connected tenant and account synchronization.'
        }
        $user | Add-Member -NotePropertyName UserPrincipalName -NotePropertyValue $script:scannedUPN -Force
        $log.AppendText("Resolved user: $($user.DisplayName) | $script:scannedUPN | $($user.PrimarySmtpAddress)`r`n")
        $domains=@(); if ($awmi.Checked) { $domains+='awmi.net' }; if ($charis.Checked) { $domains+='awmcharis.com' }
        $mailboxes=@(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox -Properties GrantSendOnBehalfTo -ErrorAction Stop |
            Where-Object { (([string]$_.PrimarySmtpAddress -split '@')[-1]).ToLowerInvariant() -in $domains })
        $total=$mailboxes.Count; $progress.Maximum=[Math]::Max(1,$total)
        foreach ($mb in $mailboxes) {
            [Windows.Forms.Application]::DoEvents(); if ($script:cancel) { break }
            $id=[string]$mb.PrimarySmtpAddress
            $status.Text="Scanning $($completed+1) of $total : $id"
            try {
                foreach ($p in @(Get-EXOMailboxPermission -Identity $id -ErrorAction Stop)) {
                    if ('FullAccess' -in @($p.AccessRights | ForEach-Object { [string]$_ })) {
                        $match=Match-Trustee ([string]$p.User) $user
                        if ($match) { Add-Result $mb 'Full Access' $match ([bool]$p.Deny) ([bool]$p.IsInherited) }
                    }
                }
            } catch { Issue $id "Full Access check failed: $($_.Exception.Message)" }
            try {
                foreach ($p in @(Get-RecipientPermission -Identity $id -ErrorAction Stop)) {
                    if ('SendAs' -in @($p.AccessRights | ForEach-Object { [string]$_ })) {
                        $match=Match-Trustee ([string]$p.Trustee) $user
                        if ($match) { Add-Result $mb 'Send As' $match ([string]$p.AccessControlType -eq 'Deny') ([bool]$p.IsInherited) }
                    }
                }
            } catch { Issue $id "Send As check failed: $($_.Exception.Message)" }
            try {
                foreach ($delegate in @($mb.GrantSendOnBehalfTo)) {
                    if (-not $delegate) { continue }
                    $match=Match-Trustee ([string]$delegate) $user
                    if ($match) { Add-Result $mb 'Send on Behalf' $match $false $false }
                }
            } catch { Issue $id "Send on Behalf check failed: $($_.Exception.Message)" }
            $completed++; $progress.Value=$completed
        }
    } catch { $fatal=$true; Issue 'Scan' $_.Exception.Message }
    finally {
        $table=New-Object System.Data.DataTable
        foreach ($col in @('UserUPN','MailboxName','MailboxEmail','MailboxType','Permission','Assignment','Via','Trustee','Inherited','ScanTime')) { [void]$table.Columns.Add($col) }
        foreach ($r in ($script:rows | Sort-Object MailboxEmail,Permission,Trustee)) {
            $dr=$table.NewRow(); foreach ($col in $table.Columns) { $dr[$col.ColumnName]=[string]$r.($col.ColumnName) }; $table.Rows.Add($dr)
        }
        $grid.DataSource=$table
        $state=if ($fatal) { 'FAILED / PARTIAL' } elseif ($script:cancel) { 'CANCELLED / PARTIAL' } elseif ($script:issues.Count) { 'FINISHED WITH ISSUES / PARTIAL' } else { 'COMPLETE' }
        $status.Text="$state - $completed/$total mailboxes scanned; $($script:rows.Count) matching assignments; $($script:issues.Count) issues."
        $log.AppendText("$($status.Text)`r`nGroup matches reflect Exchange membership; deny entries can affect effective access.`r`n")
        $script:busy=$false; $stop.Enabled=$false
        foreach ($c in @($connect,$scan,$upn,$awmi,$charis,$groups)) { $c.Enabled=$true }
        $export.Enabled=($script:rows.Count -gt 0); $exportIssues.Enabled=($script:issues.Count -gt 0)
    }
})
function Export-Rows($items, [string]$suffix) {
    $dialog=New-Object Windows.Forms.SaveFileDialog
    $dialog.Filter='CSV files (*.csv)|*.csv'; $dialog.FileName=($script:scannedUPN -replace '[^a-zA-Z0-9._-]','_')+"-$suffix-"+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.csv'
    if ($dialog.ShowDialog() -eq 'OK') {
        try { $items | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8; [Windows.Forms.MessageBox]::Show("Saved: $($dialog.FileName)") | Out-Null }
        catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Export failed') | Out-Null }
    }
    $dialog.Dispose()
}
$export.Add_Click({ Export-Rows $script:rows 'MailboxDelegateAccess' })
$exportIssues.Add_Click({ Export-Rows $script:issues 'ScanIssues' })
[void]$form.ShowDialog()
$form.Dispose()
