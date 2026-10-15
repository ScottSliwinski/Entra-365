#Requires -Version 5.1
<#
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Run on Windows: powershell.exe -STA -File .\Set-365CalendarDelegate-GUI.ps1
Requires Exchange Online administrative mailbox folder permissions.
Targets Exchange Online calendars, including hybrid synced users.
#>
$ErrorActionPreference='Stop'
if ([Environment]::OSVersion.Platform -ne 'Win32NT') { throw 'This GUI requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Launch with powershell.exe -STA -File .\Set-365CalendarDelegate-GUI.ps1' }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$script:connected=$false
$script:report=@()
$form=New-Object Windows.Forms.Form
$form.Text='Microsoft 365 Calendar Delegate'
$form.Size=New-Object Drawing.Size(850,780)
$form.MinimumSize=$form.Size
$form.StartPosition='CenterScreen'
$form.Font=New-Object Drawing.Font('Segoe UI',10)
function Label($text,$y) {
 $c=New-Object Windows.Forms.Label
 $c.Text=$text; $c.Location=New-Object Drawing.Point(25,$y)
 $labelWidth=760
 if ($y -in @(108,150,192)) { $labelWidth=235 }
 $c.Size=New-Object Drawing.Size($labelWidth,25)
 $form.Controls.Add($c)
}
function InputBox($y) {
 $c=New-Object Windows.Forms.TextBox
 $c.Location=New-Object Drawing.Point(270,$y); $c.Size=New-Object Drawing.Size(525,28)
 $form.Controls.Add($c); return $c
}
function Button($text,$x,$y,$w) {
 $c=New-Object Windows.Forms.Button
 $c.Text=$text; $c.Location=New-Object Drawing.Point($x,$y); $c.Size=New-Object Drawing.Size($w,36)
 $form.Controls.Add($c); return $c
}
function CheckBox($text,$y,$checked) {
 $c=New-Object Windows.Forms.CheckBox
 $c.Text=$text; $c.Location=New-Object Drawing.Point(25,$y); $c.Size=New-Object Drawing.Size(770,28); $c.Checked=$checked
 $form.Controls.Add($c); return $c
}
Label 'Sign in with an Exchange Online administrator account.' 20
$connect=Button 'Connect to Microsoft 365' 25 52 300
Label 'Calendar owner UPN:' 108
$owner=InputBox 105
Label 'Delegate UPN:' 150
$delegate=InputBox 147
Label 'Calendar permission:' 192
$role=New-Object Windows.Forms.ComboBox
$role.Location=New-Object Drawing.Point(270,189); $role.Size=New-Object Drawing.Size(525,30)
$role.DropDownStyle='DropDownList'
$role.Items.AddRange([object[]]@('Editor','Reviewer','AvailabilityOnly','LimitedDetails','Owner','PublishingEditor','PublishingAuthor','Author','NonEditingAuthor','Contributor','None'))
$role.SelectedIndex=0; $form.Controls.Add($role)
$meeting=CheckBox 'Receive meeting requests and responses as a delegate (requires Editor)' 232 $true
$private=CheckBox 'Allow delegate to view private calendar items' 264 $false
$notify=CheckBox 'Send a calendar sharing invitation to the delegate' 296 $false
Label 'Editor: create, edit and delete items. Reviewer: read items. Changes require confirmation.' 334
$check=Button 'Check current access' 25 370 230
$apply=Button 'Assign / Update access' 278 370 250
$export=Button 'Export displayed access' 550 370 245
$remove=Button 'Remove selected delegate' 25 412 280
$grid=New-Object Windows.Forms.DataGridView
$grid.Location=New-Object Drawing.Point(25,460); $grid.Size=New-Object Drawing.Size(770,145)
$grid.ReadOnly=$true; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false
$grid.SelectionMode='FullRowSelect'; $grid.MultiSelect=$false
$grid.RowHeadersVisible=$false; $grid.AutoSizeColumnsMode='Fill'
foreach ($name in @('User','UserPrincipalName','Resolution','AccessRights','SharingPermissionFlags')) {
 $column=New-Object Windows.Forms.DataGridViewTextBoxColumn
 $column.Name=$name; $column.HeaderText=$name; [void]$grid.Columns.Add($column)
}
$form.Controls.Add($grid)
$log=New-Object Windows.Forms.TextBox
$log.Location=New-Object Drawing.Point(25,620); $log.Size=New-Object Drawing.Size(770,95)
$log.Multiline=$true; $log.ReadOnly=$true; $log.ScrollBars='Vertical'; $form.Controls.Add($log)
function Log($message) { $log.AppendText(('{0:HH:mm:ss}  {1}' -f (Get-Date),$message)+[Environment]::NewLine); $log.Refresh() }
function Run([scriptblock]$action) {
 $form.UseWaitCursor=$true
 foreach ($b in @($connect,$check,$apply,$export,$remove)) { $b.Enabled=$false }
 try { & $action } catch {
  Log "ERROR: $($_.Exception.Message)"
  [void][Windows.Forms.MessageBox]::Show($form,$_.Exception.Message,'Operation failed','OK','Error')
 } finally {
  $form.UseWaitCursor=$false
  foreach ($b in @($connect,$check,$apply,$export,$remove)) { $b.Enabled=$true }
 }
}
function Resolve-User($upn) {
 $upn=$upn.Trim()
 if ($upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw 'Enter the full UPN, such as user@domain.com.' }
 $mailbox=Get-Mailbox -Identity $upn -ErrorAction Stop
 if ([string]$mailbox.UserPrincipalName -ine $upn) { throw "UPN resolved to $($mailbox.UserPrincipalName). Enter that actual UPN." }
 return $mailbox
}
function Get-Calendar {
 if (-not $script:connected) { throw 'Connect to Microsoft 365 first.' }
 $m=Resolve-User $owner.Text
 $folders=@(Get-MailboxFolderStatistics -Identity $m.UserPrincipalName -FolderScope Calendar -ErrorAction Stop | Where-Object { $_.FolderType -eq 'Calendar' })
 if ($folders.Count -ne 1) { throw 'Could not uniquely identify the default calendar.' }
 $path=([string]$folders[0].FolderPath).Replace('/','\')
 return [pscustomobject]@{Mailbox=$m; Identity=('{0}:{1}' -f $m.PrimarySmtpAddress,$path)}
}
function Show-Access($cal) {
 $entries=@(Get-MailboxFolderPermission -Identity $cal.Identity -ErrorAction Stop)
 $script:displayedCalendar=$cal
 $script:report=@($entries | ForEach-Object {
  $entry=$_; $upn=''; $resolution='Unresolved'
  if ([string]$entry.User -in @('Default','Anonymous')) { $resolution='Built-in' }
  else {
   try { $upn=Resolve-PermissionUPN $entry $cal; $resolution='Resolved' }
   catch { Log "UPN unavailable for $([string]$entry.User): $($_.Exception.Message)" }
  }
  [pscustomobject]@{CalendarOwner=[string]$cal.Mailbox.UserPrincipalName; CalendarFolder=$cal.Identity; User=[string]$entry.User; UserPrincipalName=$upn; Resolution=$resolution; AccessRights=($entry.AccessRights -join ', '); SharingPermissionFlags=($entry.SharingPermissionFlags -join ', ')}
 })
 $grid.Rows.Clear()
 for ($i=0; $i -lt $script:report.Count; $i++) {
  $r=$script:report[$i]
  $index=$grid.Rows.Add($r.User,$r.UserPrincipalName,$r.Resolution,$r.AccessRights,$r.SharingPermissionFlags)
  $grid.Rows[$index].Tag=$entries[$i]
 }
 $grid.ClearSelection()
 Log "Loaded permissions for $($cal.Identity)."
}
function Get-PermissionUserKey($entry) {
 # Preserve the recipient identity embedded in the permission entry.
 # A display name alone is not a unique Exchange recipient identifier.
 foreach ($obj in @($entry.User.ADRecipient,$entry.User)) {
  if (-not $obj) { continue }
  foreach ($name in @('PrimarySmtpAddress','UserPrincipalName','Guid','DistinguishedName','Sid')) {
   $property=$obj.PSObject.Properties[$name]
   if ($property -and $property.Value -and [string]$property.Value -ne '00000000-0000-0000-0000-000000000000') { return [string]$property.Value }
  }
 }
 $text=[string]$entry.User
 if ($text -match '^[^@\s]+@[^@\s]+$|^S-1-\d+-') { return $text }
 # This lookup must resolve uniquely; never choose the first name match.
 $recipient=Get-Recipient -Identity $text -ErrorAction Stop
 return [string]$recipient.PrimarySmtpAddress
}
function Resolve-PermissionUPN($entry,$cal) {
 try {
  $m=Get-Mailbox -Identity (Get-PermissionUserKey $entry) -ErrorAction Stop
  if ($m.UserPrincipalName) { return [string]$m.UserPrincipalName }
 } catch { }
 # When EXO returns only a name, test matching recipients against the
 # calendar ACL using their unique SMTP addresses. Never guess by name.
 $display=[string]$entry.User
 $escaped=$display.Replace("'","''")
 $candidates=@(Get-Recipient -Filter "(DisplayName -eq '$escaped') -or (Name -eq '$escaped') -or (Alias -eq '$escaped')" -ResultSize Unlimited -ErrorAction Stop)
 $matches=@(foreach ($candidate in $candidates) {
  try {
   $m=Get-Mailbox -Identity ([string]$candidate.PrimarySmtpAddress) -ErrorAction Stop
   $acl=Get-MailboxFolderPermission -Identity $cal.Identity -User ([string]$m.PrimarySmtpAddress) -ErrorAction Stop
   if ([string]$acl.User -eq $display -and $m.UserPrincipalName) { [string]$m.UserPrincipalName }
  } catch {
   $detail=@($_.FullyQualifiedErrorId,$_.Exception.Message,[string]$_.ErrorDetails) -join ' '
   if ($detail -notmatch 'UserNotFoundInPermissionEntry|There is no existing permission entry found for user') { throw }
  }
 })
 $matches=@($matches | Select-Object -Unique)
 if ($matches.Count -ne 1) { throw 'No unique mailbox UPN could be verified for this permission entry.' }
 return $matches[0]
}
$role.Add_SelectedIndexChanged({
 $meeting.Enabled=($role.SelectedItem -eq 'Editor')
 if (-not $meeting.Enabled) { $meeting.Checked=$false }
 $private.Enabled=$meeting.Checked
 if (-not $private.Enabled) { $private.Checked=$false }
})
$meeting.Add_CheckedChanged({ $private.Enabled=$meeting.Checked; if (-not $private.Enabled) { $private.Checked=$false } })
$connect.Add_Click({ Run {
 if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) { throw 'Install the module first: Install-Module ExchangeOnlineManagement -Scope CurrentUser' }
 Import-Module ExchangeOnlineManagement -ErrorAction Stop
 Log 'Connecting. Complete the Microsoft sign-in window.'
 Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
 $script:connected=$true
 Log 'Connected to Exchange Online.'
}})
$check.Add_Click({ Run { $cal=Get-Calendar; Show-Access $cal } })
$apply.Add_Click({ Run {
 $cal=Get-Calendar
 $d=Resolve-User $delegate.Text
 if ([string]$cal.Mailbox.ExchangeGuid -eq [string]$d.ExchangeGuid) { throw 'Calendar owner and delegate must be different users.' }
 $flags='None'
 if ($meeting.Checked) { $flags='Delegate'; if ($private.Checked) { $flags='Delegate,CanViewPrivateItems' } }
 $existing=$null
 try { $existing=Get-MailboxFolderPermission -Identity $cal.Identity -User $d.PrimarySmtpAddress -ErrorAction Stop }
 catch {
  # EXO REST sessions may return a generic error ID with the actual
  # missing-entry error only in ErrorDetails or Exception.Message.
  $errorText=@($_.FullyQualifiedErrorId,$_.Exception.Message,[string]$_.ErrorDetails) -join ' '
  if ($errorText -notmatch 'UserNotFoundInPermissionEntry|There is no existing permission entry found for user') { throw }
  $existing=$null
  Log "No existing calendar permission for $($d.UserPrincipalName); a new entry will be added after confirmation."
 }
 $verb='Add'; if ($existing) { $verb='Update' }
 $lines=@("$verb calendar access?",'',"Owner: $($cal.Mailbox.UserPrincipalName)","Delegate: $($d.UserPrincipalName)","Rights: $($role.SelectedItem)","Delegate flags: $flags","Send invitation: $($notify.Checked)")
 if ($existing) { $lines+=@('',"Previous rights: $($existing.AccessRights -join ', ')","Previous flags: $($existing.SharingPermissionFlags -join ', ')") }
 $answer=[Windows.Forms.MessageBox]::Show($form,($lines -join [Environment]::NewLine),'Confirm calendar access','YesNo','Question')
 if ($answer -ne 'Yes') { return }
 $p=@{Identity=$cal.Identity; User=[string]$d.PrimarySmtpAddress; AccessRights=[string]$role.SelectedItem; SharingPermissionFlags=$flags; SendNotificationToUser=[bool]$notify.Checked; ErrorAction='Stop'; Confirm=$false}
 if ($existing) { Set-MailboxFolderPermission @p } else { Add-MailboxFolderPermission @p }
 Log "Permission command succeeded for $($d.UserPrincipalName)."
 try {
  $v=Get-MailboxFolderPermission -Identity $cal.Identity -User $d.PrimarySmtpAddress -ErrorAction Stop
  Log "Verified rights: $($v.AccessRights -join ', '); flags: $($v.SharingPermissionFlags -join ', ')."
  Show-Access $cal
 } catch { throw "Permission command succeeded, but verification failed: $($_.Exception.Message). Check current access before retrying." }
}})
# A row keeps the original Exchange permission entry, so removal does not require
# the delegate to still have an active mailbox. Default/Anonymous are protected.
$grid.Add_CellClick({
 param($sender,$eventArgs)
 if ($eventArgs.RowIndex -lt 0) { return }
 $entry=$grid.Rows[$eventArgs.RowIndex].Tag
 if (-not $entry -or [string]$entry.User -in @('Default','Anonymous')) { return }
 # Keep a manually entered UPN if automatic name resolution is unavailable.
 try {
  $listedUPN=[string]$grid.Rows[$eventArgs.RowIndex].Cells['UserPrincipalName'].Value
  if (-not $listedUPN) { throw 'UPN was not resolved in the current access list.' }
  $resolved=Get-Mailbox -Identity $listedUPN -ErrorAction Stop
  $delegate.Text=[string]$resolved.UserPrincipalName
  $rights=@($entry.AccessRights | ForEach-Object { [string]$_ })
  if ($rights.Count -eq 1 -and $role.Items.Contains($rights[0])) { $role.SelectedItem=$rights[0] }
  else { Log 'This entry has custom rights. Choose the desired permission before updating.' }
  $meeting.Checked=($role.SelectedItem -eq 'Editor' -and ([string]$entry.SharingPermissionFlags -match 'Delegate'))
  $private.Checked=($meeting.Checked -and ([string]$entry.SharingPermissionFlags -match 'CanViewPrivateItems'))
 } catch {
  Log "Selected calendar entry: $([string]$entry.User). Enter or verify its UPN in Delegate UPN before removing or updating."
 }
})
$remove.Add_Click({ Run {
 if (-not $script:connected) { throw 'Connect to Microsoft 365 first.' }
 if ($grid.SelectedRows.Count -ne 1) { throw 'Select one delegate row in the current access list first.' }
 $selected=$grid.SelectedRows[0].Tag
 if (-not $selected) { throw 'Check current access and select a delegate.' }
 $userKey=[string]$selected.User
 if ($userKey -in @('Default','Anonymous')) { throw 'Default and Anonymous entries cannot be removed with this tool.' }
 $cal=Get-Calendar
 if ($cal.Identity -ine $script:displayedCalendar.Identity) { throw 'The calendar owner changed. Click Check current access and select a delegate again.' }
 # Use a unique recipient key for both lookup and removal.
 if ($delegate.Text.Trim()) {
  # Some REST responses contain only a display name. Permit an explicit
  # UPN only after confirming it maps to the sole selected-name ACL entry.
  $resolved=Resolve-User $delegate.Text
  $removeKey=[string]$resolved.PrimarySmtpAddress
  $sameName=@(Get-MailboxFolderPermission -Identity $cal.Identity -ErrorAction Stop | Where-Object { [string]$_.User -eq $userKey })
  $byUPN=Get-MailboxFolderPermission -Identity $cal.Identity -User $removeKey -ErrorAction Stop
  if ($sameName.Count -ne 1 -or [string]$byUPN.User -ne $userKey) { throw 'Entered UPN cannot be uniquely matched to the selected row. Removal stopped.' }
 } else {
  try { $removeKey=Get-PermissionUserKey $selected }
  catch { throw 'Enter the selected delegate full UPN in Delegate UPN, then click Remove selected delegate again.' }
 }
 $current=@(Get-MailboxFolderPermission -Identity $cal.Identity -User $removeKey -ErrorAction Stop)
 if ($current.Count -ne 1) { throw 'The selected entry changed or is no longer unique. Refresh current access before removing.' }
 $lines=@('Remove this delegate from this calendar?','',"Calendar owner: $($cal.Mailbox.UserPrincipalName)","Calendar: $($cal.Identity)","Selected delegate: $userKey","Unique identity: $removeKey","Current rights: $($current[0].AccessRights -join ', ')","Delegate flags: $($current[0].SharingPermissionFlags -join ', ')",'','This removes all permissions for this entry on this calendar.')
 if ([Windows.Forms.MessageBox]::Show($form,($lines -join [Environment]::NewLine),'Confirm delegate removal','YesNo','Warning') -ne 'Yes') { return }
 Remove-MailboxFolderPermission -Identity $cal.Identity -User $removeKey -Force -Confirm:$false -ErrorAction Stop
 Log "Removal command succeeded for $userKey."
 try {
  $remaining=$null
  try { $remaining=Get-MailboxFolderPermission -Identity $cal.Identity -User $removeKey -ErrorAction Stop }
  catch {
   $errorText=@($_.FullyQualifiedErrorId,$_.Exception.Message,[string]$_.ErrorDetails) -join ' '
   if ($errorText -notmatch 'UserNotFoundInPermissionEntry|There is no existing permission entry found for user') { throw }
  }
  if ($remaining) { throw 'The selected delegate permission is still present.' }
  Show-Access $cal
  Log "Verified removal of $userKey."
  $delegate.Clear()
 } catch { throw "Removal command succeeded, but verification failed: $($_.Exception.Message). Check current access before retrying." }
}})
$export.Add_Click({ Run {
 if ($script:report.Count -eq 0) { throw 'Check current access before exporting.' }
 $save=New-Object Windows.Forms.SaveFileDialog
 $save.Filter='CSV files (*.csv)|*.csv'; $save.FileName='CalendarPermissions.csv'
 try { if ($save.ShowDialog($form) -eq 'OK') { $script:report | Export-Csv -LiteralPath $save.FileName -NoTypeInformation -Encoding UTF8; Log "Exported $($save.FileName)." } }
 finally { $save.Dispose() }
}})
Log 'Connect and check access. Select a delegate row to update permissions or remove access.'
[void]$form.ShowDialog()
$form.Dispose()
