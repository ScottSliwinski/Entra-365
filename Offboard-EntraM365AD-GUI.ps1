#requires -Version 7.2
<#
WINDOWS POWERSHELL 7 GUI - run in a fresh process:
HYBRID additions: install Windows RSAT AD tools and the ActiveDirectory module.
Optional OneDrive: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
SPO is imported through Windows PowerShell compatibility. Requires Windows PowerShell 5.1,
SharePoint administrator rights, and a supplied SharePoint admin URL for your tenant.
AD uses the current Windows credentials and selected DC. Entra on-prem SID must match.
Select Hybrid actions to open an OU/group/action dialog after tenant verification.
All hybrid options start unchecked; check the steps needed. Passwords are cryptographically
random (16 characters), never exported, and must satisfy your domain policy.
The primary AD group is preserved. Add/remove applies to direct memberships only.
AD changes await your normal directory sync. No automatic sync is triggered.
The destination OU must stay in sync scope; moving out of scope can delete the cloud account.
Hidden GAL uses AD Exchange schema attributes unless IsExchangeCloudManaged is true.
OneDrive site admin access grants full control, preserves existing administrators, and
exports the existing URL; owner-changed or unprovisioned sites are reported as not found.
Test mode previews all new writes without performing them.

TEST MODE is checked by default. Select the cleanup actions to inspect, leave
TEST MODE checked, then run. Test mode does NOT revoke sessions, block sign-in,
delete methods/devices/events/rules/photos, or remove memberships. It only reads
service data, authenticates administrators, and writes local reports. Actions.csv
and Preview CSVs list proposed changes, skips and inspection failures. Read-only
permissions do not prove live write authorization; state may change between runs.
Uncheck TEST MODE for live execution; confirmations still apply. Required export
failure stops the run and produces an incomplete preview, rather than guessing.
 pwsh.exe -STA -File .\Offboard-EntraM365-GUI.ps1
Install once:
 Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
 Install-Module ExchangeOnlineManagement -Scope CurrentUser

Graph sign-in offers a GUI device-code dialog (recommended) or Microsoft browser
sign-in. The code is shown in the dialog, with Open sign-in and Copy code buttons;
it uses Microsoft's actual expiration time, not a fixed 120-second wait. Tokens
are held in memory, renewed for longer runs, and cleared at the end of the run.
No custom app registration is needed for the GUI device-code sign-in: it uses
Microsoft Graph Command Line Tools. Conditional Access can still block device
code flow; use the browser option in that case. This script targets public cloud.
Requires consent to the scopes below AND appropriate Entra/Exchange admin roles.
Privileged Authentication Administrator is needed for privileged targets; Exchange
recipient management is needed for rules/groups. Hidden membership requires its
read permission. Role-assignable groups require Privileged Role Administrator.
Calendar deletion requires actual access to the target mailbox: delegated Graph
Calendars.ReadWrite.Shared plus calendar sharing/delegation, or a certificate app
with admin-consented APPLICATION Calendars.ReadWrite (prefer Exchange App RBAC
scoping to the target). App fields are used ONLY for calendar work in a child
PowerShell process; directory operations always use the signed-in administrator.
No automatic mailbox permission grants. Missing calendar access is logged.

Reports are snapshots, not restorable backups. Licenses/app assignments/shared
mailbox permissions are EXPORTED, not explicitly removed. Group removal can
remove inherited licenses/app access. Group ownership and directory roles are
not removed. Synced/dynamic groups cannot be manually removed here.
Device deletion is Entra object deletion, NOT Intune retire/wipe or local data erase.
Authentication password methods cannot be deleted. Default MFA dependencies and
other API restrictions can prevent deletion; remaining methods are exported.
Calendar cleanup covers writable calendars listed by Graph, all historical/future
items and series masters. It does not purge deleted items or retention copies.
Rules means visible/hidden inbox rules (including Outlook rules blob on removal),
not tenant mail flow rules. Photo removal propagation can take time.
Session revocation is not instant and does not terminate every third-party app
session; block sign-in is optional and confirmed separately.
Shared mailbox scan is optional: a Yes/No prompt appears in both Test and Live mode.
Choosing No skips the scan and its two CSVs, records Skipped in Actions.csv,
and continues the other selected steps.
Mailbox report: direct FullAccess/SendAs/SendOnBehalf plus permissions assigned
to groups Graph reports transitively. Exchange-only nested distribution groups,
folder/calendar-only permissions, and application access aren't exhaustively
resolved; group trustees are exported separately for review. Errors are never
reported as proof of no access. Empty CSVs retain headers.
#>
[CmdletBinding()]
param([string]$CalendarUserId,[string]$CalendarTenantId,[string]$CalendarClientId,
      [string]$CalendarCertificateThumbprint,[string]$CalendarResultPath,[switch]$TestMode)
$ErrorActionPreference = 'Stop'
$script:TestRun=[bool]$TestMode
$script:ManualGraphToken=$null
# Device-code OAuth runs directly in a GUI dialog, so it never relies on
# Connect-MgGraph writing a prompt through a Windows Forms event handler.
# Tokens stay in memory and are never placed in reports, logs, or arguments.
function Set-GraphGuiToken($Token,$Scope) {
 if(-not $Token.access_token){throw 'Microsoft did not return an access token.'}
 $refresh=if($Token.refresh_token){$Token.refresh_token}elseif($script:ManualGraphToken){$script:ManualGraphToken.Refresh}else{$null}
 $script:ManualGraphToken=@{Refresh=$refresh;Scope=$Scope;Expires=[DateTime]::UtcNow.AddSeconds([int]$Token.expires_in)}
 $secure=ConvertTo-SecureString $Token.access_token -AsPlainText -Force
 Connect-MgGraph -AccessToken $secure -NoWelcome -ErrorAction Stop | Out-Null
 $secure=$null;$Token=$null
}
function Refresh-GraphGuiToken {
 if([DateTime]::UtcNow -lt $script:ManualGraphToken.Expires.AddSeconds(-120)){return}
 if(-not $script:ManualGraphToken.Refresh){throw 'Graph sign-in expired. Start another run and sign in again.'}
 try {
  $token=Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' -TimeoutSec 30 -Body @{
   client_id='14d82eec-204b-4c2f-b7e8-296a70dab67e';grant_type='refresh_token'
   refresh_token=$script:ManualGraphToken.Refresh;scope=$script:ManualGraphToken.Scope
  } -ErrorAction Stop
  Set-GraphGuiToken $token $script:ManualGraphToken.Scope
 }catch{throw 'Graph sign-in could not be renewed. Start another run and sign in again; no further actions were attempted.'}
}
function Connect-GraphGui([string[]]$Scopes) {
 $script:ManualGraphToken=$null
 $choice=[Windows.Forms.MessageBox]::Show($form,
  "Choose how to sign in to Microsoft Graph.`n`nYES: Show a device code directly in this GUI (recommended).`nNO: Use Microsoft's browser/interactive sign-in.`nCANCEL: Stop this run.`n`nSign in with an administrator, not the account being offboarded.",
  'Microsoft Graph sign-in','YesNoCancel','Information','Button1')
 if($choice -eq 'Cancel'){throw 'Sign-in canceled. No account changes made.'}
 if($choice -eq 'No'){
  Log 'Opening Microsoft interactive sign-in. Check for a browser or Microsoft sign-in window.'
  Connect-MgGraph -Scopes $Scopes -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
  return
 }
 $scopeText=(@($Scopes|ForEach-Object{'https://graph.microsoft.com/'+$_})+@('openid','profile','offline_access')) -join ' '
 $device=Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode' -TimeoutSec 30 -Body @{
  client_id='14d82eec-204b-4c2f-b7e8-296a70dab67e';scope=$scopeText
 } -ErrorAction Stop
 if(-not $device.user_code -or -not $device.device_code){throw 'Microsoft did not return a device sign-in code.'}
 $dialog=[Windows.Forms.Form]::new();$dialog.Text='Microsoft Graph - device sign-in';$dialog.Size=[Drawing.Size]::new(670,360);$dialog.StartPosition='CenterParent';$dialog.FormBorderStyle='FixedDialog';$dialog.MaximizeBox=$false;$dialog.MinimizeBox=$false
 $instructions=[Windows.Forms.Label]::new();$instructions.Text="1. Open https://microsoft.com/devicelogin`n2. Enter the code below and sign in with your administrator account.`n3. Leave this dialog open; it closes automatically after successful sign-in.";$instructions.Location=[Drawing.Point]::new(15,15);$instructions.Size=[Drawing.Size]::new(620,70);$dialog.Controls.Add($instructions)
 $code=[Windows.Forms.TextBox]::new();$code.Text=$device.user_code;$code.ReadOnly=$true;$code.Font=[Drawing.Font]::new('Consolas',22);$code.Location=[Drawing.Point]::new(15,95);$code.Size=[Drawing.Size]::new(620,50);$dialog.Controls.Add($code)
 $open=[Windows.Forms.Button]::new();$open.Text='Open Microsoft sign-in';$open.Location=[Drawing.Point]::new(15,160);$open.Size=[Drawing.Size]::new(210,32);$dialog.Controls.Add($open)
 $copy=[Windows.Forms.Button]::new();$copy.Text='Copy code';$copy.Location=[Drawing.Point]::new(240,160);$copy.Size=[Drawing.Size]::new(150,32);$dialog.Controls.Add($copy)
 $cancel=[Windows.Forms.Button]::new();$cancel.Text='Cancel';$cancel.Location=[Drawing.Point]::new(405,160);$cancel.Size=[Drawing.Size]::new(150,32);$dialog.Controls.Add($cancel)
 $status=[Windows.Forms.Label]::new();$status.Location=[Drawing.Point]::new(15,210);$status.Size=[Drawing.Size]::new(620,85);$dialog.Controls.Add($status)
 $open.Add_Click({Start-Process 'https://microsoft.com/devicelogin'})
 $copy.Add_Click({[Windows.Forms.Clipboard]::SetText($code.Text)}.GetNewClosure())
 $cancel.Add_Click({$dialog.Close()}.GetNewClosure())
 $deadline=[DateTime]::UtcNow.AddSeconds([int]$device.expires_in)
 $interval=[Math]::Max(5,[int]$device.interval);$next=[DateTime]::UtcNow.AddSeconds($interval);$token=$null
 # Parent is disabled to prevent another action while waiting; dialog is modeless
 # so this loop can pump GUI messages and poll only after the server interval.
 $form.Enabled=$false
 try{
  $dialog.Show($form)
  while(-not $dialog.IsDisposed -and $dialog.Visible -and [DateTime]::UtcNow -lt $deadline){
   $seconds=[Math]::Max(0,[int]($deadline-[DateTime]::UtcNow).TotalSeconds)
   $status.Text="Waiting for Microsoft sign-in... Code expires in $seconds seconds.`nNo account changes have been made. You can cancel this run."
   [Windows.Forms.Application]::DoEvents()
   if($dialog.IsDisposed -or -not $dialog.Visible){break}
   if([DateTime]::UtcNow -ge $next){
    try{
     $token=Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' -TimeoutSec 20 -Body @{
      client_id='14d82eec-204b-4c2f-b7e8-296a70dab67e';grant_type='urn:ietf:params:oauth:grant-type:device_code';device_code=$device.device_code
     } -ErrorAction Stop
     break
    }catch{
     $oauth=$null
     try{$oauth=$_.ErrorDetails.Message|ConvertFrom-Json -ErrorAction Stop}catch{}
     if($oauth.error -eq 'authorization_pending'){}
     elseif($oauth.error -eq 'slow_down'){$interval+=5}
     elseif($oauth.error -eq 'authorization_declined'){throw 'Microsoft sign-in was declined. No account changes made.'}
     elseif($oauth.error -eq 'expired_token'){throw 'The Microsoft device code expired. Run again to request a new code.'}
     elseif($oauth.error){throw "Microsoft device sign-in failed: $($oauth.error). Review Microsoft sign-in policy or try the browser option."}
     else{throw 'Unable to contact Microsoft sign-in. Check network/proxy access and retry.'}
    }
    $next=[DateTime]::UtcNow.AddSeconds($interval)
   }
   Start-Sleep -Milliseconds 100
  }
  if(-not $token){if([DateTime]::UtcNow -ge $deadline){throw 'Microsoft device-code sign-in expired. Run again for a new code.'}else{throw 'Sign-in canceled. No account changes made.'}}
  # Respect a close/cancel that occurred while the last network call was pending.
  [Windows.Forms.Application]::DoEvents()
  if($dialog.IsDisposed -or -not $dialog.Visible){throw 'Sign-in canceled. No account changes made.'}
  Set-GraphGuiToken $token $scopeText
  Log 'Microsoft Graph sign-in completed.'
 }finally{
  $form.Enabled=$true
  if(-not $dialog.IsDisposed){$dialog.Close();$dialog.Dispose()}
  $token=$null;$device=$null
 }
}
function G($Path,$Method='GET',$Body=$null) {
 if($script:ManualGraphToken){Refresh-GraphGuiToken}
 if($script:TestRun -and $Method -notin @('GET','HEAD')){throw "Test mode blocked Graph mutation: $Method $Path"}
 $uri = if($Path -like 'https://*'){$Path}else{'https://graph.microsoft.com/v1.0/'+$Path}
 $args=@{Uri=$uri;Method=$Method;OutputType='Hashtable';ErrorAction='Stop';Headers=@{ConsistencyLevel='eventual'}}
 if($null -ne $Body){$args.Body=$Body|ConvertTo-Json -Depth 15;$args.ContentType='application/json'}
 Invoke-MgGraphRequest @args
}
function All($Path) {
 $items=[Collections.Generic.List[object]]::new()
 do {$r=G $Path;foreach($v in $r.value){$items.Add($v)};$Path=$r.'@odata.nextLink'}while($Path)
 $items.ToArray()
}
function E($Value){[uri]::EscapeDataString([string]$Value)}
function CalendarClean($uid) {
 $result=[Collections.Generic.List[object]]::new()
 $calendars=@(All "users/$uid/calendars?`$select=id,name,canEdit,owner")
 foreach($c in $calendars){
  if(-not $c.canEdit){$result.Add([pscustomobject]@{Calendar=$c.name;Item='';Status='Skipped';Detail='Calendar is not writable'});continue}
  try {
   # Snapshot complete pagination before deleting, to avoid shifted pages.
   $events=@(All ("users/$uid/calendars/$(E $c.id)/events?`$select=id,subject,type&`$top=100"))
   foreach($ev in $events){try{G "users/$uid/calendars/$(E $c.id)/events/$(E $ev.id)" DELETE|Out-Null
    $result.Add([pscustomobject]@{Calendar=$c.name;Item=$ev.id;Status='Deleted';Detail=$ev.subject})
   }catch{$result.Add([pscustomobject]@{Calendar=$c.name;Item=$ev.id;Status='Failed';Detail=$_.Exception.Message})}}
   $left=@(All "users/$uid/calendars/$(E $c.id)/events?`$select=id&`$top=100")
   $result.Add([pscustomobject]@{Calendar=$c.name;Item='';Status='Verified';Detail="$($left.Count) items remain"})
  }catch{$result.Add([pscustomobject]@{Calendar=$c.name;Item='';Status='Failed';Detail=$_.Exception.Message})}
 }
 $result.ToArray()
}
function CalendarPreview($uid) {
 $result=[Collections.Generic.List[object]]::new()
 try {
  foreach($c in @(All "users/$uid/calendars?`$select=id,name,canEdit,owner")){
   try {
    $events=@(All "users/$uid/calendars/$(E $c.id)/events?`$select=id,subject,type,isOrganizer,start,end&`$top=100")
    foreach($ev in $events){$result.Add([pscustomobject]@{Calendar=$c.name;Item=$ev.id;Status=if($c.canEdit){'Would delete'}else{'Would skip'};Detail="$($ev.subject) | Type=$($ev.type) | Organizer=$($ev.isOrganizer) | Start=$($ev.start.dateTime) | End=$($ev.end.dateTime)"})}
    if(-not $events.Count){$result.Add([pscustomobject]@{Calendar=$c.name;Item='';Status='No items';Detail='Calendar is empty'})}
   }catch{$result.Add([pscustomobject]@{Calendar=$c.name;Item='';Status='Cannot inspect';Detail=$_.Exception.Message})}
  }
 }catch{$result.Add([pscustomobject]@{Calendar='';Item='';Status='Cannot inspect';Detail=$_.Exception.Message})}
 $result.ToArray()
}
if($CalendarUserId){
 try{Import-Module Microsoft.Graph.Authentication
 Connect-MgGraph -TenantId $CalendarTenantId -ClientId $CalendarClientId -CertificateThumbprint $CalendarCertificateThumbprint -ContextScope Process -NoWelcome
 $rows=if($TestMode){@(CalendarPreview $CalendarUserId)}else{@(CalendarClean $CalendarUserId)}
 if($rows.Count){$rows|Export-Csv $CalendarResultPath -NoTypeInformation -Encoding utf8}else{'"Calendar","Item","Status","Detail"'|Set-Content $CalendarResultPath}
 Disconnect-MgGraph|Out-Null;exit 0
 }catch{$_|Out-String|Set-Content ($CalendarResultPath+'.error.txt');exit 1}
}
if(-not $IsWindows){throw 'This Windows Forms GUI requires Windows and PowerShell 7.'}
if([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA'){throw 'Start with pwsh.exe -STA -File .\Offboard-EntraM365-GUI.ps1'}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$script:RunDir=$null;$script:Audit=[Collections.Generic.List[object]]::new()
$form=[Windows.Forms.Form]::new();$form.Text='Hybrid AD / Entra / Microsoft 365 Offboarding';$form.Size=[Drawing.Size]::new(900,830);$form.StartPosition='CenterScreen'
function Label($text,$y){$l=[Windows.Forms.Label]::new();$l.Text=$text;$l.Location=[Drawing.Point]::new(15,$y);$l.Size=[Drawing.Size]::new(845,22);$form.Controls.Add($l)}
function Box($text,$y){$b=[Windows.Forms.TextBox]::new();$b.Text=$text;$b.Location=[Drawing.Point]::new(15,$y);$b.Size=[Drawing.Size]::new(845,25);$form.Controls.Add($b);return $b}
Label 'User UPN to offboard:' 15;$upnBox=Box '' 40
Label 'Report folder (a timestamped folder with the UPN will be created):' 73;$folderBox=Box ([Environment]::GetFolderPath('MyDocuments')+'\OffboardingReports') 98
Label 'Optional calendar-only certificate app: Tenant ID | Client ID | Certificate thumbprint (leave blank for delegated access)' 131
$appBox=Box '' 156
Label 'Select cleanup steps. Each selected step still asks for confirmation. Exports run before cleanup.' 191
$options=[ordered]@{}
$names=@('Block sign-in (recommended)','Clear authentication methods','Delete Entra device objects (confirm each)','Clear all calendar entries (meeting cancellations may be sent)','Remove all inbox rules (including hidden rules)','Remove account photo','Remove cloud group memberships (confirm each)')
$y=219
foreach($name in $names){$cb=[Windows.Forms.CheckBox]::new();$cb.Text=$name;$cb.Location=[Drawing.Point]::new(15,$y);$cb.Size=[Drawing.Size]::new(840,23);$form.Controls.Add($cb);$options[$name]=$cb;$y+=25}
$testBox=[Windows.Forms.CheckBox]::new();$testBox.Text='TEST MODE - preview and export only (no account changes)';$testBox.Checked=$true;$testBox.Location=[Drawing.Point]::new(15,402);$testBox.Size=[Drawing.Size]::new(840,25);$form.Controls.Add($testBox)
$hybridBox=[Windows.Forms.CheckBox]::new();$hybridBox.Text='Configure on-prem AD actions and group changes (opens a separate dialog)';$hybridBox.SetBounds(15,432,840,24);$form.Controls.Add($hybridBox)
$oneDriveBox=[Windows.Forms.CheckBox]::new();$oneDriveBox.Text='Export OneDrive URL / optionally add a OneDrive site administrator';$oneDriveBox.SetBounds(15,460,840,24);$form.Controls.Add($oneDriveBox)
$forwardBox=[Windows.Forms.CheckBox]::new();$forwardBox.Text='Ask about mail forwarding (one internal UPN)';$forwardBox.SetBounds(15,488,840,24);$form.Controls.Add($forwardBox)
$mailAccessBox=[Windows.Forms.CheckBox]::new();$mailAccessBox.Text='Ask about granting mailbox Full Access (multiple UPNs)';$mailAccessBox.SetBounds(15,516,840,24);$form.Controls.Add($mailAccessBox)
$run=[Windows.Forms.Button]::new();$run.Text='Verify user and run';$run.Location=[Drawing.Point]::new(15,550);$run.Size=[Drawing.Size]::new(190,32);$form.Controls.Add($run)
$log=[Windows.Forms.TextBox]::new();$log.Multiline=$true;$log.ReadOnly=$true;$log.ScrollBars='Vertical';$log.Location=[Drawing.Point]::new(15,595);$log.Size=[Drawing.Size]::new(845,160);$form.Controls.Add($log)
function Log($Message){$line="$(Get-Date -Format s) $Message";$log.AppendText($line+[Environment]::NewLine);if($script:RunDir){Add-Content (Join-Path $script:RunDir 'Run.log') $line};$log.Refresh()}
function Confirm($Message){[Windows.Forms.MessageBox]::Show($form,$Message,'Confirm offboarding action','YesNo','Warning','Button2') -eq 'Yes'}
function Audit($Action,$Target,$Status,$Detail=''){
 $row=[pscustomobject]@{Time=(Get-Date -Format o);UPN=$script:UPN;Action=$Action;Target=$Target;Status=$Status;Detail=$Detail}
 $script:Audit.Add($row);$row|Export-Csv (Join-Path $script:RunDir ($script:Safe+'_Actions.csv')) -Append -NoTypeInformation -Encoding utf8
 Log "$Action [$Status] $Target $Detail"
}
function Csv($Name,$Rows,$Columns){$p=Join-Path $script:RunDir ($script:Safe+'_'+$Name+'.csv');$r=@($Rows);if($r.Count){$r|Select-Object $Columns|Export-Csv $p -NoTypeInformation -Encoding utf8}else{('"'+($Columns -join '","')+'"')|Set-Content $p -Encoding utf8};Log "Exported $Name ($($r.Count) rows)"}
function Step($Name,$Work){try{& $Work;Audit $Name $script:UPN 'Completed'}catch{Audit $Name $script:UPN 'Failed' $_.Exception.Message}}
# Hybrid additions: all writes go through HybridAction and the existing Test guard.
function HybridAction($Name,$Target,$Detail,[scriptblock]$Work) {
 if($script:TestRun){Audit $Name $Target 'Would perform' $Detail;return}
 if(Confirm "$Name`nTarget: $Target`n$Detail"){
  try{& $Work;Audit $Name $Target 'Completed' $Detail}catch{Audit $Name $Target 'Failed' $_.Exception.Message}
 }else{Audit $Name $Target 'Skipped' 'Administrator declined'}
}
function InputText($Title,$Default='') {
 $d=[Windows.Forms.Form]::new();$d.Text=$Title;$d.Size=[Drawing.Size]::new(650,170);$d.StartPosition='CenterParent'
 $t=[Windows.Forms.TextBox]::new();$t.Text=$Default;$t.SetBounds(15,15,600,25);$d.Controls.Add($t)
 $ok=[Windows.Forms.Button]::new();$ok.Text='OK';$ok.SetBounds(420,55,90,30);$ok.DialogResult='OK';$d.Controls.Add($ok)
 $no=[Windows.Forms.Button]::new();$no.Text='Cancel';$no.SetBounds(520,55,90,30);$no.DialogResult='Cancel';$d.Controls.Add($no);$d.AcceptButton=$ok;$d.CancelButton=$no
 try{if($d.ShowDialog($form) -eq 'OK'){return $t.Text.Trim()}return $null}finally{$d.Dispose()}
}
function New-OffboardPassword {
 # Cryptographic random selection; guarantee each character class and shuffle.
 $sets=@('ABCDEFGHJKLMNPQRSTUVWXYZ','abcdefghijkmnopqrstuvwxyz','23456789','!@#$%&*+-=?')
 $chars=[Collections.Generic.List[char]]::new()
 foreach($set in $sets){$chars.Add($set[[Security.Cryptography.RandomNumberGenerator]::GetInt32($set.Length)])}
 $all=$sets -join '';while($chars.Count -lt 16){$chars.Add($all[[Security.Cryptography.RandomNumberGenerator]::GetInt32($all.Length)])}
 for($i=$chars.Count-1;$i -gt 0;$i--){$j=[Security.Cryptography.RandomNumberGenerator]::GetInt32($i+1);$tmp=$chars[$i];$chars[$i]=$chars[$j];$chars[$j]=$tmp}
 ConvertTo-SecureString (-join $chars) -AsPlainText -Force
}
function HybridPrepare($User) {
 $server=InputText 'Domain controller / AD DNS domain (required)'
 if(-not $server){throw 'Hybrid setup canceled: no domain controller selected.'}
 Import-Module ActiveDirectory -ErrorAction Stop
 $adargs=@{Server=$server;ErrorAction='Stop'}
 if(-not $User.onPremisesSyncEnabled -or -not $User.onPremisesSecurityIdentifier){throw 'Hybrid AD actions require a synced account with a Graph on-premises SID. No UPN-only match is used.'}
 $ad=Get-ADUser -Identity $User.onPremisesSecurityIdentifier -Properties @('Description','Manager','MemberOf','PrimaryGroupID','mailNickname') @adargs
 if([string]$ad.SID -ne [string]$User.onPremisesSecurityIdentifier){throw 'AD SID does not match the Entra account.'}
 if(-not(Confirm "Verify hybrid identity:`nEntra: $script:UPN`nAD: $($ad.DistinguishedName)`nSID: $($ad.SID)`nDC: $server`nContinue configuring the AD actions?")){throw 'Hybrid identity review canceled.'}
 $d=[Windows.Forms.Form]::new();$d.Text="Hybrid actions: $script:UPN";$d.Size=[Drawing.Size]::new(850,720);$d.StartPosition='CenterParent'
 $checks=@{};$y=10
 foreach($n in @('Reset AD password (16 random characters)','Disable AD account','Move to selected OU','Set description','Clear manager','Hide mailbox from address lists')){
  $c=[Windows.Forms.CheckBox]::new();$c.Text=$n;$c.SetBounds(15,$y,780,24);$d.Controls.Add($c);$checks[$n]=$c;$y+=26
 }
 $desc=[Windows.Forms.TextBox]::new();$desc.Text="Disabled: Prep'd for Decom";$desc.SetBounds(15,175,790,25);$d.Controls.Add($desc)
 $ou=[Windows.Forms.ComboBox]::new();$ou.DropDownStyle='DropDownList';$ou.SetBounds(15,210,790,25);$d.Controls.Add($ou)
 foreach($o in @(Get-ADOrganizationalUnit -Filter * @adargs | Sort-Object DistinguishedName)){[void]$ou.Items.Add([string]$o.DistinguishedName)}
 $label=[Windows.Forms.Label]::new();$label.Text='Direct AD groups: checked = keep; unchecked = propose removal. Primary group is preserved.';$label.SetBounds(15,250,800,25);$d.Controls.Add($label)
 $list=[Windows.Forms.CheckedListBox]::new();$list.CheckOnClick=$true;$list.SetBounds(15,280,790,265);$d.Controls.Add($list)
 $existing=@($ad.MemberOf|ForEach-Object {Get-ADGroup -Identity $_ @adargs})
 $map=[Collections.Generic.List[object]]::new()
 foreach($g in $existing){$map.Add($g);[void]$list.Items.Add("$($g.Name) | $($g.DistinguishedName)",$true)}
 $add=[Windows.Forms.Button]::new();$add.Text='Add group (name or DN)';$add.SetBounds(15,555,230,30);$d.Controls.Add($add)
 $add.Add_Click({
  try{$q=InputText 'AD group name, GUID or distinguished name';if(-not $q){return};$g=Get-ADGroup -Identity $q @adargs
   if($g.DistinguishedName -notin @($map|ForEach-Object DistinguishedName)){$map.Add($g);[void]$list.Items.Add("$($g.Name) | $($g.DistinguishedName)",$true)}
  }catch{[Windows.Forms.MessageBox]::Show($d,$_.Exception.Message,'Group lookup failed')|Out-Null}
 }.GetNewClosure())
 $removeAll=[Windows.Forms.Button]::new();$removeAll.Text='Remove all direct groups';$removeAll.SetBounds(260,555,240,30);$d.Controls.Add($removeAll)
 $removeAll.Add_Click({for($j=0;$j -lt $list.Items.Count;$j++){$list.SetItemChecked($j,$false)}}.GetNewClosure())
 $keepAll=[Windows.Forms.Button]::new();$keepAll.Text='Keep all groups';$keepAll.SetBounds(515,555,180,30);$d.Controls.Add($keepAll)
 $keepAll.Add_Click({for($j=0;$j -lt $list.Items.Count;$j++){$list.SetItemChecked($j,$true)}}.GetNewClosure())
 $ok=[Windows.Forms.Button]::new();$ok.Text='Save action selections';$ok.SetBounds(560,605,245,32);$ok.DialogResult='OK';$d.Controls.Add($ok)
 $cancel=[Windows.Forms.Button]::new();$cancel.Text='Cancel';$cancel.SetBounds(440,605,100,32);$cancel.DialogResult='Cancel';$d.Controls.Add($cancel)
 try{
  if($d.ShowDialog($form) -ne 'OK'){throw 'Hybrid action configuration canceled.'}
  $selected=@($checks.Keys|Where-Object {$checks[$_].Checked})
  if('Move to selected OU' -in $selected -and -not $ou.SelectedItem){throw 'Select a destination OU for the move.'}
  $changes=for($i=0;$i -lt $map.Count;$i++){
   $g=$map[$i];$was=$g.DistinguishedName -in $ad.MemberOf;$want=$list.GetItemChecked($i)
   if($was -and -not $want){[pscustomobject]@{Group=$g;Action='Remove'}}
   elseif(-not $was -and $want){[pscustomobject]@{Group=$g;Action='Add'}}
  }
  $primarySid=([string]$ad.SID -replace '-\d+$','')+'-'+$ad.PrimaryGroupID
  $primary=Get-ADGroup -Identity $primarySid @adargs
  Csv 'ADGroupsBefore' (@($existing|ForEach-Object {[pscustomobject]@{Name=$_.Name;DN=$_.DistinguishedName;Primary=$false}})+@([pscustomobject]@{Name=$primary.Name;DN=$primary.DistinguishedName;Primary=$true})) @('Name','DN','Primary')
  Csv 'ADAccountBefore' @([pscustomobject]@{UPN=$script:UPN;ADUPN=$ad.UserPrincipalName;SID=[string]$ad.SID;DN=$ad.DistinguishedName;Enabled=$ad.Enabled;Description=$ad.Description;Manager=$ad.Manager;DC=$server}) @('UPN','ADUPN','SID','DN','Enabled','Description','Manager','DC')
  Csv 'ADGroupChanges' @($changes|ForEach-Object {[pscustomobject]@{Group=$_.Group.Name;DN=$_.Group.DistinguishedName;Action=$_.Action;Mode=if($script:TestRun){'Test'}else{'Pending confirmation'}}}) @('Group','DN','Action','Mode')
  return @{AD=$ad;Args=$adargs;Selected=$selected;OU=[string]$ou.SelectedItem;Description=$desc.Text;Changes=@($changes)}
 }finally{$d.Dispose()}
}
function HybridExecute($Plan,$User) {
 $ad=$Plan.AD;$a=$Plan.Args;$id=$ad.ObjectGUID
 foreach($n in $Plan.Selected){
  switch($n){
   'Reset AD password (16 random characters)' {HybridAction $n $ad.SamAccountName 'Generate a new 16-character password; password is not displayed or exported. Domain password policy still applies.' { $pw=New-OffboardPassword;try{Set-ADAccountPassword -Identity $id -Reset -NewPassword $pw @a}finally{$pw.Dispose()} }}
   'Disable AD account' {HybridAction $n $ad.SamAccountName 'Disable the on-premises account.' {Disable-ADAccount -Identity $id @a}}
   'Move to selected OU' {HybridAction $n $ad.SamAccountName ("Destination: $($Plan.OU). Confirm this OU remains in Entra sync scope; an out-of-scope move can delete the cloud account.") {Get-ADOrganizationalUnit -Identity $Plan.OU @a|Out-Null;Move-ADObject -Identity $id -TargetPath $Plan.OU @a}}
   'Set description' {HybridAction $n $ad.SamAccountName $Plan.Description {Set-ADUser -Identity $id -Description $Plan.Description @a}}
   'Clear manager' {HybridAction $n $ad.SamAccountName "Current manager: $($ad.Manager)" {Set-ADUser -Identity $id -Clear manager @a}}
   'Hide mailbox from address lists' {HybridAction $n $script:UPN 'Use Exchange Online when Exchange attributes are cloud managed; otherwise update the AD Exchange attribute and await normal synchronization.' {
    $mb=Get-Mailbox -Identity $script:UPN -ErrorAction Stop
    if($mb.IsExchangeCloudManaged -eq $true){Set-Mailbox -Identity $script:UPN -HiddenFromAddressListsEnabled $true -ErrorAction Stop}
    else{if(-not $ad.mailNickname){throw 'AD mailNickname is empty. Configure the Exchange alias with your supported Exchange management tools, then retry.'};Set-ADUser -Identity $id -Replace @{msExchHideFromAddressLists=$true} @a}
   }}
  }
 }
 if(-not @($Plan.Changes).Count){Audit 'AD group changes' $ad.SamAccountName 'No changes selected' 'Existing groups were kept. In the Hybrid dialog, uncheck groups or click Remove all direct groups.'}
 foreach($c in $Plan.Changes){$g=$c.Group
  HybridAction ("AD group "+$c.Action) $g.DistinguishedName "User: $($ad.SamAccountName). Group licensing and application access may change after synchronization." {
   $currentAD=Get-ADUser -Identity $id -Properties MemberOf @a
   if($c.Action -eq 'Add'){Add-ADGroupMember -Identity $g.ObjectGUID -Members $currentAD.DistinguishedName @a}else{Remove-ADGroupMember -Identity $g.ObjectGUID -Members $currentAD.DistinguishedName -Confirm:$false @a}
   $verifiedAD=Get-ADUser -Identity $id -Properties MemberOf @a
   $stillMember=$g.DistinguishedName -in @($verifiedAD.MemberOf)
   if(($c.Action -eq 'Remove' -and $stillMember) -or ($c.Action -eq 'Add' -and -not $stillMember)){throw 'AD group change could not be verified on the selected DC.'}
  }
 }
 Csv 'ADGroupResults' @($script:Audit|Where-Object {$_.Action -like 'AD group *'}) @('Time','UPN','Action','Target','Status','Detail')
 $after=Get-ADUser -Identity $id -Properties @('Description','Manager','MemberOf') @a
 Csv 'ADAccountAfter' @([pscustomobject]@{UPN=$script:UPN;DN=$after.DistinguishedName;Enabled=$after.Enabled;Description=$after.Description;Manager=$after.Manager;Mode=if($script:TestRun){'Unchanged test snapshot'}else{'Live'}}) @('UPN','DN','Enabled','Description','Manager','Mode')
 Csv 'ADGroupsAfter' @($after.MemberOf|ForEach-Object {[pscustomobject]@{DN=$_}}) @('DN')
}
function OneDriveAccess {
 $adminUrl=InputText 'SharePoint admin URL (https://TENANT-admin.sharepoint.com)'
 if($adminUrl -notmatch '^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$'){throw 'Enter a valid public-cloud SharePoint admin URL.'}
 $delegate=InputText 'UPN to add as OneDrive site administrator (blank = export URL only)'
 if($delegate){$du=G "users/$(E $delegate)?`$select=id,userPrincipalName,userType";if($du.userType -ne 'Member' -or $du.userPrincipalName -ine $delegate -or $du.id -eq $uid){throw 'Select the exact UPN of another internal member account for OneDrive delegation.'};$delegate=$du.userPrincipalName}
 Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -ErrorAction Stop
 Connect-SPOService -Url $adminUrl -ErrorAction Stop
 Log 'Finding existing OneDrive sites by exact owner UPN; no OneDrive is provisioned.'
 $sites=@(Get-SPOSite -IncludePersonalSite $true -Limit All -ErrorAction Stop | Where-Object {$_.Template -like 'SPSPERS*' -and (([string]$_.Owner -replace '^i:0#\.f\|membership\|','') -ieq $script:UPN)})
 if(-not $sites.Count){
  $siteInput=InputText 'OneDrive not found by owner. Paste the existing personal SITE URL (blank = skip)'
  if($siteInput){
   $siteUri=[uri]$siteInput;$expectedHost=([uri]$adminUrl).Host -replace '-admin\.','-my.'
   if($siteUri.Scheme -ne 'https' -or $siteUri.Host -ine $expectedHost -or $siteUri.AbsolutePath -notmatch '^/personal/[^/]+/?$' -or $siteUri.Query -or $siteUri.Fragment){throw 'Use the personal site root URL in the selected tenant, not a document or sharing link.'}
   $candidate=Get-SPOSite -Identity $siteUri.AbsoluteUri -ErrorAction Stop
   if($candidate.Template -notlike 'SPSPERS*'){throw 'Selected URL is not a OneDrive personal site.'}
   if(Confirm "Verify this existing OneDrive belongs to the offboarded account before continuing:`nAccount: $script:UPN`nURL: $($candidate.Url)`nCurrent owner: $($candidate.Owner)"){$sites=@($candidate)}
  }
 }
 $rows=[Collections.Generic.List[object]]::new()
 if(-not $sites.Count){$rows.Add([pscustomobject]@{UPN=$script:UPN;URL='';DelegateUPN=$delegate;Status='Not found';Detail='No existing personal site with matching owner. Renamed/owner-changed sites need manual review.'});Audit 'OneDrive' $script:UPN 'Not found' 'No matching personal site; no permission change.'}
 foreach($site in $sites){
  $status='URL exported';$detail=''
  if($delegate){
   if($script:TestRun){$status='Would add site administrator';Audit 'OneDrive delegation' $site.Url $status $delegate}
   elseif(Confirm "Grant OneDrive site administrator access?`nOwner: $script:UPN`nSite: $($site.Url)`nDelegate: $delegate`nThis grants full control of this OneDrive site."){
    try{Set-SPOUser -Site $site.Url -LoginName $delegate -IsSiteCollectionAdmin $true -ErrorAction Stop|Out-Null
     $status='Granted; verification pending'
     try{$siteUser=Get-SPOUser -Site $site.Url -LoginName $delegate -ErrorAction Stop;if(-not $siteUser.IsSiteAdmin){throw 'User is not a site collection administrator after the write.'};$status='Granted and verified'}catch{$detail='Set-SPOUser completed; verification could not confirm access: '+$_.Exception.Message}
     Audit 'OneDrive delegation' $site.Url $status ($delegate+' '+$detail)}catch{$status='Failed';$detail=$_.Exception.Message+' Verify SharePoint admin permissions and that the delegate is resolvable on this personal site.';Audit 'OneDrive delegation' $site.Url $status $detail}
   }else{$status='Declined';Audit 'OneDrive delegation' $site.Url $status $delegate}
  }
  $rows.Add([pscustomobject]@{UPN=$script:UPN;URL=$site.Url;DelegateUPN=$delegate;Status=$status;Detail=$detail})
 }
 Csv 'OneDriveAccess' $rows.ToArray() @('UPN','URL','DelegateUPN','Status','Detail')
}


# Mailbox handover: internal users only, exact Graph UPN/object ID matching.
function Resolve-HandoverUser([string]$RecipientInputUPN) {
 if($RecipientInputUPN -notmatch '^[^\s@,;]+@[^\s@,;]+$'){throw 'Enter one complete UPN per recipient; forwarding accepts only one UPN.'}
 $u=G "users/$(E $RecipientInputUPN)?`$select=id,userPrincipalName,userType"
 if($u.userPrincipalName -ine $RecipientInputUPN -or $u.userType -ne 'Member'){throw "Use the exact UPN of an internal member account: $RecipientInputUPN"}
 if($u.userPrincipalName -ieq $script:UPN){throw 'The recipient cannot be the account being offboarded.'}
 $r=Get-Recipient -Identity $u.userPrincipalName -ErrorAction Stop
 if([string]$r.ExternalDirectoryObjectId -ne [string]$u.id){throw "Exchange recipient does not match Entra user: $RecipientInputUPN"}
 if([string]$r.RecipientTypeDetails -notin @('UserMailbox','SharedMailbox')){throw "Recipient must have an Exchange Online user or shared mailbox: $RecipientInputUPN"}
 [pscustomobject]@{UPN=$u.userPrincipalName;Id=$u.id;Address=[string]$r.PrimarySmtpAddress}
}
function MailboxHandover($Forward,$Access,[string]$SourceObjectId) {
 $mb=Get-Mailbox -Identity $script:UPN -ErrorAction Stop
 if([string]$mb.ExternalDirectoryObjectId -ne $SourceObjectId){throw 'Mailbox object ID does not match the verified Entra account.'}
 $mailboxId=[string]$mb.ExchangeGuid
 Csv 'MailboxForwardingBefore' @([pscustomobject]@{UPN=$script:UPN;ForwardingAddress=[string]$mb.ForwardingAddress;ForwardingSmtpAddress=[string]$mb.ForwardingSmtpAddress;DeliverToMailboxAndForward=$mb.DeliverToMailboxAndForward}) @('UPN','ForwardingAddress','ForwardingSmtpAddress','DeliverToMailboxAndForward')
 $permissions=@(Get-MailboxPermission -Identity $mailboxId -ErrorAction Stop)
 Csv 'MailboxPermissionsBefore' ($permissions|ForEach-Object{[pscustomobject]@{UPN=$script:UPN;Trustee=[string]$_.User;AccessRights=($_.AccessRights -join ';');Deny=$_.Deny;IsInherited=$_.IsInherited}}) @('UPN','Trustee','AccessRights','Deny','IsInherited')
 $rows=[Collections.Generic.List[object]]::new()
 if($Forward){
  $status='Skipped';$detail='Administrator declined';$target='';$keep=$null
  try{
   if(Confirm "Does mail forwarding need to be enabled for $script:UPN?"){
    $entered=InputText 'Forward mail to ONE internal UPN'
    if($entered){
     $forwardRecipient=Resolve-HandoverUser $entered;$target=$forwardRecipient.UPN
     $forwardSmtp=[string]$forwardRecipient.Address
     if(-not $forwardSmtp -or $forwardRecipient.Id -eq $SourceObjectId -or $forwardSmtp -ieq [string]$mb.PrimarySmtpAddress){throw 'Self-forwarding or empty destination refused.'}
     $choice=[Windows.Forms.MessageBox]::Show($form,'Keep a copy in the original mailbox as well? Yes = keep and forward; No = forward only; Cancel = skip forwarding.','Forwarding delivery','YesNoCancel','Question','Button1')
     if($choice -ne 'Cancel'){
      $keep=($choice -eq 'Yes');$detail="Source mailbox: $script:UPN. Destination: $target ($forwardSmtp); keep original copy=$keep. Replaces existing mailbox forwarding: $($mb.ForwardingAddress) / $($mb.ForwardingSmtpAddress)."
      if($script:TestRun){$status='Would enable'}
      elseif(Confirm "Enable forwarding for $script:UPN?`n$detail"){
       Set-Mailbox -Identity $mailboxId -ForwardingAddress $null -ForwardingSmtpAddress $forwardSmtp -DeliverToMailboxAndForward $keep -ErrorAction Stop
       $after=Get-Mailbox -Identity $mailboxId -ErrorAction Stop
       $savedSmtp=([string]$after.ForwardingSmtpAddress -replace '^smtp:','')
       if($after.ForwardingAddress -or $savedSmtp -ine $forwardSmtp -or $after.DeliverToMailboxAndForward -ne $keep){throw 'Forwarding write returned but verification failed; inspect mailbox settings.'}
       $status='Enabled and verified'
      }else{$detail='Administrator declined final confirmation'}
     }else{$detail='Delivery choice cancelled'}
    }else{$detail='Recipient entry cancelled or empty'}
   }
  }catch{$status='Failed';$detail=$_.Exception.Message}
  Audit 'Mailbox forwarding' $script:UPN $status $detail
  $rows.Add([pscustomobject]@{UPN=$script:UPN;Action='Forwarding';RecipientUPN=$target;Permission='';KeepCopy=$keep;Status=$status;Detail=$detail})
 }
 if($Access){
  if(Confirm "Grant Full Access to the mailbox of $script:UPN?`nYou can enter multiple internal UPNs. Existing permissions are preserved."){
   $entered=InputText 'Mailbox Full Access UPNs (separate with commas or semicolons)'
   if($entered){
    $targets=@($entered -split '[,;\s]+'|Where-Object {$_}|Sort-Object -Unique)
    # Validate the complete list before granting any access.
    $validated=[Collections.Generic.List[object]]::new()
    try{foreach($target in $targets){$validated.Add((Resolve-HandoverUser $target))}}
    catch{Audit 'Mailbox access' $script:UPN 'Failed' $_.Exception.Message;$rows.Add([pscustomobject]@{UPN=$script:UPN;Action='Full Access';RecipientUPN=$entered;Permission='FullAccess';KeepCopy='';Status='Failed validation';Detail=$_.Exception.Message});$validated.Clear()}
    foreach($recipient in $validated){
     $status='Skipped';$detail='Administrator declined';$target=$recipient.UPN
     try{
      $existing=@(Get-MailboxPermission -Identity $mailboxId -User $recipient.Address -ErrorAction Stop)
      if(@($existing|Where-Object {$_.Deny -and 'FullAccess' -in $_.AccessRights}).Count){throw 'An explicit or inherited Full Access deny exists; review it before granting access.'}
      if(@($existing|Where-Object {-not $_.Deny -and 'FullAccess' -in $_.AccessRights}).Count){$status='Already granted';$detail='Existing allow permission preserved'}
      elseif($script:TestRun){$status='Would grant';$detail='FullAccess; AutoMapping=True'}
      elseif(Confirm "Grant Full Access to $target?`nMailbox being offboarded: $script:UPN`nAllows reading and changing mailbox contents; enables Outlook automapping."){
       Add-MailboxPermission -Identity $mailboxId -User $recipient.Address -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false -ErrorAction Stop|Out-Null
       $verify=@(Get-MailboxPermission -Identity $mailboxId -User $recipient.Address -ErrorAction Stop|Where-Object {-not $_.Deny -and 'FullAccess' -in $_.AccessRights})
       if(-not $verify.Count){throw 'Access write returned but verification failed; inspect permissions.'}
       $status='Granted and verified';$detail='FullAccess; AutoMapping=True'
      }
     }catch{$status='Failed';$detail=$_.Exception.Message}
     Audit 'Mailbox Full Access' $target $status $detail
     $rows.Add([pscustomobject]@{UPN=$script:UPN;Action='Full Access';RecipientUPN=$target;Permission='FullAccess';KeepCopy='';Status=$status;Detail=$detail})
    }
   }else{Audit 'Mailbox access' $script:UPN 'Skipped' 'Recipient entry cancelled or empty'}
  }else{Audit 'Mailbox access' $script:UPN 'Skipped' 'Administrator declined'}
 }
 Csv 'MailboxHandover' $rows.ToArray() @('UPN','Action','RecipientUPN','Permission','KeepCopy','Status','Detail')
}

$run.Add_Click({
 $script:RunDir=$null
 $forwardSelected=$forwardBox.Checked;$mailAccessSelected=$mailAccessBox.Checked;$forwardBox.Enabled=$false;$mailAccessBox.Enabled=$false
 $hybridSelected=$hybridBox.Checked;$oneDriveSelected=$oneDriveBox.Checked;$hybridBox.Enabled=$false;$oneDriveBox.Enabled=$false
 $script:TestRun=$testBox.Checked;$testBox.Enabled=$false
 $run.Enabled=$false;$upnBox.Enabled=$false;$folderBox.Enabled=$false;$appBox.Enabled=$false
 foreach($cb in $options.Values){$cb.Enabled=$false}
 try{
  foreach($m in @('Microsoft.Graph.Authentication','ExchangeOnlineManagement')){if(-not(Get-Module -ListAvailable $m)){throw "Install-Module $m -Scope CurrentUser in a separate PowerShell window first."}}
  Import-Module Microsoft.Graph.Authentication
  $scopes=@('User.ReadWrite.All','User.RevokeSessions.All','Directory.Read.All','Application.Read.All','GroupMember.ReadWrite.All','Member.Read.Hidden','UserAuthenticationMethod.ReadWrite.All','Device.ReadWrite.All','ProfilePhoto.ReadWrite.All','Calendars.ReadWrite.Shared','RoleManagement.ReadWrite.Directory')
  if($script:TestRun){$scopes=@('User.Read.All','Directory.Read.All','Application.Read.All','Member.Read.Hidden','UserAuthenticationMethod.Read.All','Device.Read.All','ProfilePhoto.Read.All','Calendars.Read.Shared')}
  $scopes=@($scopes)+@('User.Read')
  Log 'Choose Microsoft Graph sign-in in the dialog.'
  Connect-GraphGui -Scopes $scopes
  $inputUPN=$upnBox.Text.Trim();if($inputUPN -notmatch '^[^\s@]+@[^\s@]+$'){throw 'Enter a valid user UPN.'}
  $user=G "users/$(E $inputUPN)?`$select=id,displayName,userPrincipalName,accountEnabled,onPremisesSyncEnabled,onPremisesSecurityIdentifier,licenseAssignmentStates"
  $script:UPN=$user.userPrincipalName;$uid=$user.id
  # User-supplied access-token contexts don't populate SDK Account/TenantId.
  # Resolve both through authenticated Graph instead of assuming cached context.
  $me=G 'me?$select=id,userPrincipalName'
  $tenantInfo=@(All 'organization?$select=id')
  if($tenantInfo.Count -ne 1){throw 'Unable to verify the signed-in tenant.'}
  $context=@{Account=$me.userPrincipalName;TenantId=$tenantInfo[0].id}
  if($me.id -eq $uid -or $context.Account -eq $script:UPN){throw 'Offboarding your signed-in administrator is blocked.'}
  if(-not(Confirm "Tenant: $($context.TenantId)`nUser: $($user.displayName)`nUPN: $script:UPN`nObject ID: $uid`nEnabled: $($user.accountEnabled)`nSynced: $($user.onPremisesSyncEnabled)`n`nMode: $(if($script:TestRun){'TEST - no account changes'}else{'LIVE - confirmed changes'})`n$(if($script:TestRun){'Inspect the user and export an action preview?'}else{'Verify both connections, revoke sessions, then export reports?'})")){return}
  $script:Safe=$script:UPN -replace '[<>:"/\\|?*\x00-\x1f]','_'
  $script:RunDir=Join-Path $folderBox.Text ($script:Safe+$(if($script:TestRun){'_TEST_'}else{'_LIVE_'})+(Get-Date -Format 'yyyyMMdd_HHmmss_fff'));New-Item $script:RunDir -ItemType Directory|Out-Null
  $script:Audit.Clear();$user|ConvertTo-Json -Depth 20|Set-Content (Join-Path $script:RunDir ($script:Safe+'_User.json'))
  Log 'Connect to Exchange Online with an administrator from the same tenant.'
  Import-Module ExchangeOnlineManagement
  $exargs=@{ShowBanner=$false;ErrorAction='Stop';UserPrincipalName=$context.Account}
  if((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')){$exargs.DisableWAM=$true}
  Connect-ExchangeOnline @exargs
  # Verify the connection actually used by the unprefixed Exchange cmdlets.
  $connections=@(Get-ConnectionInformation -ErrorAction Stop | Where-Object {
   $_.State -eq 'Connected' -and -not $_.IsEopSession -and -not $_.ModulePrefix
  })
  $active=@($connections | Where-Object {$_.ConnectionUsedForInbuiltCmdlets -eq $true})
  if($active.Count -eq 0){$active=$connections}
  $tenantIds=@($active | ForEach-Object {[string]$_.TenantID} | Where-Object {$_} | Sort-Object -Unique)
  if($tenantIds.Count -gt 1){throw "Multiple Exchange tenants are connected. Close this window and launch the script in a fresh PowerShell process. Graph tenant: $($context.TenantId); Exchange tenants: $($tenantIds -join ', '). No cleanup started."}
  $exchangeTenant=if($tenantIds.Count -eq 1){$tenantIds[0]}else{''}
  if(-not $exchangeTenant){
   $org=Get-OrganizationConfig -ErrorAction Stop
   $exchangeTenant=[string]$org.ExternalDirectoryOrganizationId
  }
  $graphGuid=[guid]::Empty;$exchangeGuid=[guid]::Empty
  if(-not [guid]::TryParse([string]$context.TenantId,[ref]$graphGuid) -or
     -not [guid]::TryParse($exchangeTenant,[ref]$exchangeGuid) -or
     $graphGuid -eq [guid]::Empty -or $exchangeGuid -eq [guid]::Empty){
   throw "Unable to verify tenant IDs. Graph: '$($context.TenantId)'; Exchange: '$exchangeTenant'. Close this window, update ExchangeOnlineManagement, and retry in a fresh process. No cleanup started."
  }
  Log "Tenant verification: Graph=$graphGuid; Exchange=$exchangeGuid"
  if($exchangeGuid -ne $graphGuid){throw "Exchange and Graph tenants do not match. Graph: $graphGuid; Exchange: $exchangeGuid. Sign in to both with an administrator in the intended tenant. No cleanup started."}
  Audit 'Tenant verification' $script:UPN 'Verified' "Graph=$graphGuid; Exchange=$exchangeGuid"
  $hybridPlan=$null
  if($user.onPremisesSyncEnabled -and -not $hybridSelected){$hybridSelected=Confirm 'This account is synced from Windows AD. Configure on-prem AD actions and direct group removals now? Cloud group cleanup cannot remove synced memberships.'}
  if($hybridSelected){$hybridPlan=HybridPrepare $user}
  if(-not $script:TestRun){Step 'Revoke sessions' {$r=G "users/$uid/revokeSignInSessions" POST @{};if(-not $r.value){throw 'Graph did not return successful revocation.'}}}
  if(-not $script:TestRun -and $options[$names[0]].Checked -and (Confirm "Block sign-in for $script:UPN? For synced accounts, also disable the source AD account.")){Step 'Block sign-in' {G "users/$uid" PATCH @{accountEnabled=$false}|Out-Null}}
  # Required exports are a gate: failure aborts all subsequent cleanup.
  $direct=@(All "users/$uid/memberOf/microsoft.graph.group?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,onPremisesSyncEnabled,isAssignableToRole,mail")
  $transitive=@(All "users/$uid/transitiveMemberOf/microsoft.graph.group?`$count=true&`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,onPremisesSyncEnabled,isAssignableToRole,mail")
  $groups=foreach($g in $transitive){[pscustomobject]@{UPN=$script:UPN;Id=$g.id;Name=$g.displayName;Mail=$g.mail;Direct=($g.id -in $direct.id);GroupTypes=($g.groupTypes -join ';');MailEnabled=$g.mailEnabled;SecurityEnabled=$g.securityEnabled;Synced=$g.onPremisesSyncEnabled;RoleAssignable=$g.isAssignableToRole}}
  Csv 'Groups' $groups @('UPN','Id','Name','Mail','Direct','GroupTypes','MailEnabled','SecurityEnabled','Synced','RoleAssignable')
  # Include user app assignments and all transitive group app assignments.
  $apps=[Collections.Generic.List[object]]::new()
  $principals=@(@{id=$uid;name=$script:UPN;kind='users'})+@($transitive|ForEach-Object{@{id=$_.id;name=$_.displayName;kind='groups'}})
  foreach($p in $principals){foreach($a in @(All "$($p.kind)/$($p.id)/appRoleAssignments?`$count=true")){
   if($a.principalId -ne $p.id){continue}
   $sp=G "servicePrincipals/$($a.resourceId)?`$select=id,displayName,appId,appRoles"
   $role=@($sp.appRoles|Where-Object id -eq $a.appRoleId)
   $apps.Add([pscustomobject]@{UPN=$script:UPN;Source=$p.name;SourceType=$p.kind;EnterpriseApplication=$sp.displayName;AppId=$sp.appId;ServicePrincipalId=$sp.id;AppRoleId=$a.appRoleId;Role=($role.displayName -join ';')})
  }}
  Csv 'EnterpriseApplications' $apps.ToArray() @('UPN','Source','SourceType','EnterpriseApplication','AppId','ServicePrincipalId','AppRoleId','Role')
  $lic=@(All "users/$uid/licenseDetails")
  $licRows=foreach($l in $lic){$states=@($user.licenseAssignmentStates|Where-Object skuId -eq $l.skuId);[pscustomobject]@{UPN=$script:UPN;SkuId=$l.skuId;SkuPartNumber=$l.skuPartNumber;AssignmentSources=(($states|ForEach-Object{if($_.assignedByGroup){$_.assignedByGroup}else{'Direct'}})-join ';');AssignmentState=($states.state -join ';');ServicePlans=(($l.servicePlans|ForEach-Object{"$($_.servicePlanName):$($_.provisioningStatus)"})-join ';')}}
  Csv 'Licenses' $licRows @('UPN','SkuId','SkuPartNumber','AssignmentSources','AssignmentState','ServicePlans')
  if(Confirm "Run the Shared Mailbox Access scan for $script:UPN?`nThis scans all shared mailboxes for Full Access, Send As and Send on Behalf permissions and exports the results.`nChoose No to skip this scan and continue the other selected steps."){
  $trustees=@($uid)+@($transitive.id);$access=[Collections.Generic.List[object]]::new();$groupReview=[Collections.Generic.List[object]]::new()
  $recipient=Get-Recipient -Identity $script:UPN
  $cache=@{}
  function MatchTrustee($value){
   $key=[string]$value;if(-not $cache.ContainsKey($key)){try{$r=Get-Recipient -Identity $key -ErrorAction Stop;$cache[$key]=$r}catch{$cache[$key]=$null}}
   $r=$cache[$key];if($r){return ([string]$r.ExternalDirectoryObjectId -in $trustees -or [string]$r.PrimarySmtpAddress -eq [string]$recipient.PrimarySmtpAddress)}
   return ($key -in @($script:UPN,[string]$recipient.Sid,[string]$recipient.DistinguishedName))
  }
  $mailboxes=@(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited)
  $i=0
  foreach($mb in $mailboxes){$i++;Log "Shared mailbox scan $i / $($mailboxes.Count): $($mb.PrimarySmtpAddress)"
   $perms=@(Get-MailboxPermission -Identity $mb.Identity |Where-Object{ -not $_.IsInherited -and $_.User -ne 'NT AUTHORITY\SELF'})
   $sas=@(Get-RecipientPermission -Identity $mb.Identity|Where-Object{ $_.Trustee -ne 'NT AUTHORITY\SELF'})
   $entries=@($perms|ForEach-Object{[pscustomobject]@{Trustee=[string]$_.User;Permission=($_.AccessRights -join ';');Deny=$_.Deny}})+@($sas|ForEach-Object{[pscustomobject]@{Trustee=[string]$_.Trustee;Permission=($_.AccessRights -join ';');Deny=($_.AccessControlType -eq 'Deny')}})+@($mb.GrantSendOnBehalfTo|ForEach-Object{[pscustomobject]@{Trustee=[string]$_;Permission='SendOnBehalf';Deny=$false}})
   foreach($p in $entries){$matched=MatchTrustee $p.Trustee;$resolved=$cache[$p.Trustee]
    $row=[pscustomobject]@{UPN=$script:UPN;SharedMailbox=[string]$mb.PrimarySmtpAddress;Trustee=$p.Trustee;Permission=$p.Permission;Deny=$p.Deny;Match=if($matched){'Direct or Graph-transitive group'}else{'Requires review'};TrusteeType=if($resolved){[string]$resolved.RecipientTypeDetails}else{'Unresolved'}}
    if($matched){$access.Add($row)}elseif(-not $resolved -or [string]$resolved.RecipientTypeDetails -match 'Group'){$groupReview.Add($row)}
   }
  }
  Csv 'SharedMailboxAccess' $access.ToArray() @('UPN','SharedMailbox','Trustee','Permission','Deny','Match','TrusteeType')
  Csv 'SharedMailboxGroupTrusteesReview' $groupReview.ToArray() @('UPN','SharedMailbox','Trustee','Permission','Deny','Match','TrusteeType')
   Audit 'Shared mailbox access scan' $script:UPN 'Completed' 'Shared mailbox access and group trustee review CSVs exported.'
  }else{
   Audit 'Shared mailbox access scan' $script:UPN 'Skipped' 'Administrator declined. Shared mailbox permissions were not inspected; access CSVs were not created.'
  }
  if($forwardSelected -or $mailAccessSelected){Step 'Mailbox handover' {MailboxHandover $forwardSelected $mailAccessSelected $uid}}
  if($hybridPlan){Step 'Hybrid AD actions' {HybridExecute $hybridPlan $user}}
  if($oneDriveSelected){Step 'OneDrive access' {OneDriveAccess}}
  if($script:TestRun){
   Audit 'Revoke sessions' $script:UPN 'Would perform' 'Live mode revokes sessions after user verification; not performed in test mode.'
   for($n=0;$n -lt $names.Count;$n++){Audit 'Selected cleanup' $names[$n] $(if($options[$names[$n]].Checked){'Selected'}else{'Not selected'})}
   if($options[$names[0]].Checked){Audit 'Block sign-in' $script:UPN 'Would perform' "Current accountEnabled=$($user.accountEnabled). Synced source AD must also be managed separately."}
   if($options[$names[1]].Checked){
    try{
     $methods=@(All "users/$uid/authentication/methods")
     Csv 'AuthenticationMethodsPreview' ($methods|ForEach-Object{[pscustomobject]@{Id=$_.id;Type=$_.'@odata.type';Action=if($_.'@odata.type' -eq '#microsoft.graph.passwordAuthenticationMethod'){'Cannot delete password'}else{'Would attempt deletion; API restrictions may apply'}}}) @('Id','Type','Action')
     foreach($m in $methods){Audit 'Authentication cleanup' $m.id $(if($m.'@odata.type' -eq '#microsoft.graph.passwordAuthenticationMethod'){'Would skip'}else{'Would attempt'}) $m.'@odata.type'}
    }catch{Audit 'Authentication preview' $script:UPN 'Cannot inspect' $_.Exception.Message}
   }
   if($options[$names[2]].Checked){
    try{
     $devices=@(All "users/$uid/registeredDevices/microsoft.graph.device?`$select=id,displayName,deviceId,operatingSystem,onPremisesSyncEnabled")+@(All "users/$uid/ownedDevices/microsoft.graph.device?`$select=id,displayName,deviceId,operatingSystem,onPremisesSyncEnabled")
     $deviceRows=foreach($d in @($devices|Sort-Object id -Unique)){
      $detail='';$action='Would ask to delete Entra object'
      try{
       $owners=@(All "devices/$($d.id)/registeredOwners");$users=@(All "devices/$($d.id)/registeredUsers")
       $others=@(($owners+$users)|Where-Object id -ne $uid|Sort-Object id -Unique)
       if($d.onPremisesSyncEnabled -or $others.Count){$action='Would skip';$detail='Synced or associated with other users'}
      }catch{$action='Cannot inspect associations';$detail=$_.Exception.Message}
      Audit 'Device preview' $d.id $action "$($d.displayName) $detail"
      [pscustomobject]@{Id=$d.id;Name=$d.displayName;DeviceId=$d.deviceId;OS=$d.operatingSystem;Synced=$d.onPremisesSyncEnabled;Action=$action;Detail=$detail}
     }
     Csv 'DevicesPreview' $deviceRows @('Id','Name','DeviceId','OS','Synced','Action','Detail')
    }catch{Audit 'Device preview' $script:UPN 'Cannot inspect' $_.Exception.Message}
   }
   if($options[$names[3]].Checked){
    try{
     $app=$appBox.Text.Trim();$path=Join-Path $script:RunDir ($script:Safe+'_CalendarPreview.csv')
     if($app){
      $parts=@($app -split '\|'|ForEach-Object{$_.Trim()})
      if($parts.Count -ne 3 -or @($parts|Where-Object{ -not $_}).Count){throw 'Enter Tenant ID | Client ID | Certificate thumbprint.'}
      if($parts[0] -ne $context.TenantId){throw 'Calendar app tenant must match Graph tenant.'}
      & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $PSCommandPath -CalendarUserId $uid -CalendarTenantId $parts[0] -CalendarClientId $parts[1] -CalendarCertificateThumbprint $parts[2] -CalendarResultPath $path -TestMode
      if($LASTEXITCODE -ne 0){throw "Calendar preview child failed. See $path.error.txt"}
     }else{Csv 'CalendarPreview' @(CalendarPreview $uid) @('Calendar','Item','Status','Detail')}
     $calrows=@(Import-Csv $path)
     foreach($r in $calrows){Audit 'Calendar preview' ($r.Calendar+':'+$r.Item) $r.Status $r.Detail}
     Audit 'Calendar preview summary' $script:UPN 'Preview only' "$(@($calrows|Where-Object Status -eq 'Would delete').Count) items proposed. A read-only preview does not prove deletion permissions. Organizer meetings would send cancellations."
    }catch{Audit 'Calendar preview' $script:UPN 'Cannot inspect' $_.Exception.Message}
   }
   if($options[$names[4]].Checked){
    try{
     $rules=@(Get-InboxRule -Mailbox $script:UPN -IncludeHidden)
     Csv 'InboxRulesPreview' ($rules|Select-Object -Property @('Identity','Name','Enabled','Description')) @('Identity','Name','Enabled','Description')
     foreach($r in $rules){Audit 'Inbox rule preview' ([string]$r.Identity) 'Would delete' $r.Name}
     Audit 'Inbox rule preview summary' $script:UPN 'Preview only' "$($rules.Count) rules; Outlook client-side rules blob would also be removed by rule changes."
    }catch{Audit 'Inbox rule preview' $script:UPN 'Cannot inspect' $_.Exception.Message}
   }
   if($options[$names[5]].Checked){
    try{$photo=G "users/$uid/photo";Audit 'Photo preview' $script:UPN 'Would delete' "Photo exists: $($photo.width)x$($photo.height)"}
    catch{Audit 'Photo preview' $script:UPN 'Unavailable or cannot inspect' $_.Exception.Message}
   }
   if($options[$names[6]].Checked){
    $groupRows=foreach($g in $direct){
     $action=if($g.onPremisesSyncEnabled -or 'DynamicMembership' -in $g.groupTypes){'Would skip synced/dynamic group'}else{'Would ask to remove membership'}
     $route=if($g.mailEnabled -and 'Unified' -notin $g.groupTypes){'Exchange'}else{'Graph'}
     Audit 'Group membership preview' $g.displayName $action "Route=$route; RoleAssignable=$($g.isAssignableToRole). May remove inherited licenses and application access."
     [pscustomobject]@{Id=$g.id;Name=$g.displayName;Mail=$g.mail;Action=$action;Route=$route;RoleAssignable=$g.isAssignableToRole}
    }
    Csv 'GroupRemovalPreview' $groupRows @('Id','Name','Mail','Action','Route','RoleAssignable')
   }
   Audit 'TEST completed' $script:UPN 'No account changes' 'Only reads, sign-ins and local report writes occurred. Live permissions were not validated.'
   [Windows.Forms.MessageBox]::Show($form,"TEST completed - no account changes were made.`nReview Actions.csv and the Preview CSVs.`nUnchecked actions were not selected for cleanup.`n$script:RunDir",'Test results')|Out-Null
   return
  }
  Audit 'Required exports' $script:RunDir 'Completed' 'Proceeding to confirmed cleanup.'
  if($options[$names[1]].Checked -and (Confirm "Delete all removable authentication methods for $script:UPN? Password cannot be deleted. API restrictions will be logged.")){
   Step 'Authentication cleanup' {
    $map=@{microsoftAuthenticatorAuthenticationMethod='microsoftAuthenticatorMethods';phoneAuthenticationMethod='phoneMethods';emailAuthenticationMethod='emailMethods';fido2AuthenticationMethod='fido2Methods';softwareOathAuthenticationMethod='softwareOathMethods';temporaryAccessPassAuthenticationMethod='temporaryAccessPassMethods';windowsHelloForBusinessAuthenticationMethod='windowsHelloForBusinessMethods';platformCredentialAuthenticationMethod='platformCredentialMethods'}
    # Repeat for dependencies, such as alternate phone before primary phone.
    for($pass=0;$pass -lt 3;$pass++){
     $methods=@(All "users/$uid/authentication/methods");$progress=0
     foreach($m in $methods){$type=$m.'@odata.type' -replace '^#microsoft.graph.','';if($map.ContainsKey($type)){
      try{G "users/$uid/authentication/$($map[$type])/$(E $m.id)" DELETE|Out-Null;Audit 'Delete authentication method' $m.id 'Deleted' $type;$progress++}
      catch{if($pass -eq 2){Audit 'Delete authentication method' $m.id 'Failed' $_.Exception.Message}}
     }}
     if($progress -eq 0){break}
    }
    $remaining=@(All "users/$uid/authentication/methods")
    Csv 'AuthenticationMethodsRemaining' ($remaining|ForEach-Object{[pscustomobject]@{Id=$_.id;Type=$_.'@odata.type'}}) @('Id','Type')
    foreach($m in $remaining){Audit 'Authentication method remaining' $m.id 'Remaining' $m.'@odata.type'}
   }
  }
  if($options[$names[2]].Checked){Step 'Device cleanup' {
   $devices=@(All "users/$uid/registeredDevices/microsoft.graph.device?`$select=id,displayName,deviceId,operatingSystem,onPremisesSyncEnabled")+@(All "users/$uid/ownedDevices/microsoft.graph.device?`$select=id,displayName,deviceId,operatingSystem,onPremisesSyncEnabled")
   $devices=@($devices|Sort-Object id -Unique)
   Csv 'Devices' ($devices|ForEach-Object{[pscustomobject]@{Id=$_.id;Name=$_.displayName;DeviceId=$_.deviceId;OS=$_.operatingSystem;Synced=$_.onPremisesSyncEnabled}}) @('Id','Name','DeviceId','OS','Synced')
   foreach($d in $devices){try{
    $owners=@(All "devices/$($d.id)/registeredOwners");$users=@(All "devices/$($d.id)/registeredUsers")
    $others=@(($owners+$users)|Where-Object id -ne $uid|Sort-Object id -Unique)
    if($d.onPremisesSyncEnabled -or $others.Count){Audit 'Delete device' $d.id 'Skipped' 'Synced or associated with other users; manual review required.';continue}
    if(Confirm "Delete Entra device object?`nUser: $script:UPN`nDevice: $($d.displayName)`nObject ID: $($d.id)`nDevice ID: $($d.deviceId)`nThis affects the device tenant-wide, may break access, and does not wipe/retire Intune or erase local data."){G "devices/$($d.id)" DELETE|Out-Null;Audit 'Delete device' $d.id 'Deleted'}else{Audit 'Delete device' $d.id 'Skipped' 'Administrator declined'}
   }catch{Audit 'Delete device' $d.id 'Failed' $_.Exception.Message}}
  }}
  if($options[$names[3]].Checked -and (Confirm "Delete ALL historical and future entries and recurring series in writable calendars for $script:UPN?`nOrganizer meetings send cancellations to attendees. This is not just cancellation of future meetings. Reports are NOT calendar backups.")){
   Step 'Calendar cleanup' {
    $app=$appBox.Text.Trim();$path=Join-Path $script:RunDir ($script:Safe+'_CalendarResults.csv')
    if($app){$parts=@($app -split '\|'|ForEach-Object{$_.Trim()});if($parts.Count -ne 3 -or @($parts|Where-Object{ -not $_}).Count){throw 'Enter Tenant ID | Client ID | Certificate thumbprint.'}
     if($parts[0] -ne $context.TenantId){throw 'Calendar app tenant must match Graph tenant.'}
     # Child isolates certificate Graph context from interactive directory operations.
     & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -File $PSCommandPath -CalendarUserId $uid -CalendarTenantId $parts[0] -CalendarClientId $parts[1] -CalendarCertificateThumbprint $parts[2] -CalendarResultPath $path
     if($LASTEXITCODE -ne 0){throw "Calendar child failed. See $path.error.txt"}
    }else{Csv 'CalendarResults' @(CalendarClean $uid) @('Calendar','Item','Status','Detail')}
    foreach($r in @(Import-Csv $path)){Audit 'Calendar item' ($r.Calendar+':'+$r.Item) $r.Status $r.Detail}
   }
  }
  if($options[$names[4]].Checked -and (Confirm "Remove all visible and hidden inbox rules for $script:UPN? Exchange rule changes also remove Outlook client-side rules.")){
   Step 'Inbox rule cleanup' {
    $rules=@(Get-InboxRule -Mailbox $script:UPN -IncludeHidden)
    Csv 'InboxRules' ($rules|Select-Object -Property @('Identity','Name','Enabled','Description')) @('Identity','Name','Enabled','Description')
    foreach($rule in $rules){try{Remove-InboxRule -Mailbox $script:UPN -Identity $rule.Identity -AlwaysDeleteOutlookRulesBlob -Confirm:$false;Audit 'Remove rule' ([string]$rule.Identity) 'Deleted'}catch{Audit 'Remove rule' ([string]$rule.Identity) 'Failed' $_.Exception.Message}}
    $left=@(Get-InboxRule -Mailbox $script:UPN -IncludeHidden);Audit 'Verify rules' $script:UPN 'Verified' "$($left.Count) rules remain"
   }
  }
  if($options[$names[5]].Checked -and (Confirm "Remove the account photo for $script:UPN?")){Step 'Remove photo' {G "users/$uid/photo/`$value" DELETE|Out-Null}}
  if($options[$names[6]].Checked -and (Confirm "Review and remove direct cloud group memberships for $script:UPN?`nThis may remove group-based licenses, Teams access and enterprise application access. Synced/dynamic memberships are skipped. Each eligible group asks separately.")){
   Step 'Cloud group cleanup' {
    $current=@(All "users/$uid/memberOf/microsoft.graph.group?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,onPremisesSyncEnabled,isAssignableToRole,mail")
    foreach($g in $current){
     if($g.onPremisesSyncEnabled -or 'DynamicMembership' -in $g.groupTypes){Audit 'Remove group membership' $g.displayName 'Skipped' 'Synced or dynamic group';continue}
     if(-not(Confirm "Remove $script:UPN from $($g.displayName)?`nGroup ID: $($g.id)`nRole-assignable: $($g.isAssignableToRole)`nGroup-based licensing and application access can be lost.")){Audit 'Remove group membership' $g.displayName 'Skipped' 'Administrator declined';continue}
     try{
      if($g.mailEnabled -and 'Unified' -notin $g.groupTypes){
       $dg=Get-DistributionGroup -Identity $g.id
       if($dg.IsDirSynced){throw 'Exchange says this group is synced; removal refused.'}
       Remove-DistributionGroupMember -Identity $dg.Identity -Member $script:UPN -BypassSecurityGroupManagerCheck -Confirm:$false
      }else{G ("groups/$($g.id)/members/$uid/"+'$ref') DELETE|Out-Null}
      Audit 'Remove group membership' $g.displayName 'Removed'
     }catch{Audit 'Remove group membership' $g.displayName 'Failed' $_.Exception.Message}
    }
    $left=@(All "users/$uid/memberOf/microsoft.graph.group?`$select=id,displayName")
    Csv 'GroupsRemaining' ($left|ForEach-Object{[pscustomobject]@{Id=$_.id;Name=$_.displayName}}) @('Id','Name')
   }
  }
  Log "Run finished. Review Actions.csv and remaining-item reports. Output: $script:RunDir"
  [Windows.Forms.MessageBox]::Show($form,"Run finished. Some actions may be skipped or failed; review the log and CSVs.`n$script:RunDir",'Reports ready')|Out-Null
 }catch{Log "STOPPED: $($_.Exception.Message)";[Windows.Forms.MessageBox]::Show($form,$_.Exception.Message,'Stopped')|Out-Null}
 finally{
  if($script:ManualGraphToken){try{Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null}catch{};$script:ManualGraphToken=$null}
  $hybridBox.Enabled=$true;$oneDriveBox.Enabled=$true;$forwardBox.Enabled=$true;$mailAccessBox.Enabled=$true
  $testBox.Enabled=$true;$run.Enabled=$true;$upnBox.Enabled=$true;$folderBox.Enabled=$true;$appBox.Enabled=$true
  foreach($cb in $options.Values){$cb.Enabled=$true}
 }
})
$form.Add_FormClosing({param($sender,$eventArgs) if(-not $run.Enabled){$eventArgs.Cancel=$true}})
[Windows.Forms.Application]::EnableVisualStyles()
[Windows.Forms.Application]::Run($form)
