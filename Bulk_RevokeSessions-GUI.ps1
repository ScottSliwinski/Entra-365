#Requires -Version 7.0
<# Windows PowerShell 7 GUI. CSV: UserPrincipalName; optional OnPremUPN.
Requires RSAT Active Directory tools for optional AD changes.
Run: pwsh.exe -STA -File .\Revoke-EntraSessions-GUI.ps1
AD updates and cloud revocation are independent; both outcomes are exported.
#>
$ErrorActionPreference='Stop'
if (-not $IsWindows) { throw 'Requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run with pwsh.exe -STA -File .\Revoke-EntraSessions-GUI.ps1' }
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$script:busy=$false; $script:tenant=$null; $script:adArgs=$null; $script:results=@()
$form=[Windows.Forms.Form]::new(); $form.Text='Entra Bulk Sessions + On-prem Password Change'; $form.ClientSize=[Drawing.Size]::new(1120,750); $form.MinimumSize=$form.Size; $form.StartPosition='CenterScreen'
function Control($type,$text,$x,$y,$w,$h=28) {
 $c=New-Object "System.Windows.Forms.$type"; $c.Text=$text; $c.SetBounds($x,$y,$w,$h); $form.Controls.Add($c); return $c
}
$null=Control Label 'Tenant ID or verified domain:' 15 15 260
$tenantBox=Control TextBox '' 15 43 420
$connect=Control Button 'Connect to Entra' 450 40 180 32
$device=Control CheckBox 'Device-code sign-in (see console)' 650 43 440
$connection=Control Label 'Entra: not connected' 15 82 1080
$adOption=Control CheckBox 'Also require on-prem AD password change at next logon' 15 119 580
$null=Control Label 'Writable domain controller (FQDN):' 15 157 320
$server=Control TextBox '' 15 185 420
$adConnect=Control Button 'Connect / Test AD' 450 182 180 32
$alternate=Control CheckBox 'Prompt for AD credentials' 650 185 400
$adLabel=Control Label 'AD: not connected. Uses RSAT ActiveDirectory module and AD Web Services.' 15 222 1080
$null=Control Label 'AD setting does not guarantee a password prompt in Microsoft 365; hybrid configuration determines cloud behavior.' 15 254 1080
$browse=Control Button 'Browse / Import CSV' 15 290 180 32
$template=Control Button 'Save CSV Template' 210 290 180 32
$pathLabel=Control Label 'No CSV selected' 410 294 680
$grid=Control DataGridView '' 15 338 1090 270
$grid.ReadOnly=$true; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.RowHeadersVisible=$false; $grid.AutoSizeColumnsMode='Fill'; $grid.SelectionMode='FullRowSelect'; $grid.Anchor='Top,Bottom,Left,Right'
foreach($n in 'UserPrincipalName','OnPremUPN','SessionStatus','ADStatus','Error') { $null=$grid.Columns.Add($n,$n) }
$progress=Control ProgressBar '' 15 625 1090 22; $progress.Anchor='Bottom,Left,Right'
$status=Control Label 'Import CSV, connect, then run.' 15 660 1090; $status.Anchor='Bottom,Left,Right'
$run=Control Button 'Run selected actions' 15 701 180 32; $run.Anchor='Bottom,Left'
$export=Control Button 'Export Results' 210 701 180 32; $export.Anchor='Bottom,Left'
function Note($s) { $status.Text=$s; [Windows.Forms.Application]::DoEvents() }
function ErrorBox($s) { [void][Windows.Forms.MessageBox]::Show($form,$s,'Error','OK','Error') }
function Busy([bool]$b) {
 $script:busy=$b
 foreach($c in @($tenantBox,$connect,$device,$adOption,$server,$adConnect,$alternate,$browse,$template)) { $c.Enabled=-not $b }
 $run.Enabled=(-not $b -and $script:tenant -and $grid.Rows.Count -gt 0 -and (-not $adOption.Checked -or $script:adArgs))
 $export.Enabled=(-not $b -and $script:results.Count -gt 0)
}
function SaveResults {
 if (-not $script:results.Count) {return}
 $d=[Windows.Forms.SaveFileDialog]::new(); $d.Filter='CSV (*.csv)|*.csv'; $d.FileName="RevokeSessions-Results-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
 try { if($d.ShowDialog($form) -eq 'OK') { $script:results | Export-Csv -LiteralPath $d.FileName -NoTypeInformation -Encoding utf8; $status.Text="Saved: $($d.FileName)" } } finally {$d.Dispose()}
}
function EscapeLdap([string]$s) { $s.Replace('\','\5c').Replace('*','\2a').Replace('(','\28').Replace(')','\29').Replace([string][char]0,'\00') }
$connect.Add_Click({
 if($script:busy){return}; Busy $true; $script:tenant=$null
 try {
  if(-not $tenantBox.Text.Trim()){throw 'Enter a tenant ID or verified domain.'}
  Note 'Loading Graph modules and connecting...'
  foreach($m in 'Microsoft.Graph.Authentication','Microsoft.Graph.Users.Actions') {
   if(-not(Get-Module -ListAvailable $m)){Install-Module $m -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop}
   Import-Module $m -ErrorAction Stop
  }
  if(Get-MgContext){Disconnect-MgGraph | Out-Null}
  $p=@{TenantId=$tenantBox.Text.Trim();Scopes=@('User.RevokeSessions.All');ContextScope='Process';NoWelcome=$true;ErrorAction='Stop'}
  if($device.Checked){$p.UseDeviceCode=$true; Note 'Follow device-code instructions in PowerShell.'}
  Connect-MgGraph @p
  $ctx=Get-MgContext; if(-not $ctx.TenantId){throw 'No connected tenant returned.'}
  $script:tenant=$ctx.TenantId; $connection.Text="Connected: $($ctx.Account) | Tenant: $($ctx.TenantId)"
 }catch{$connection.Text='Entra: not connected'; ErrorBox $_.Exception.Message}finally{Busy $false}
})
$adConnect.Add_Click({
 if($script:busy){return}; Busy $true; $script:adArgs=$null
 try {
  if(-not $server.Text.Trim()){throw 'Enter a writable domain controller FQDN.'}
  Note 'Loading ActiveDirectory and testing AD connection...'
  $env:ADPS_LoadDefaultDrive='0'
  if(-not(Get-Module -ListAvailable ActiveDirectory)){throw 'Install RSAT Active Directory tools on this computer, then retry. Windows client (elevated): Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'}
  Import-Module ActiveDirectory -ErrorAction Stop
  $a=@{Server=$server.Text.Trim();ErrorAction='Stop'}
  if($alternate.Checked){$cred=Get-Credential -Message 'On-prem AD administrator / delegated credentials'; if(-not $cred){throw 'Credentials canceled.'}; $a.Credential=$cred}
  $root=Get-ADRootDSE @a
  $dc=Get-ADDomainController -Identity $root.dnsHostName @a
  if($dc.IsReadOnly){throw 'Choose a writable domain controller.'}
  $a.Server=[string]$root.dnsHostName; $script:adArgs=$a
  $adLabel.Text="AD connected: $($a.Server) | $($root.defaultNamingContext)"
 }catch{$adLabel.Text='AD: not connected';ErrorBox $_.Exception.Message}finally{Busy $false}
})
$server.Add_TextChanged({$script:adArgs=$null; $adLabel.Text='AD server changed; reconnect.'; Busy $false})
$alternate.Add_CheckedChanged({$script:adArgs=$null; $adLabel.Text='AD credentials option changed; reconnect.'; Busy $false})
$adOption.Add_CheckedChanged({Busy $false})
$browse.Add_Click({
 if($script:busy){return}; $d=[Windows.Forms.OpenFileDialog]::new();$d.Filter='CSV (*.csv)|*.csv'
 try {
  if($d.ShowDialog($form) -ne 'OK'){return}
  $data=@(Import-Csv -LiteralPath $d.FileName)
  if(-not $data.Count -or 'UserPrincipalName' -notin $data[0].PSObject.Properties.Name){throw 'CSV requires UserPrincipalName and at least one account.'}
  $map=@{}; $seenAD=@{}
  foreach($r in $data){
   $upn=([string]$r.UserPrincipalName).Trim(); if(-not $upn){continue}
   $adupn=([string]$r.OnPremUPN).Trim(); if(-not $adupn){$adupn=$upn}
   foreach($u in @($upn,$adupn)){if($u -notmatch '^[^@\s]+@[^@\s]+$'){throw "Invalid UPN: $u"}}
   if($map.ContainsKey($upn) -and $map[$upn] -ine $adupn){throw "Conflicting AD mapping for $upn"}
   if($seenAD.ContainsKey($adupn) -and $seenAD[$adupn] -ine $upn){throw "Multiple cloud accounts map to $adupn"}
   $map[$upn]=$adupn; $seenAD[$adupn]=$upn
  }
  if(-not $map.Count){throw 'No accounts found.'}
  $grid.Rows.Clear(); foreach($u in ($map.Keys | Sort-Object)){$null=$grid.Rows.Add($u,$map[$u],'Ready','Ready','')}
  $script:results=@();$progress.Value=0; $pathLabel.Text=$d.FileName;Note "$($map.Count) unique accounts loaded.";Busy $false
 }catch{ErrorBox $_.Exception.Message}finally{$d.Dispose()}
})
$template.Add_Click({
 $d=[Windows.Forms.SaveFileDialog]::new();$d.Filter='CSV (*.csv)|*.csv';$d.FileName='RevokeSessions-Template.csv'
 try{if($d.ShowDialog($form) -eq 'OK'){'UserPrincipalName,OnPremUPN'|Set-Content -LiteralPath $d.FileName -Encoding utf8}}catch{ErrorBox $_.Exception.Message}finally{$d.Dispose()}
})
$run.Add_Click({
 if($script:busy){return}
 try {
  $ctx=Get-MgContext;if(-not $script:tenant -or $ctx.TenantId -ne $script:tenant){throw 'Reconnect to Entra before running.'}
  $doAD=$adOption.Checked
  if($doAD -and -not $script:adArgs){throw 'Connect / Test AD first.'}
  $extra=if($doAD){"`nAlso set 'Must change password at next logon' on AD server $($script:adArgs.Server)."}else{''}
  if([Windows.Forms.MessageBox]::Show($form,"Revoke sessions for all $($grid.Rows.Count) previewed accounts?`nTenant: $script:tenant$extra`n`nAD failures do not prevent session revocation. Results show each action separately.",'Confirm bulk actions','YesNo','Warning','Button2') -ne 'Yes'){return}
  Busy $true; $script:results=@();$progress.Value=0;$progress.Maximum=$grid.Rows.Count
  foreach($row in $grid.Rows){
   $upn=[string]$row.Cells['UserPrincipalName'].Value; $adupn=[string]$row.Cells['OnPremUPN'].Value
   Note "Processing $($progress.Value+1) of $($grid.Rows.Count): $upn"
   $adStatus='Not requested';$adError='';$sessionStatus='Failed';$sessionError='';$guid=''
   # Apply the AD flag before revocation. Never use a UPN prefix as an identity.
   if($doAD){
    try{
     $a=$script:adArgs; $escaped=EscapeLdap $adupn
     $users=@(Get-ADUser -LDAPFilter "(userPrincipalName=$escaped)" -Properties pwdLastSet,PasswordNeverExpires,CannotChangePassword @a)
     if($users.Count -ne 1){throw "Expected exactly one AD account for $adupn; found $($users.Count)."}
     $u=$users[0];if($u.UserPrincipalName -ine $adupn){throw 'AD UPN mismatch.'};$guid=[string]$u.ObjectGUID
     if($u.PasswordNeverExpires){throw 'Password never expires is enabled. Review this account in AD; no password policy flags were changed.'}
     if($u.CannotChangePassword){throw 'User cannot change password is enabled. Review this account in AD.'}
     if([long]$u.pwdLastSet -ne 0){Set-ADUser -Identity $u.ObjectGUID -ChangePasswordAtLogon $true @a}
     $check=Get-ADUser -Identity $u.ObjectGUID -Properties pwdLastSet @a
     if([long]$check.pwdLastSet -ne 0){throw 'AD write could not be verified (pwdLastSet is not zero).'}
     $adStatus='Success (verified)'
    }catch{$adStatus='Failed';$adError=$_.Exception.Message}
   }
   try{
    $current=Get-MgContext;if($current.TenantId -ne $script:tenant){throw 'Graph tenant changed; reconnect.'}
    # Inspect the HTTP result directly; SDK versions may wrap the Boolean in Value.
    $encodedUPN=[Uri]::EscapeDataString($upn)
    $http=$null
    try {
     $http=Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$encodedUPN/revokeSignInSessions" -OutputType HttpResponseMessage -ErrorAction Stop
     if(-not $http.IsSuccessStatusCode){throw "Graph returned HTTP $([int]$http.StatusCode)."}
     $body=$http.Content.ReadAsStringAsync().GetAwaiter().GetResult()
     if(-not [string]::IsNullOrWhiteSpace($body)){
      $payload=ConvertFrom-Json -InputObject $body -ErrorAction Stop
      if($payload.PSObject.Properties['value'] -and $payload.value -eq $false){throw 'Graph explicitly returned value=false.'}
     }
     $sessionStatus='Success'
    } finally {if($null -ne $http){$http.Dispose()}}
   }catch{$sessionError=$_.Exception.Message}
   $errors=@();if($adError){$errors+="AD: $adError"};if($sessionError){$errors+="Entra: $sessionError"}
   $row.Cells['SessionStatus'].Value=$sessionStatus;$row.Cells['ADStatus'].Value=$adStatus;$row.Cells['Error'].Value=$errors -join ' | '
   $script:results += [pscustomobject]@{UserPrincipalName=$upn;OnPremUPN=$adupn;TenantId=$script:tenant;ADServer=$(if($doAD){$script:adArgs.Server}else{''});ADObjectGUID=$guid;SessionStatus=$sessionStatus;SessionError=$sessionError;ADStatus=$adStatus;ADError=$adError;TimeUTC=[datetime]::UtcNow.ToString('o')}
   $progress.Value++;$grid.Refresh()
  }
  $s=@($script:results|Where-Object SessionStatus -eq 'Success').Count
  $ac=@($script:results|Where-Object ADStatus -eq 'Success (verified)').Count
  Note "Complete: sessions accepted $s/$($script:results.Count); AD flags verified $ac. Review errors, then export."
  SaveResults
 }catch{ErrorBox $_.Exception.Message}finally{Busy $false}
})
$export.Add_Click({try{SaveResults}catch{ErrorBox $_.Exception.Message}})
$form.Add_FormClosing({param($sender,$e) if($script:busy){$e.Cancel=$true}})
Busy $false
try{[void]$form.ShowDialog()}finally{$script:adArgs=$null;$form.Dispose()}
