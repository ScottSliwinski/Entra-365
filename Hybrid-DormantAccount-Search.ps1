#Requires -Version 7.2
<#
Hybrid Dormant Account Search v1.3 - read-only; Windows / PowerShell 7.2+.
Missing Graph / ThreadJob prerequisites install automatically for CurrentUser.
Windows RSAT AD tools install on first AD connection if absent (UAC required).
Internet / approved Windows feature source access is required for installation.
Run: pwsh.exe -STA -File .\Hybrid-DormantAccount-Search.ps1
Sign-in activity requires Entra P1/P2, consent for User.Read.All and
AuditLog.Read.All, and an appropriate role such as Reports Reader.
IMPORTANT: lastSignInDateTime is an interactive ATTEMPT (success OR failure).
lastSuccessfulSignInDateTime includes interactive AND non-interactive success.
Missing history is unknown, not proof of never signing in. No account changes.
#>
[CmdletBinding()]param()
$ErrorActionPreference='Stop'
if (-not $IsWindows) { throw 'Windows is required for the WPF interface.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Launch with pwsh.exe -STA -File <script path>.' }
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,Microsoft.Win32.Primitives
function Ensure-RequiredModule {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing missing prerequisite: $Name (CurrentUser)..." -ForegroundColor Cyan
        try {
            if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
                throw 'PowerShellGet is unavailable. Repair your PowerShell 7 installation to restore Install-Module.'
            }
            if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
                Register-PSRepository -Default -ErrorAction Stop
            }
            $install = @{
                Name = $Name
                Repository = 'PSGallery'
                Scope = 'CurrentUser'
                Force = $true
                AllowClobber = $true
                ErrorAction = 'Stop'
            }
            if ((Get-Command Install-Module).Parameters.ContainsKey('AcceptLicense')) {
                $install.AcceptLicense = $true
            }
            Install-Module @install
        } catch {
            throw "Could not install prerequisite '$Name'. Check internet/proxy access to PowerShell Gallery and your organization's module-install policy. Details: $($_.Exception.Message)"
        }
    }
    # A present but broken module is reported rather than silently overwritten.
    Import-Module $Name -ErrorAction Stop
}
function Ensure-ADModule {
    if (-not (Get-Module -ListAvailable ActiveDirectory)) {
        Write-Host 'Installing missing Windows RSAT Active Directory tools. Approve the Windows UAC prompt if shown.' -ForegroundColor Cyan
        # RSAT is a Windows component, not a PowerShell Gallery module.
        # Use Windows PowerShell for inbox Windows servicing cmdlets.
        $setup = @'
$ErrorActionPreference = 'Stop'
try {
    $os = Get-CimInstance Win32_OperatingSystem
    if ($os.ProductType -eq 1) {
        $r = Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ErrorAction Stop
        if ($r.RestartNeeded) { Write-Warning 'Restart Windows before using AD tools.' }
    } else {
        $r = Install-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction Stop
        if (-not $r.Success) { throw 'Windows Server RSAT installation failed.' }
        if ([string]$r.RestartNeeded -eq 'Yes') { Write-Warning 'Restart Windows before using AD tools.' }
    }
    exit 0
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    Read-Host 'Press Enter to close this installation error'
    exit 1
}
'@
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($setup))
        $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $process = Start-Process -FilePath $psExe -Verb RunAs -ArgumentList @('-NoProfile','-EncodedCommand',$encoded) -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -ne 0) {
            throw 'RSAT installation was cancelled or failed. Windows Update / your managed feature source must be available. Entra-only search is still available.'
        }
    }
    try { Import-Module ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue }
    catch { throw "AD tools could not load. A Windows restart may be needed after RSAT installation. Details: $($_.Exception.Message)" }
}
try {
    Ensure-RequiredModule 'Microsoft.Graph.Authentication'
    # Some PowerShell installations expose the built-in Microsoft.PowerShell.ThreadJob
    # rather than the Gallery module named ThreadJob. Accept either command provider.
    if (-not (Get-Command Start-ThreadJob -ErrorAction SilentlyContinue)) {
        Ensure-RequiredModule 'ThreadJob'
    }
} catch {
    [Windows.MessageBox]::Show($_.Exception.Message, 'Prerequisite setup failed', 'OK', 'Error') | Out-Null
    throw
}
[xml]$x=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Hybrid Dormant Account Search • v1.3" Width="1380" Height="870" MinWidth="1000" MinHeight="700" WindowStartupLocation="CenterScreen" Background="#F2F5F9">
<DockPanel Margin="16">
<StackPanel DockPanel.Dock="Top">
<TextBlock Text="Hybrid dormant account search" FontSize="25" FontWeight="Bold" Margin="0,0,0,10"/>
<WrapPanel><TextBlock Text="Tenant ID / verified domain:" VerticalAlignment="Center"/><TextBox x:Name="Tenant" Width="230" Margin="8"/><CheckBox x:Name="Device" Content="Use device-code sign-in" VerticalAlignment="Center"/><Button x:Name="Connect" Content="Connect Entra" Padding="12,5" Margin="12,5"/><TextBlock x:Name="CloudStatus" Text="Not connected" VerticalAlignment="Center"/></WrapPanel>
<WrapPanel><CheckBox x:Name="UseAD" Content="Include on-prem AD" VerticalAlignment="Center"/><TextBlock Text="DC / ADWS server:" Margin="15,0,5,0" VerticalAlignment="Center"/><TextBox x:Name="Server" Width="220"/><CheckBox x:Name="CurrentAD" Content="Current Windows credentials" IsChecked="True" VerticalAlignment="Center" Margin="10,0"/><Button x:Name="ADConnect" Content="Connect / test AD" Padding="12,5" Margin="12,5"/><TextBlock x:Name="ADStatus" Text="Current Windows credentials unless changed" VerticalAlignment="Center"/></WrapPanel>
<TextBlock Text="Organizational filters: case-insensitive contains; all filled fields must match. Blank means any." Margin="0,12,0,4"/>
<WrapPanel x:Name="FilterPanel"/>
<WrapPanel Margin="0,6"><TextBlock Text="UPN domains (optional):" VerticalAlignment="Center"/><TextBox x:Name="Domains" Width="570" Margin="8,0" ToolTip="Enter domains separated by commas, semicolons or spaces. Leading @ is optional. Exact UPN domain matches; blank searches all domains."/><TextBlock Text="Comma-separated; blank = all domains" VerticalAlignment="Center"/></WrapPanel>
<WrapPanel Margin="0,8"><TextBlock Text="No interactive attempt for:" VerticalAlignment="Center"/><ComboBox x:Name="Age" Width="115" Margin="8,0" SelectedIndex="0"><ComboBoxItem Content="3 months"/><ComboBoxItem Content="6 months"/><ComboBoxItem Content="1 year"/></ComboBox><ComboBox x:Name="State" Width="150" SelectedIndex="0"><ComboBoxItem Content="Enabled accounts"/><ComboBoxItem Content="Disabled accounts"/><ComboBoxItem Content="All accounts"/></ComboBox><CheckBox x:Name="Guests" Content="Include guests" Margin="12,0" VerticalAlignment="Center"/><CheckBox x:Name="Unknown" Content="Include missing history / unmatched" IsChecked="True" VerticalAlignment="Center"/><CheckBox x:Name="AllActivity" Content="Show active matches too" Margin="12,0" VerticalAlignment="Center"/></WrapPanel>
<WrapPanel><TextBlock Text="Organizational values from:" VerticalAlignment="Center"/><ComboBox x:Name="Source" SelectedIndex="0" Width="170" Margin="8,0"><ComboBoxItem Content="AD when matched"/><ComboBoxItem Content="Entra"/></ComboBox><Button x:Name="Search" Content="Search" Padding="20,6"/><Button x:Name="Cancel" Content="Cancel" IsEnabled="False" Padding="15,6" Margin="8,0"/><Button x:Name="Export" Content="Export CSV" IsEnabled="False" Padding="15,6"/></WrapPanel>
<TextBlock Text="Interactive timestamp includes FAILED attempts. Successful timestamp includes background sign-ins. Accounts under 30 days old are excluded. Results are review candidates; missing history is unknown. Times are UTC." TextWrapping="Wrap" Foreground="#8A4800" Margin="0,10"/>
<ProgressBar x:Name="Progress" Height="7" Minimum="0" Maximum="100"/>
<TextBlock x:Name="Status" Text="Connect Entra, optionally connect AD, then search." Margin="0,7" TextWrapping="Wrap"/>
</StackPanel>
<DataGrid x:Name="Results" AutoGenerateColumns="True" IsReadOnly="True" CanUserAddRows="False" EnableRowVirtualization="True" EnableColumnVirtualization="True" SelectionMode="Extended"/>
</DockPanel></Window>
'@
$w=[Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
$ui=@{}; foreach($n in 'Tenant Device Connect CloudStatus UseAD Server CurrentAD ADConnect ADStatus FilterPanel Domains Age State Guests Unknown AllActivity Source Search Cancel Export Progress Status Results'.Split(' ')) { $ui[$n]=$w.FindName($n) }
$fields=[ordered]@{Name='Name / UPN';JobTitle='Job title / role';Department='Department';Office='Office';Company='Company';Manager='Manager name / UPN';EmployeeId='Employee ID';City='City';State='State / province';Country='Country'}
$inputs=@{}
foreach($key in $fields.Keys){$p=[Windows.Controls.StackPanel]::new();$p.Margin='0,4,12,4';$l=[Windows.Controls.TextBlock]::new();$l.Text=$fields[$key];$t=[Windows.Controls.TextBox]::new();$t.Width=185;$p.Children.Add($l)|Out-Null;$p.Children.Add($t)|Out-Null;$ui.FilterPanel.Children.Add($p)|Out-Null;$inputs[$key]=$t}
$script:adCredential=$null;$script:adServer=$null;$script:tenantId=$null;$script:job=$null;$script:rows=@()
function Show-Error($e){[Windows.MessageBox]::Show([string]$e,'Hybrid Search','OK','Error')|Out-Null}
$ui.Connect.Add_Click({try{
 if(!$ui.Tenant.Text.Trim()){throw 'Enter the intended tenant ID or verified domain.'}
 $p=@{TenantId=$ui.Tenant.Text.Trim();Scopes=@('User.Read.All','AuditLog.Read.All');ContextScope='Process';NoWelcome=$true;ErrorAction='Stop'}
 if($ui.Device.IsChecked){$p.UseDeviceAuthentication=$true;[Windows.MessageBox]::Show('The device code will appear in the PowerShell window. Complete sign-in there.','Device sign-in')|Out-Null}
 Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null
 Connect-MgGraph @p
 $c=Get-MgContext;$script:tenantId=$c.TenantId;$ui.CloudStatus.Text="$($c.Account) | Tenant $($c.TenantId)"
}catch{$script:tenantId=$null;$ui.CloudStatus.Text='Not connected';Show-Error $_}})
$ui.ADConnect.Add_Click({try{
 if(!$ui.Server.Text.Trim()){throw 'Enter your AD domain controller FQDN.'}
 Ensure-ADModule
 $cred=$null;if(!$ui.CurrentAD.IsChecked){$cred=Get-Credential -Message 'Credentials for the AD server';if(!$cred){throw 'AD credential entry cancelled.'}}
 $p=@{Server=$ui.Server.Text.Trim();ErrorAction='Stop'};if($cred){$p.Credential=$cred}
 $root=Get-ADRootDSE @p
 $script:adCredential=$cred;$script:adServer=$p.Server;$ui.UseAD.IsChecked=$true;$ui.ADStatus.Text="$($root.defaultNamingContext) | $($p.Server)"
}catch{$script:adServer=$null;$ui.ADStatus.Text='AD connection failed';Show-Error $_}})
$worker={param($cfg,$cred)
$ErrorActionPreference='Stop'
Import-Module Microsoft.Graph.Authentication
function Notice($s){Write-Output @{Kind='Progress';Text=$s}}
function Pages($url){
 while($url){
  $reply=$null
  for($try=0;$try -lt 6;$try++){
   try{$reply=Invoke-MgGraphRequest -Method GET -Uri $url -OutputType Hashtable;break}catch{
    $status=0;try{$status=[int]$_.Exception.Response.StatusCode}catch{}
    if($status -notin @(429,500,502,503,504) -or $try -eq 5){throw}
    $delay=[math]::Min(60,[math]::Pow(2,$try+1));try{$h=$_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds;if($h){$delay=[math]::Min(120,$h)}}catch{}
    Start-Sleep -Seconds $delay
   }
  }
  foreach($v in $reply.value){Write-Output $v}
  $url=$reply['@odata.nextLink']
 }
}
function DateValue($s){if($s){return [datetimeoffset]::Parse([string]$s).UtcDateTime};return $null}
function Stamp($d){if($null -ne $d){$d.ToString('yyyy-MM-ddTHH:mm:ssZ')}else{''}}
try{
 if((Get-MgContext).TenantId -ne $cfg.Tenant){throw 'Entra context changed. Reconnect and search again.'}
 Notice 'Loading Entra users and activity (500 per page)...'
 $cloud=@(Pages 'https://graph.microsoft.com/v1.0/users?$top=500&$select=id,displayName,userPrincipalName,userType,accountEnabled,createdDateTime,jobTitle,department,officeLocation,companyName,employeeId,city,state,country,onPremisesSecurityIdentifier,onPremisesSyncEnabled,signInActivity')
 if($cloud.Count -and !(@($cloud | Where-Object {$null -ne $_.signInActivity}).Count)){throw 'No sign-in activity was returned for any user. Verify P1/P2 licensing, consent and Reports Reader role. No inactivity conclusions produced.'}
 Notice "Loaded $($cloud.Count) Entra users. Loading managers..."
 $managers=@{}
 foreach($m in (Pages 'https://graph.microsoft.com/v1.0/users?$top=999&$select=id&$expand=manager($select=id,displayName,userPrincipalName)')){if($m.manager){$managers[$m.id]="$($m.manager.displayName) <$($m.manager.userPrincipalName)>"}}
 $bySid=@{};$byUpn=@{};foreach($u in $cloud){if($u.onPremisesSecurityIdentifier){$bySid[[string]$u.onPremisesSecurityIdentifier]=$u};if($u.userPrincipalName){$byUpn[[string]$u.userPrincipalName]=$u}}
 $records=[Collections.Generic.List[object]]::new();$used=@{}
 if($cfg.UseAD){
  Notice 'Loading AD users and organizational properties...'
  Import-Module ActiveDirectory -WarningAction SilentlyContinue
  $p=@{Server=$cfg.Server;Filter='*';Properties=@('displayName','userPrincipalName','title','department','physicalDeliveryOfficeName','company','manager','employeeID','l','st','co','whenCreated');ErrorAction='Stop'};if($cred){$p.Credential=$cred}
  $ad=@(Get-ADUser @p);$dnMap=@{};foreach($a in $ad){$dnMap[$a.DistinguishedName]=$a}
  foreach($a in $ad){
   $u=$null;$match='AD only';$sid=[string]$a.SID
   if($bySid.ContainsKey($sid)){$u=$bySid[$sid];$match='SID'}
   elseif($a.UserPrincipalName -and $byUpn.ContainsKey($a.UserPrincipalName)){
    $candidate=$byUpn[$a.UserPrincipalName]
    if(!$candidate.onPremisesSecurityIdentifier){$u=$candidate;$match='UPN fallback - verify'}else{$match='SID conflict - verify'}
   }
   if($u){if($used.ContainsKey($u.id)){throw "Duplicate AD match for Entra UID $($u.id). Resolve identity ambiguity."};$used[$u.id]=$true}
   $manager=[string]$a.Manager;if($manager -and $dnMap.ContainsKey($manager)){$ma=$dnMap[$manager];$manager="$($ma.DisplayName) <$($ma.UserPrincipalName)>"}
   $records.Add(@{AD=$a;Cloud=$u;Match=$match;ADManager=$manager})
  }
 }
 foreach($u in $cloud){if(!$used.ContainsKey($u.id)){$records.Add(@{AD=$null;Cloud=$u;Match='Entra only';ADManager=''})}}
 $now=[datetime]::UtcNow;$cutoff=$now.AddMonths(-$cfg.Months);$out=[Collections.Generic.List[object]]::new();$i=0
 foreach($r in $records){
  $i++;if($i%250 -eq 0){Notice "Evaluating $i / $($records.Count)..."}
  $u=$r.Cloud;$a=$r.AD
  if($u.userType -eq 'Guest' -and !$cfg.Guests){continue}
  $enabled=if($u){[bool]$u.accountEnabled}else{[bool]$a.Enabled}
  if($cfg.State -eq 0 -and !$enabled){continue};if($cfg.State -eq 1 -and $enabled){continue}
  $adValues=$a -and $cfg.Source -eq 0
  $upn=if($u){$u.userPrincipalName}else{$a.UserPrincipalName}
  # Domains are typed by the operator; compare complete UPN suffixes, not substrings.
  $suffix=if($upn -match '@([^@]+)$'){$Matches[1]}else{''}
  if($cfg.Domains.Count -gt 0 -and $suffix -notin $cfg.Domains){continue}
  $vals=@{
   Name=if($adValues){$a.DisplayName}else{$u.displayName}
   JobTitle=if($adValues){$a.Title}else{$u.jobTitle}
   Department=if($adValues){$a.Department}else{$u.department}
   Office=if($adValues){$a.physicalDeliveryOfficeName}else{$u.officeLocation}
   Company=if($adValues){$a.Company}else{$u.companyName}
   Manager=if($adValues){$r.ADManager}else{if($u){$managers[$u.id]}else{''}}
   EmployeeId=if($adValues){$a.EmployeeID}else{$u.employeeId}
   City=if($adValues){$a.l}else{$u.city}
   State=if($adValues){$a.st}else{$u.state}
   Country=if($adValues){$a.co}else{$u.country}
   UPNSuffix=if($upn -match '@(.+)$'){$Matches[1]}else{''}
  }
  $ok=$true
  foreach($key in $cfg.Filters.Keys){$needle=$cfg.Filters[$key];if(!$needle){continue};$value=[string]$vals[$key];if($key -eq 'Name'){$value+=" $upn"};if($value.IndexOf($needle,[StringComparison]::OrdinalIgnoreCase) -lt 0){$ok=$false;break}}
  if(!$ok){continue}
  $last=DateValue $u.signInActivity.lastSignInDateTime;$success=DateValue $u.signInActivity.lastSuccessfulSignInDateTime;$non=DateValue $u.signInActivity.lastNonInteractiveSignInDateTime
  $created=DateValue $u.createdDateTime
  if(!$created -and $a.whenCreated){$created=$a.whenCreated.ToUniversalTime()}
  # Always exclude new accounts, including when missing history / active matches are shown.
  # Exactly 30 days old is eligible; an unavailable creation date is not assumed recent.
  if($created -and $created -gt $now.AddDays(-30)){continue}
  $status=if(!$u){'Unmatched AD - unknown'}elseif(!$last){'Missing interactive history - unknown'}elseif($last -lt $cutoff){'Dormant interactive candidate'}else{'Recent interactive attempt'}
  if($created -and $created -ge $cutoff -and !$last){$status='New account - missing history'}
  if($status -like '*unknown*' -or $status -like 'New account*'){if(!$cfg.Unknown){continue}}
  elseif($status -eq 'Recent interactive attempt' -and !$cfg.AllActivity){continue}
  $out.Add([pscustomobject][ordered]@{
   Classification=$status;DisplayName=$vals.Name;UPN=$upn;EntraUID=$u.id;IdentityMatch=$r.Match
   EntraEnabled=if($u){$u.accountEnabled}else{$null};ADEnabled=if($a){$a.Enabled}else{$null};UserType=$u.userType
   LastInteractiveAttemptUTC=Stamp $last;DaysSinceInteractiveAttempt=if($last){[int][math]::Floor(($now-$last).TotalDays)}else{$null}
   LastSuccessfulAnySignInUTC=Stamp $success;SuccessfulActivityInsideThreshold=if($success){$success -ge $cutoff}else{$null}
   LastNonInteractiveAttemptUTC=Stamp $non;CreatedUTC=Stamp $created;MinimumAccountAgeDays=30;CutoffUTC=Stamp $cutoff;ReportUTC=Stamp $now
   JobTitle=$vals.JobTitle;Department=$vals.Department;Office=$vals.Office;Company=$vals.Company;Manager=$vals.Manager;EmployeeID=$vals.EmployeeId
   City=$vals.City;State=$vals.State;Country=$vals.Country;OrganizationalSource=if($adValues){'AD'}else{'Entra'}
   ADSamAccountName=$a.SamAccountName;ADDistinguishedName=$a.DistinguishedName;TenantId=$cfg.Tenant;ADServer=$cfg.Server
  })
 }
 Write-Output @{Kind='Complete';Rows=$out.ToArray();Text="Complete: $($out.Count) matches; $($cloud.Count) Entra users; $($records.Count) combined identities. Cutoff: $(Stamp $cutoff)."}
}catch{Write-Output @{Kind='Error';Text=$_.Exception.Message}}
}
$timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[timespan]::FromMilliseconds(400)
function Set-Busy($busy){foreach($n in 'Search','Connect','ADConnect'){$ui[$n].IsEnabled=!$busy};$ui.Cancel.IsEnabled=$busy;$ui.Progress.IsIndeterminate=$busy}
$timer.Add_Tick({
 if(!$script:job){return}
 foreach($m in @(Receive-Job $script:job)){
  if($m.Kind -eq 'Progress'){$ui.Status.Text=$m.Text}
  elseif($m.Kind -eq 'Complete'){$script:rows=@($m.Rows);$ui.Results.ItemsSource=$script:rows;$ui.Export.IsEnabled=$script:rows.Count -gt 0;$ui.Status.Text=$m.Text}
  elseif($m.Kind -eq 'Error'){$ui.Status.Text="Search failed: $($m.Text)";Show-Error $m.Text}
 }
 if($script:job.State -in @('Completed','Failed','Stopped')){
  if($script:job.State -eq 'Failed'){$ui.Status.Text="Search failed: $($script:job.ChildJobs[0].JobStateInfo.Reason)"}
  Remove-Job $script:job -Force;$script:job=$null;$timer.Stop();Set-Busy $false
 }
})
$ui.Search.Add_Click({try{
 if(!$script:tenantId -or (Get-MgContext).TenantId -ne $script:tenantId){throw 'Connect to Entra first.'}
 if($ui.UseAD.IsChecked -and (!$script:adServer -or $ui.Server.Text.Trim() -ne $script:adServer)){throw 'Connect / test the selected AD server first.'}
 $domains=@($ui.Domains.Text -split '[,;\s]+' | ForEach-Object { $_.Trim().TrimStart('@').ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
 foreach($domain in $domains){
  if($domain.Length -gt 253 -or $domain -notmatch '^(?=.{1,253}$)[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)+$' -or @($domain.Split('.') | Where-Object {$_.Length -gt 63}).Count){
   throw "Invalid UPN domain '$domain'. Enter domain names only, separated by commas; for example: example.org, another.net."
  }
 }
 $f=@{};foreach($k in $inputs.Keys){$f[$k]=$inputs[$k].Text.Trim()}
 $cfg=@{Tenant=$script:tenantId;Server=$script:adServer;UseAD=[bool]$ui.UseAD.IsChecked;Months=@(3,6,12)[$ui.Age.SelectedIndex];State=$ui.State.SelectedIndex;Guests=[bool]$ui.Guests.IsChecked;Unknown=[bool]$ui.Unknown.IsChecked;AllActivity=[bool]$ui.AllActivity.IsChecked;Source=$ui.Source.SelectedIndex;Filters=$f;Domains=$domains}
 $script:rows=@();$ui.Results.ItemsSource=$null;$ui.Export.IsEnabled=$false;Set-Busy $true;$ui.Status.Text='Starting search...'
 $script:job=Start-ThreadJob -ScriptBlock $worker -ArgumentList $cfg,$script:adCredential;$timer.Start()
}catch{Set-Busy $false;Show-Error $_}})
$ui.Cancel.Add_Click({if($script:job){Stop-Job $script:job;$ui.Status.Text='Cancelled. No partial results exported.'}})
$ui.Export.Add_Click({try{$d=[Microsoft.Win32.SaveFileDialog]::new();$d.Filter='CSV files (*.csv)|*.csv';$d.FileName="Hybrid-DormantAccounts-$(Get-Date -Format yyyyMMdd-HHmmss).csv";if($d.ShowDialog()){
 # Neutralize spreadsheet formula injection in directory-controlled text.
 $safe=foreach($row in $script:rows){$h=[ordered]@{};foreach($p in $row.PSObject.Properties){$v=$p.Value;if($v -is [string] -and $v -match '^[\s]*[=+@-]'){$v="'"+$v};$h[$p.Name]=$v};[pscustomobject]$h}
 $safe|Export-Csv -LiteralPath $d.FileName -NoTypeInformation -Encoding utf8BOM;$ui.Status.Text="Exported $($script:rows.Count) rows to $($d.FileName)"
}}catch{Show-Error $_}})
$w.Add_Closed({$timer.Stop();if($script:job){Stop-Job $script:job;Remove-Job $script:job -Force}})
$null=$w.ShowDialog()
