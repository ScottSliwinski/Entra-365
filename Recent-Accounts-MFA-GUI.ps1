#requires -Version 5.1
<#
Windows GUI: selectable account creation windows, per-user MFA disabled.
Checkboxes: last 30, last 60, last 90 days (cumulative), or older than 90 days.
Multiple selections combine the windows; 90 days is checked by default.
Run in a fresh PowerShell console (not ISE):
  powershell.exe -STA -File .\Recent-Accounts-MFA-GUI.ps1
  pwsh.exe -STA -File .\Recent-Accounts-MFA-GUI.ps1
Prerequisites:
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
Exchange role: permission to read all mailboxes. ExchangeOnlineManagement 3.0.0 or later required.
Role: Authentication Policy Administrator (or a supported higher role).
Admin consent: User.Read.All, Policy.Read.All, Policy.ReadWrite.AuthenticationMethod.
Uses Microsoft Graph BETA for per-user MFA requirements; API subject to change.
Disabled per-user MFA does NOT establish absence of MFA via Conditional Access
or Security Defaults. Do not use per-user enabling where CA manages MFA.
Scan includes all tenant domains and disabled accounts; Guest accounts and Shared Mailboxes are excluded. Only enabled
Member accounts with a verified disabled MFA state can be selected for changes.
CSV includes EVERY scanned matching non-Guest, non-shared-mailbox account, including lookup failures.
Enabling sets 'enabled', not 'enforced'; it does not register an MFA method.
No tenant changes occur during scanning. Graph login occurs in the console/browser.
#>
[CmdletBinding()]
param([ValidateSet(30,60,90)][int]$Days = 90, [string]$TenantId, [switch]$DeviceCode)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This GUI requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run with powershell.exe -STA -File or pwsh.exe -STA -File.' }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    throw 'Install prerequisite: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser; then open a fresh console.'
}
if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
    throw 'Install prerequisite: Install-Module ExchangeOnlineManagement -Scope CurrentUser; then open a fresh console.'
}
Import-Module Microsoft.Graph.Authentication
Import-Module ExchangeOnlineManagement -MinimumVersion 3.0.0
$scopes = @('User.Read.All','Policy.Read.All','Policy.ReadWrite.AuthenticationMethod')
# Always authenticate explicitly in this process; no silent reuse of another tenant.
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
$connect = @{ Scopes=$scopes; ContextScope='Process'; NoWelcome=$true }
if ($TenantId) { $connect.TenantId=$TenantId }
if ($DeviceCode) { $connect.UseDeviceAuthentication=$true }
Write-Host 'Signing in to Graph. Complete browser login (or use the device code shown HERE).' -ForegroundColor Cyan
Connect-MgGraph @connect
$context = Get-MgContext
if (-not $context.TenantId) { throw 'Graph connection did not return a tenant.' }
Write-Host 'Signing in to Exchange Online. Use an administrator in the SAME tenant as Graph.' -ForegroundColor Cyan
# This tool runs in a fresh console; remove stale Exchange connections before sign-in.
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
$exoConnect = @{ShowBanner=$false; ErrorAction='Stop'}
if ($context.Account) { $exoConnect.UserPrincipalName=[string]$context.Account }
if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) { $exoConnect.DisableWAM=$true }
if ($DeviceCode -and (Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('Device') -and $PSVersionTable.PSVersion.Major -ge 7) {
    $exoConnect.Remove('DisableWAM'); $exoConnect.Device=$true
}
try { Connect-ExchangeOnline @exoConnect } catch {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    throw
}
function Get-SharedMailboxIndex {
    # Read the tenant GUID from the active REST connection, not organization metadata.
    $connections=@(Get-ConnectionInformation -ErrorAction Stop | Where-Object {
        $_.State -eq 'Connected' -and -not $_.IsEopSession
    })
    if ($connections.Count -ne 1) {
        throw "Expected one active Exchange Online connection; found $($connections.Count). Close this console and run the script in a fresh console."
    }
    $exchangeTenant=[string]$connections[0].TenantID
    if ([string]::IsNullOrWhiteSpace($exchangeTenant)) {
        throw 'Exchange did not return a tenant ID. Update ExchangeOnlineManagement, then rerun from a fresh console.'
    }
    if ($exchangeTenant.Trim() -ne ([string]$context.TenantId).Trim()) {
        throw "Tenant mismatch confirmed. Graph tenant: $($context.TenantId). Exchange tenant: $exchangeTenant. Close the console and sign into the same organization for both connections."
    }
    $ids=@{}
    foreach($mailbox in @(Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties ExternalDirectoryObjectId -ErrorAction Stop)) {
        $id=[string]$mailbox.ExternalDirectoryObjectId
        if (-not $id) { throw 'A shared mailbox has no Entra object ID. Scan stopped because exclusions cannot be verified.' }
        $ids[$id]=$true
    }
    return $ids
}

$form = New-Object Windows.Forms.Form
$form.Text = "Recent accounts - per-user MFA | Tenant $($context.TenantId)"
$form.Size = New-Object Drawing.Size(1180,790)
$form.MinimumSize = New-Object Drawing.Size(1000,670)
$form.StartPosition = 'CenterScreen'
$note = New-Object Windows.Forms.Label
$note.SetBounds(15,12,1130,55)
$note.Text = "Selected creation windows, all domains, Guests and Shared Mailboxes excluded. Disabled here means PER-USER MFA only; Conditional Access / Security Defaults may already protect these users.`r`nOnly enabled Member accounts can be selected. Scan errors are shown and exported, never treated as MFA disabled."
$note.Anchor = 'Top,Left,Right'
$form.Controls.Add($note)
$rangeChecks=@{}
$x=15
foreach($item in @(@('30','30 Days (last 30)'),@('60','60 Days (last 60)'),@('90','90 Days (last 90)'),@('Older','Over 90 Days'))) {
    $cb=New-Object Windows.Forms.CheckBox
    $cb.Text=$item[1]; $cb.SetBounds($x,70,180,26); $x+=185
    $cb.Checked=($item[0] -eq [string]$Days)
    $rangeChecks[$item[0]]=$cb; $form.Controls.Add($cb)
}
$rangeHint=New-Object Windows.Forms.Label
$rangeHint.Text='30 / 60 / 90 are cumulative windows. Multiple checks combine results. Change selections, then click Scan accounts.'
$rangeHint.SetBounds(15,98,1130,22); $rangeHint.Anchor='Top,Left,Right'; $form.Controls.Add($rangeHint)
$buttons = @{}
$x = 15
foreach ($item in @(@('Scan','Scan accounts'),@('All','Select All'),@('None','Clear Selection'),@('Apply','Enable selected MFA'),@('Export','Export CSV'))) {
    $b=New-Object Windows.Forms.Button
    $b.Text=$item[1]; $b.SetBounds($x,123,175,32); $x+=185
    $form.Controls.Add($b); $buttons[$item[0]]=$b
}
$grid = New-Object Windows.Forms.DataGridView
$grid.SetBounds(15,170,1130,495); $grid.Anchor='Top,Bottom,Left,Right'
$grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false
$grid.AutoSizeColumnsMode='DisplayedCells'; $grid.RowHeadersVisible=$false
$grid.SelectionMode='FullRowSelect'; $grid.MultiSelect=$false
$check=New-Object Windows.Forms.DataGridViewCheckBoxColumn
$check.Name='Select'; $check.HeaderText='Enable'; [void]$grid.Columns.Add($check)
foreach ($name in @('UPN','DisplayName','CreatedUTC','AccountEnabled','UserType','Before','After','Result','Error')) {
    $col=New-Object Windows.Forms.DataGridViewTextBoxColumn
    $col.Name=$name; $col.HeaderText=$name; $col.ReadOnly=$true
    $col.SortMode='NotSortable'; [void]$grid.Columns.Add($col)
}
$form.Controls.Add($grid)
$progress=New-Object Windows.Forms.ProgressBar
$progress.SetBounds(15,678,1130,20); $progress.Anchor='Bottom,Left,Right'; $form.Controls.Add($progress)
$status=New-Object Windows.Forms.Label
$status.SetBounds(15,706,1130,40); $status.Anchor='Bottom,Left,Right'; $status.Text='Click Scan accounts to start.'; $form.Controls.Add($status)
$script:records=New-Object 'System.Collections.Generic.List[object]'
$script:busy=$false; $script:scanComplete=$false; $script:lastCsv=$null; $script:scanWindow=''
function Set-Busy([bool]$value) {
    $script:busy=$value
    foreach($b in $buttons.Values) { $b.Enabled=-not $value }
    foreach($cb in $rangeChecks.Values) { $cb.Enabled=-not $value }
    $grid.Enabled=-not $value
}
function Show-Error($message) { [void][Windows.Forms.MessageBox]::Show($form,[string]$message,'MFA tool','OK','Error') }
function Update-Progress($i,$total,$message) {
    $progress.Value=if($total -gt 0){ [Math]::Min(100,[int](100*$i/$total)) }else{0}
    $status.Text=$message; [Windows.Forms.Application]::DoEvents()
}
function Get-State($id) {
    $response=Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$id/authentication/requirements" -OutputType PSObject
    $state=[string]$response.perUserMfaState
    if($state -notin @('disabled','enabled','enforced')) { throw "Unexpected MFA state '$state'." }
    return $state
}
function Render-Records {
    $grid.Rows.Clear()
    foreach($r in $script:records) {
        if ($r.Before -eq 'disabled' -or $r.Result -eq 'Scan failed') {
            $index=$grid.Rows.Add($false,$r.UPN,$r.DisplayName,$r.CreatedUTC,$r.AccountEnabled,$r.UserType,$r.Before,$r.After,$r.Result,$r.Error)
            $row=$grid.Rows[$index]; $row.Tag=$r
            $eligible=($r.Before -eq 'disabled' -and $r.AccountEnabled -and $r.UserType -eq 'Member' -and $r.Result -in @('Not selected','Enable failed'))
            $row.Cells['Select'].ReadOnly=-not $eligible
            if(-not $eligible){$row.DefaultCellStyle.BackColor=[Drawing.Color]::Gainsboro}
        }
    }
}
function Save-Report([string]$path) {
    $script:records | Select-Object TenantId,UserId,UPN,DisplayName,CreatedUTC,CreationWindow,AccountEnabled,UserType,Before,After,Selected,Result,Error,ScanUTC,ActionUTC |
        Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    # Write headers even when there are no recent accounts.
    if ($script:records.Count -eq 0) {
        '"TenantId","UserId","UPN","DisplayName","CreatedUTC","CreationWindow","AccountEnabled","UserType","Before","After","Selected","Result","Error","ScanUTC","ActionUTC"' | Set-Content -LiteralPath $path -Encoding UTF8
    }
    $script:lastCsv=$path
}
function Choose-Csv {
    $dialog=New-Object Windows.Forms.SaveFileDialog
    $dialog.Filter='CSV files (*.csv)|*.csv'; $dialog.DefaultExt='csv'
    $dialog.FileName="RecentAccounts-MFA-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
    $dialog.InitialDirectory=[Environment]::GetFolderPath('MyDocuments')
    try { if($dialog.ShowDialog($form) -eq 'OK'){return $dialog.FileName} } finally {$dialog.Dispose()}
}
$grid.Add_CurrentCellDirtyStateChanged({if($grid.IsCurrentCellDirty){[void]$grid.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit)}})
foreach($cb in $rangeChecks.Values) {
    $cb.Add_CheckedChanged({
        if(-not $script:busy) {
            $script:scanComplete=$false; $script:records.Clear(); $grid.Rows.Clear()
            $status.Text='Date selection changed. Click Scan accounts to refresh results.'
        }
    })
}
$buttons.Scan.Add_Click({
    $recentDays=@(30,60,90 | Where-Object { $rangeChecks[[string]$_].Checked })
    $includeOlder=$rangeChecks['Older'].Checked
    if($recentDays.Count -eq 0 -and -not $includeOlder) { Show-Error 'Check at least one creation window.'; return }
    $maxRecent=if($recentDays.Count -gt 0){($recentDays | Measure-Object -Maximum).Maximum}else{0}
    $labels=@($recentDays | ForEach-Object { "Last $_ days" })
    if($includeOlder){$labels+='Over 90 days'}
    $script:scanWindow=$labels -join '; '
    Set-Busy $true; $script:scanComplete=$false; $script:records.Clear(); $grid.Rows.Clear()
    try {
        Update-Progress 0 1 'Loading shared mailbox exclusions from Exchange Online...'
        $sharedIds=Get-SharedMailboxIndex
        $excludedShared=0
        # Use one UTC instant for all date boundaries and the report.
        $end=[DateTime]::UtcNow
        $stamp=$end.ToString('o'); $cutoff90=$end.AddDays(-90)
        $recentStart=$end.AddDays(-$maxRecent)
        $uri='https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,createdDateTime,accountEnabled,userType&$top=999'
        # Retrieve a superset, then apply exact boundaries locally before MFA lookups.
        if(-not $includeOlder) {
            $filter=[Uri]::EscapeDataString("createdDateTime ge $($recentStart.ToString('yyyy-MM-ddTHH:mm:ssZ'))")
            $uri+='&$filter='+$filter
        }
        $users=New-Object 'System.Collections.Generic.List[object]'
        do {
            Update-Progress 0 1 "Loading accounts for selected windows: $($users.Count) retrieved..."
            $page=Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
            foreach($u in $page.value){
                # Exclude Guests and shared mailbox backing accounts before MFA lookups or report collection.
                if($sharedIds.ContainsKey([string]$u.id)) { $excludedShared++; continue }
                if($u.userType -eq 'Guest' -or -not $u.createdDateTime){continue}
                $created=[DateTimeOffset]::Parse($u.createdDateTime).UtcDateTime
                $matchesRecent=($maxRecent -gt 0 -and $created -ge $recentStart -and $created -le $end)
                $matchesOlder=($includeOlder -and $created -lt $cutoff90)
                if($matchesRecent -or $matchesOlder){$users.Add($u)}
            }
            $uri=$page.'@odata.nextLink'
        } while($uri)
        $i=0
        foreach($u in ($users | Sort-Object createdDateTime -Descending)) {
            $i++; Update-Progress $i $users.Count "Checking MFA $i / $($users.Count): $($u.userPrincipalName)"
            $r=[pscustomobject]@{TenantId=$context.TenantId;UserId=$u.id;UPN=$u.userPrincipalName;DisplayName=$u.displayName;CreatedUTC=$u.createdDateTime;CreationWindow=$script:scanWindow;AccountEnabled=[bool]$u.accountEnabled;UserType=$u.userType;Before='Unknown';After='Unknown';Selected=$false;Result='Scan failed';Error='';ScanUTC=$stamp;ActionUTC=''}
            try {
                $r.Before=Get-State $u.id; $r.After=$r.Before
                $r.Result=if($r.Before -ne 'disabled'){'Already enabled/enforced'}elseif(-not $r.AccountEnabled -or $r.UserType -ne 'Member'){'Skipped: disabled account or non-Member'}else{'Not selected'}
            } catch {$r.Error=$_.Exception.Message}
            $script:records.Add($r)
        }
        $script:scanComplete=$true; Render-Records
        $errors=@($script:records | Where-Object Result -eq 'Scan failed').Count
        $status.Text="Scan complete. $($users.Count) matching accounts ($script:scanWindow); $($grid.Rows.Count) disabled MFA / error rows. Shared mailboxes excluded: $excludedShared. Lookup errors: $errors. Export CSV for the full audit."
    } catch {Show-Error $_.Exception.Message; $status.Text='Scan incomplete. Resolve the error and scan again.'}
    finally {Set-Busy $false}
})
$buttons.All.Add_Click({foreach($row in $grid.Rows){if(-not $row.Cells['Select'].ReadOnly){$row.Cells['Select'].Value=$true}}})
$buttons.None.Add_Click({foreach($row in $grid.Rows){$row.Cells['Select'].Value=$false}})
$buttons.Export.Add_Click({
    if(-not $script:scanComplete){Show-Error 'Complete a scan first.';return}
    try {foreach($row in $grid.Rows){$row.Tag.Selected=($row.Cells['Select'].Value -eq $true)}; $path=Choose-Csv; if($path){Save-Report $path; $status.Text="CSV saved: $path"}}catch{Show-Error $_.Exception.Message}
})
$buttons.Apply.Add_Click({
    if(-not $script:scanComplete){Show-Error 'Complete a scan first.';return}
    [void]$grid.EndEdit()
    $selected=@($grid.Rows | Where-Object { $_.Cells['Select'].Value -eq $true -and -not $_.Cells['Select'].ReadOnly } | ForEach-Object {$_.Tag})
    if($selected.Count -eq 0){Show-Error 'Check at least one eligible account.';return}
    $message="Enable PER-USER MFA for $($selected.Count) checked accounts in tenant $($context.TenantId)?`r`n`r`nDo not proceed if Conditional Access or Security Defaults manages MFA for these accounts. This sets Enabled and may prompt registration on their next modern sign-in.`r`n`r`nA CSV location is required before any changes."
    if([Windows.Forms.MessageBox]::Show($form,$message,'Confirm MFA changes','YesNo','Warning') -ne 'Yes'){return}
    $path=Choose-Csv; if(-not $path){return}
    Set-Busy $true
    try {
        Update-Progress 0 1 'Refreshing shared mailbox exclusions before changes...'
        $sharedIds=Get-SharedMailboxIndex
        # A mailbox converted since the scan is removed from both GUI and CSV.
        foreach($r in @($script:records.ToArray())) {
            if($sharedIds.ContainsKey([string]$r.UserId)) { [void]$script:records.Remove($r) }
        }
        $selected=@($selected | Where-Object { -not $sharedIds.ContainsKey([string]$_.UserId) })
        Render-Records
        foreach($r in $selected){$r.Selected=$true; $r.Result='Pending'}
        Save-Report $path # Ensure output is writable before PATCHing.
        $i=0
        foreach($r in $selected){
            $i++; Update-Progress $i $selected.Count "Enabling $i / $($selected.Count): $($r.UPN)"
            $patched=$false
            try {
                # Recheck both eligibility and state immediately before changing.
                $u=Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($r.UserId)?`$select=accountEnabled,userType" -OutputType PSObject
                $r.After=Get-State $r.UserId
                if(-not $u.accountEnabled -or $u.userType -ne 'Member'){$r.Result='Skipped: account eligibility changed'}
                elseif($r.After -ne 'disabled'){$r.Result='Skipped: MFA already enabled/enforced'}
                else {
                    $body=@{perUserMfaState='enabled'} | ConvertTo-Json -Compress
                    Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/beta/users/$($r.UserId)/authentication/requirements" -Body $body -ContentType 'application/json' | Out-Null
                    $patched=$true; $r.After='Unknown'; $r.Result='Enabled: verification pending'
                    $r.After=Get-State $r.UserId
                    $r.Result=if($r.After -in @('enabled','enforced')){'Enabled successfully'}else{'PATCH accepted: verification pending'}
                }
                $r.Error=''
            } catch {
                $r.Error=$_.Exception.Message
                $r.Result=if($patched){'PATCH accepted: verification failed'}else{'Enable failed'}
            }
            $r.ActionUTC=[DateTime]::UtcNow.ToString('o')
            Save-Report $path # Checkpoint each result; stop further changes if saving fails.
        }
        Render-Records
        $status.Text="Finished. Results saved: $path"
        [void][Windows.Forms.MessageBox]::Show($form,"Finished. Check Result and Error in the CSV.`r`n$path",'Results exported','OK','Information')
    }catch{Render-Records; Show-Error "Processing stopped: $($_.Exception.Message)`r`nLast report: $path"}
    finally{Set-Busy $false}
})
$form.Add_FormClosing({param($sender,$e) if($script:busy){$e.Cancel=$true}})
try {[void]$form.ShowDialog()} finally {$form.Dispose(); Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue; Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null}
