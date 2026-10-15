#requires -Version 7.0
# Exchange Online Mailbox Delegate Manager - Windows / PowerShell 7
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$ErrorActionPreference = 'Stop'
$script:Mailbox = $null
$script:Rows = @()
$script:RecipientCache = @{}

function Resolve-Delegate([string]$Identity) {
    $id = $Identity.Trim()
    if ($script:RecipientCache.ContainsKey($id)) { return $script:RecipientCache[$id] }
    $result = [pscustomobject]@{ UPN=$id; Name=$id; Identity=$id; Resolved=$false }
    try {
        $r = Get-Recipient -Identity $id -ErrorAction Stop
        $upn = [string]$r.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($upn)) { $upn = [string]$r.PrimarySmtpAddress }
        if (-not [string]::IsNullOrWhiteSpace($upn)) {
            $result = [pscustomobject]@{UPN=$upn;Name=[string]$r.DisplayName;Identity=$id;Resolved=$true}
        }
    } catch { }
    $script:RecipientCache[$id] = $result
    return $result
}
function Log([string]$Message) {
    $log.AppendText("$(Get-Date -Format 'HH:mm:ss')  $Message`r`n")
    $log.SelectionStart=$log.TextLength; $log.ScrollToCaret()
}
function Invoke-UI([scriptblock]$Action) {
    try { $form.UseWaitCursor=$true; $status.Text='Working...'; [System.Windows.Forms.Application]::DoEvents(); & $Action }
    catch { Log "ERROR: $($_.Exception.Message)"; [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Operation failed','OK','Error') | Out-Null }
    finally { $form.UseWaitCursor=$false; $status.Text='Ready' }
}
function Find-Mailbox {
    $identity=$search.Text.Trim()
    if([string]::IsNullOrWhiteSpace($identity)){throw 'Enter the mailbox UPN or email address.'}
    $gridMail.Rows.Clear(); $gridDelegates.Rows.Clear(); $script:Rows=@(); $script:Mailbox=$null
    $mailLabel.Text='Selected mailbox: None'
    Log "Looking up mailbox: $identity"
    # Direct server-side lookup: never enumerate every mailbox in the tenant.
    $m=Get-EXOMailbox -Identity $identity -Properties UserPrincipalName,PrimarySmtpAddress,DisplayName -ErrorAction Stop
    $smtp=[string]$m.PrimarySmtpAddress
    $upn=[string]$m.UserPrincipalName
    [void]$gridMail.Rows.Add([string]$m.DisplayName,$upn,$smtp,[string]$m.RecipientTypeDetails)
    $script:Mailbox=$smtp
    $mailLabel.Text="Selected mailbox: $smtp"
    Log "Found mailbox: $smtp"
    Load-Delegates
}
function Load-Delegates {
    if(-not $script:Mailbox){throw 'Select a mailbox first.'}
    $gridDelegates.Rows.Clear(); $script:Rows=@(); $script:RecipientCache=@{}
    $entries=@{}
    function Add-Entry([string]$id,[string]$kind) {
        if([string]::IsNullOrWhiteSpace($id)){return}
        $resolved=Resolve-Delegate $id
        $key=if($resolved.Resolved){$resolved.UPN.ToLowerInvariant()}else{$id.ToLowerInvariant()}
        if(-not $entries.ContainsKey($key)){$entries[$key]=[pscustomobject]@{Name=$resolved.Name;UPN=$resolved.UPN;Identity=$id;Resolved=$resolved.Resolved;FullAccess=$false;SendAs=$false;SendOnBehalf=$false}}
        $entries[$key].$kind=$true
    }
    $id=$script:Mailbox
    Log "Reading delegates for $id..."
    foreach($p in @(Get-EXOMailboxPermission -Identity $id)) {
        if($p.IsInherited -or $p.Deny -or ([string]$p.User -match '^(NT AUTHORITY\\SELF|S-1-5-10)$')){continue}
        if(@($p.AccessRights) -contains 'FullAccess'){Add-Entry ([string]$p.User) 'FullAccess'}
    }
    foreach($p in @(Get-RecipientPermission -Identity $id)) {
        if($p.IsInherited -or $p.Deny -or ([string]$p.Trustee -match '^(NT AUTHORITY\\SELF|S-1-5-10)$')){continue}
        if(@($p.AccessRights) -contains 'SendAs'){Add-Entry ([string]$p.Trustee) 'SendAs'}
    }
    $mb=Get-Mailbox -Identity $id
    foreach($p in @($mb.GrantSendOnBehalfTo)) {Add-Entry ([string]$p) 'SendOnBehalf'}
    $script:Rows=@($entries.Values | Sort-Object UPN)
    foreach($e in $script:Rows){
        [void]$gridDelegates.Rows.Add($e.Name,$e.UPN,$e.FullAccess,$e.SendAs,$e.SendOnBehalf,$e.Resolved)
    }
    Log "Found $($script:Rows.Count) delegate entries. Unresolved identities are displayed, not guessed."
}
function Set-DelegatePermissions([string]$delegate,[bool]$fa,[bool]$sa,[bool]$sob,[bool]$removeAll=$false) {
    if(-not $script:Mailbox){throw 'Select a mailbox.'}
    if(-not $delegate){throw 'Enter a delegate UPN.'}
    $resolved=Resolve-Delegate $delegate
    if(-not $resolved.Resolved){throw "Could not uniquely resolve delegate '$delegate'. Enter a valid UPN."}
    $delegate=$resolved.UPN
    if($removeAll){$fa=$false;$sa=$false;$sob=$false}
    $existing=$script:Rows | Where-Object {$_.UPN -ieq $delegate} | Select-Object -First 1
    $oldFA=[bool]$existing.FullAccess; $oldSA=[bool]$existing.SendAs; $oldSOB=[bool]$existing.SendOnBehalf
    $changes=@()
    if($fa -ne $oldFA){$changes+="Full Access: $oldFA -> $fa"}
    if($sa -ne $oldSA){$changes+="Send As: $oldSA -> $sa"}
    if($sob -ne $oldSOB){$changes+="Send on Behalf: $oldSOB -> $sob"}
    if(-not $changes.Count){Log 'No permission changes needed.';return}
    $msg="Mailbox: $script:Mailbox`nDelegate: $delegate`n`n$($changes -join "`n")`n`nApply changes?"
    if([System.Windows.Forms.MessageBox]::Show($msg,'Confirm permission changes','YesNo','Question') -ne 'Yes'){return}
    $failures=@()
    if($fa -ne $oldFA){try{if($fa){Add-MailboxPermission -Identity $script:Mailbox -User $delegate -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -Confirm:$false -ErrorAction Stop | Out-Null}else{Remove-MailboxPermission -Identity $script:Mailbox -User $delegate -AccessRights FullAccess -Confirm:$false -ErrorAction Stop | Out-Null};Log 'Updated Full Access.'}catch{$failures+="Full Access: $($_.Exception.Message)"}}
    if($sa -ne $oldSA){try{if($sa){Add-RecipientPermission -Identity $script:Mailbox -Trustee $delegate -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null}else{Remove-RecipientPermission -Identity $script:Mailbox -Trustee $delegate -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null};Log 'Updated Send As.'}catch{$failures+="Send As: $($_.Exception.Message)"}}
    if($sob -ne $oldSOB){try{if($sob){Set-Mailbox -Identity $script:Mailbox -GrantSendOnBehalfTo @{Add=$delegate} -ErrorAction Stop}else{Set-Mailbox -Identity $script:Mailbox -GrantSendOnBehalfTo @{Remove=$delegate} -ErrorAction Stop};Log 'Updated Send on Behalf.'}catch{$failures+="Send on Behalf: $($_.Exception.Message)"}}
    Load-Delegates
    if($failures.Count){throw ($failures -join "`n")}
}
$form=New-Object System.Windows.Forms.Form
$form.Text='Microsoft 365 Mailbox Delegate Manager';$form.Size=New-Object System.Drawing.Size(1190,820);$form.StartPosition='CenterScreen';$form.MinimumSize=New-Object System.Drawing.Size(1050,700)
$font=New-Object System.Drawing.Font('Segoe UI',9)
$form.Font=$font
function Button($text,$x,$y,$w,$handler){$b=New-Object System.Windows.Forms.Button;$b.Text=$text;$b.SetBounds($x,$y,$w,32);$b.Add_Click($handler);$form.Controls.Add($b);return $b}
function Label($text,$x,$y,$w){$l=New-Object System.Windows.Forms.Label;$l.Text=$text;$l.SetBounds($x,$y,$w,25);$form.Controls.Add($l);return $l}
function Grid($x,$y,$w,$h){$g=New-Object System.Windows.Forms.DataGridView;$g.SetBounds($x,$y,$w,$h);$g.ReadOnly=$true;$g.AllowUserToAddRows=$false;$g.AllowUserToDeleteRows=$false;$g.SelectionMode='FullRowSelect';$g.MultiSelect=$false;$g.AutoSizeColumnsMode='Fill';$g.RowHeadersVisible=$false;$form.Controls.Add($g);return $g}
Label 'Mailbox UPN / email:' 15 17 145 | Out-Null
$search=New-Object System.Windows.Forms.TextBox;$search.SetBounds(160,13,405,28);$form.Controls.Add($search)
$null=Button 'Connect Exchange Online' 580 11 205 {Invoke-UI {Import-Module ExchangeOnlineManagement -ErrorAction Stop;Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop;Log 'Connected to Exchange Online.'}}
$null=Button 'Search UPN / Load Delegates' 800 11 265 {Invoke-UI {Find-Mailbox}}
$search.Add_KeyDown({if($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter){$_.SuppressKeyPress=$true;Invoke-UI {Find-Mailbox}}})
Label 'Enter an exact mailbox UPN or SMTP address; search retrieves only that mailbox.' 15 53 850 | Out-Null
Label 'Mailbox lookup result' 15 90 260 | Out-Null
$gridMail=Grid 15 115 1140 210
foreach($c in @('Display Name','UPN','Primary SMTP','Type')){[void]$gridMail.Columns.Add($c,$c)}
$gridMail.Add_SelectionChanged({if($gridMail.SelectedRows.Count -eq 1){$script:Mailbox=[string]$gridMail.SelectedRows[0].Cells[2].Value;$gridDelegates.Rows.Clear();$script:Rows=@();$mailLabel.Text="Selected mailbox: $script:Mailbox"}})
$mailLabel=Label 'Selected mailbox: None' 15 335 750
$null=Button 'Load / Refresh Delegates' 880 331 275 {Invoke-UI {Load-Delegates}}
$gridDelegates=Grid 15 375 1140 210
foreach($c in @('Display Name','Delegate UPN / Identity','Full Access','Send As','Send on Behalf','Resolved')){[void]$gridDelegates.Columns.Add($c,$c)}
$gridDelegates.Add_SelectionChanged({
    if($gridDelegates.SelectedRows.Count -eq 1){
        $r=$gridDelegates.SelectedRows[0];$delegateBox.Text=[string]$r.Cells[1].Value
        $checkFA.Checked=[bool]$r.Cells[2].Value;$checkSA.Checked=[bool]$r.Cells[3].Value;$checkSOB.Checked=[bool]$r.Cells[4].Value
    }
})
Label 'Delegate UPN:' 15 600 110 | Out-Null
$delegateBox=New-Object System.Windows.Forms.TextBox;$delegateBox.SetBounds(125,597,350,28);$form.Controls.Add($delegateBox)
$checkFA=New-Object System.Windows.Forms.CheckBox;$checkFA.Text='Full Access';$checkFA.SetBounds(490,598,115,27);$form.Controls.Add($checkFA)
$checkSA=New-Object System.Windows.Forms.CheckBox;$checkSA.Text='Send As';$checkSA.SetBounds(610,598,95,27);$form.Controls.Add($checkSA)
$checkSOB=New-Object System.Windows.Forms.CheckBox;$checkSOB.Text='Send on Behalf';$checkSOB.SetBounds(715,598,155,27);$form.Controls.Add($checkSOB)
$null=Button 'Add / Update Permissions' 15 635 225 {Invoke-UI {Set-DelegatePermissions $delegateBox.Text $checkFA.Checked $checkSA.Checked $checkSOB.Checked}}
$null=Button 'Remove All 3 Permissions' 250 635 225 {Invoke-UI {Set-DelegatePermissions $delegateBox.Text $false $false $false $true}}
$null=Button 'Export Delegates CSV' 485 635 195 {Invoke-UI {if(-not $script:Mailbox){throw 'Select a mailbox.'};$save=New-Object System.Windows.Forms.SaveFileDialog;$save.Filter='CSV files (*.csv)|*.csv';$save.FileName="Delegates_$($script:Mailbox.Replace('@','_'))_$(Get-Date -Format yyyyMMdd_HHmm).csv";if($save.ShowDialog() -eq 'OK'){$script:Rows | Select-Object @{n='Mailbox';e={$script:Mailbox}},Name,UPN,FullAccess,SendAs,SendOnBehalf,Resolved | Export-Csv -Path $save.FileName -NoTypeInformation -Encoding utf8;Log "Exported $($save.FileName)"}}}
$log=New-Object System.Windows.Forms.TextBox;$log.SetBounds(15,682,1140,65);$log.Multiline=$true;$log.ScrollBars='Vertical';$log.ReadOnly=$true;$form.Controls.Add($log)
$status=New-Object System.Windows.Forms.Label;$status.Text='Ready';$status.SetBounds(900,641,250,25);$form.Controls.Add($status)
Log 'Ready. Connect to Exchange Online, then enter a mailbox UPN and search. No tenant-wide scan.'
[void]$form.ShowDialog()
