#requires -Version 7.4
<#
Windows / PowerShell 7. Run: pwsh -STA -File .\M365-Mailbox-Cleanup-GUI.ps1
Requires Exchange Online RBAC access to the target mailboxes and these cmdlets.
CSV: UserPrincipalName (one account per row). No domain is hardcoded.
Calendar cleanup uses Remove-CalendarEvents: organized meetings with attendees,
future dates only, at most 1825 days. An intersecting recurring series is cancelled
in its entirety. The mailbox must still be able to send cancellation messages.
Rule JSON exports are reference snapshots, NOT automatically restorable backups.
Docs: https://learn.microsoft.com/powershell/module/exchangepowershell/remove-calendarevents
      https://learn.microsoft.com/powershell/module/exchangepowershell/remove-inboxrule
#>
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'This WPF GUI requires Windows and PowerShell 7.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Launch with: pwsh -STA -File .\M365-Mailbox-Cleanup-GUI.ps1'
}
Add-Type -AssemblyName @('PresentationFramework','PresentationCore','WindowsBase')
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="M365 Mailbox Rules and Organized Meeting Cleanup" Width="1150" Height="950" MinWidth="940" MinHeight="720" WindowStartupLocation="CenterScreen">
 <DockPanel Margin="14">
  <TextBox x:Name="Log" DockPanel.Dock="Bottom" Height="150" IsReadOnly="True" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" Margin="0,10,0,0"/>
  <ProgressBar x:Name="Progress" DockPanel.Dock="Bottom" Height="8" Margin="0,8,0,0"/>
  <TextBlock x:Name="Status" DockPanel.Dock="Bottom" Text="Connect to Exchange Online to begin." Margin="0,8,0,0"/>
  <StackPanel x:Name="Controls" DockPanel.Dock="Top">
   <WrapPanel Margin="0,0,0,10">
    <Button x:Name="Connect" Content="Connect / Reconnect Exchange Online" Padding="12,6"/>
    <Button x:Name="ExportLog" Content="Save Log" Padding="12,6" Margin="10,0,0,0"/>
   </WrapPanel>
   <GroupBox Header="1. Target accounts" Padding="8">
    <StackPanel>
     <WrapPanel><TextBlock Text="UPN:" VerticalAlignment="Center"/><TextBox x:Name="UPN" Width="310" Margin="8,0"/><Button x:Name="Add" Content="Add account" Padding="10,4"/><Button x:Name="Import" Content="Import CSV" Padding="10,4" Margin="8,0"/><Button x:Name="Clear" Content="Clear targets" Padding="10,4"/></WrapPanel>
     <TextBlock Text="CSV header: UserPrincipalName. Select target rows below; all actions use ONLY the selected rows." Margin="0,8,0,4"/>
     <ListBox x:Name="Targets" Height="86" SelectionMode="Extended"/>
    </StackPanel>
   </GroupBox>
   <GroupBox Header="2. Inbox rules" Padding="8" Margin="0,8,0,0">
    <WrapPanel>
     <CheckBox x:Name="Hidden" Content="Include hidden rules" VerticalAlignment="Center" Margin="0,0,12,0"/>
     <Button x:Name="ReadRules" Content="Load rules" Padding="10,5"/>
     <Button x:Name="DeleteSelected" Content="Delete selected rules" Padding="10,5" Margin="8,0"/>
     <Button x:Name="DeleteAll" Content="Delete ALL rules for selected accounts" Padding="10,5"/>
     <Button x:Name="ExportRules" Content="Export rules JSON" Padding="10,5" Margin="8,0,0,0"/>
    </WrapPanel>
   </GroupBox>
   <GroupBox Header="3. Meetings organized by selected accounts" Padding="8" Margin="0,8,0,8">
    <StackPanel>
     <WrapPanel><TextBlock Text="Start:" VerticalAlignment="Center"/><DatePicker x:Name="StartDate" Width="145" Margin="8,0"/><TextBlock Text="Days (1-1825):" VerticalAlignment="Center"/><TextBox x:Name="Days" Text="365" Width="65" Margin="8,0"/><Button x:Name="PreviewCalendar" Content="Preview meetings (no changes)" Padding="10,5"/><Button x:Name="CancelCalendar" Content="Cancel organized meetings" Padding="10,5" Margin="8,0"/></WrapPanel>
     <TextBlock Text="Cancellations are sent to attendees. Personal appointments are excluded. Recurring series are cancelled in full if any occurrence is in range. Preview details appear in the log." TextWrapping="Wrap" Margin="0,8,0,0"/>
    </StackPanel>
   </GroupBox>
  </StackPanel>
  <TabControl x:Name="Views"><TabItem Header="Inbox rules">
  <DataGrid x:Name="Rules" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended" SelectionUnit="FullRow" CanUserAddRows="False" HorizontalScrollBarVisibility="Auto">
   <DataGrid.Columns>
    <DataGridTextColumn Header="Mailbox" Binding="{Binding Mailbox}" Width="200"/>
    <DataGridTextColumn Header="Rule name" Binding="{Binding Name}" Width="180"/>
    <DataGridTextColumn Header="Enabled" Binding="{Binding Enabled}" Width="65"/>
    <DataGridTextColumn Header="Priority" Binding="{Binding Priority}" Width="65"/>
    <DataGridTextColumn Header="Description / conditions / actions" Binding="{Binding Description}" Width="500"/>
   </DataGrid.Columns>
  </DataGrid>
  </TabItem><TabItem Header="Mailbox storage" x:Name="StorageTab">
   <DockPanel Margin="8">
    <StackPanel DockPanel.Dock="Top">
     <WrapPanel><Button x:Name="ReadStorage" Content="Load storage for selected accounts" Padding="12,6"/><Button x:Name="ExportStorage" Content="Export folder CSV(s)" Padding="12,6" Margin="8,0"/></WrapPanel>
     <TextBlock Text="Primary mailbox only. Archive folder is included; separate Online Archive mailbox is not. Select a mailbox summary to view its folders." TextWrapping="Wrap" Margin="0,8"/>
     <DataGrid x:Name="StorageSummary" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" SelectionMode="Single" Height="100">
      <DataGrid.Columns>
       <DataGridTextColumn Header="Mailbox" Binding="{Binding Mailbox}" Width="230"/>
       <DataGridTextColumn Header="Used" Binding="{Binding Used}" Width="90"/>
       <DataGridTextColumn Header="Quota" Binding="{Binding Quota}" Width="90"/>
       <DataGridTextColumn Header="Used %" Binding="{Binding UsedPercent}" Width="75"/>
       <DataGridTextColumn Header="Items" Binding="{Binding Items}" Width="90"/>
       <DataGridTextColumn Header="Recoverable items (separate)" Binding="{Binding Recoverable}" Width="180"/>
      </DataGrid.Columns>
     </DataGrid>
     <TextBlock x:Name="StorageLabel" Text="Load storage to begin." Margin="0,8,0,4"/>
     <ProgressBar x:Name="StorageUsage" Height="14" Maximum="100"/>
     <TextBlock Text="Folder size/count excludes subfolders; inclusive columns include them. Do not add inclusive sizes together. Hidden/system items may be included; totals can differ from Outlook. Sizes use 1024-based units." TextWrapping="Wrap" Margin="0,8"/>
    </StackPanel>
    <DataGrid x:Name="StorageFolders" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False">
     <DataGrid.Columns>
      <DataGridTextColumn Header="Folder path" Binding="{Binding FolderPath}" Width="260"/>
      <DataGridTextColumn Header="Type" Binding="{Binding FolderType}" Width="110"/>
      <DataGridTextColumn Header="Size (MB)" Binding="{Binding SizeMB}" Width="100"/>
      <DataGridTextColumn Header="Items" Binding="{Binding Items}" Width="80"/>
      <DataGridTextColumn Header="With subfolders (MB)" Binding="{Binding InclusiveMB}" Width="150"/>
      <DataGridTextColumn Header="Items with subfolders" Binding="{Binding InclusiveItems}" Width="145"/>
     </DataGrid.Columns>
    </DataGrid>
   </DockPanel>
  </TabItem><TabItem Header="Delegation">
   <DockPanel Margin="8">
    <StackPanel DockPanel.Dock="Top">
     <WrapPanel><Button x:Name="ReadDelegation" Content="Load delegates for selected accounts" Padding="12,6"/><Button x:Name="ExportDelegation" Content="Export delegation CSV" Padding="12,6" Margin="8,0"/></WrapPanel>
     <TextBlock Text="Read-only: who has Full Access, Send As, or Send on Behalf on the selected mailboxes. Groups are listed without expanding membership. SELF/system entries and inherited/deny entries are shown as returned. Unresolved trustees retain their original identity. This does not scan other mailboxes for reverse access." TextWrapping="Wrap" Margin="0,8"/>
    </StackPanel>
    <DataGrid x:Name="Delegates" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False">
     <DataGrid.Columns>
      <DataGridTextColumn Header="Mailbox UPN" Binding="{Binding UserPrincipalName}" Width="210"/>
      <DataGridTextColumn Header="Permission" Binding="{Binding Permission}" Width="130"/>
      <DataGridTextColumn Header="Delegate UPN / email" Binding="{Binding DelegateUPN}" Width="210"/>
      <DataGridTextColumn Header="Delegate name" Binding="{Binding DelegateName}" Width="160"/>
      <DataGridTextColumn Header="Trustee identity" Binding="{Binding TrusteeIdentity}" Width="220"/>
      <DataGridTextColumn Header="Access type" Binding="{Binding AccessType}" Width="85"/>
      <DataGridTextColumn Header="Inherited" Binding="{Binding IsInherited}" Width="75"/>
      <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="280"/>
     </DataGrid.Columns>
    </DataGrid>
   </DockPanel>
  </TabItem></TabControl>
 </DockPanel>
</Window>
'@
$window = [Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
$ui = @{}
foreach ($n in 'Log','Progress','Status','Controls','Connect','ExportLog','UPN','Add','Import','Clear','Targets','Hidden','ReadRules','DeleteSelected','DeleteAll','ExportRules','StartDate','Days','PreviewCalendar','CancelCalendar','Rules','Views','StorageTab','ReadStorage','ExportStorage','StorageSummary','StorageLabel','StorageUsage','StorageFolders','ReadDelegation','ExportDelegation','Delegates') { $ui[$n] = $window.FindName($n) }
$ui.StartDate.SelectedDate = [datetime]::Today
$script:connected = $false
$script:busy = $false
$script:ruleRows = @()
$script:storageRows = @()
$script:folderRows = @()
$script:delegateRows = @()
$script:worker = $null
$script:pending = $null
$script:queue = [Collections.Concurrent.ConcurrentQueue[object]]::new()
$script:runspace = [RunspaceFactory]::CreateRunspace()
$script:runspace.ApartmentState = 'STA'
$script:runspace.ThreadOptions = 'ReuseThread'
$script:runspace.Open()
function Write-Log([string]$Message) {
    $ui.Log.AppendText(('[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ' + $Message + "`r`n"))
    $ui.Log.ScrollToEnd()
}
function Show-Error($ErrorObject) { Write-Log "ERROR: $ErrorObject"; [void][Windows.MessageBox]::Show($window,[string]$ErrorObject,'Mailbox cleanup error','OK','Error') }
function Get-Targets {
    $items = @($ui.Targets.SelectedItems | ForEach-Object { [string]$_ })
    if ($items.Count -eq 0) { throw 'Select at least one target account row.' }
    return ,$items
}
function Add-Target([string]$Value) {
    $v = $Value.Trim().ToLowerInvariant()
    if ($v -notmatch '^[^\s@,;*?]+@[^\s@,;*?]+\.[^\s@,;*?]+$') { throw "Invalid UPN: $Value" }
    if (-not $ui.Targets.Items.Contains($v)) { [void]$ui.Targets.Items.Add($v) }
    $ui.Targets.SelectedItems.Add($v)
}
function Confirm-Action([string]$Text) {
    return [Windows.MessageBox]::Show($window,$Text,'Confirm mailbox changes','YesNo','Warning','No') -eq 'Yes'
}
function Get-SavePath([string]$Filter,[string]$Name) {
    $d = [Microsoft.Win32.SaveFileDialog]::new(); $d.Filter = $Filter; $d.FileName = $Name
    if ($d.ShowDialog($window)) { return $d.FileName }
}
# All Exchange commands run in one persistent background runspace, keeping the
# Exchange connection intact and the WPF UI responsive without ThreadJob.
$script:taskCode = {
    param($Request,$Queue)
    $ErrorActionPreference = 'Stop'
    function Report([string]$Message) { $Queue.Enqueue([pscustomobject]@{Kind='Log';Message=$Message}) }
    $recipientCache = @{}
    function Make-DelegateRow($MailboxUPN,$Permission,$Trustee,$AccessType,$Inherited,$Status) {
        $key = [string]$Trustee
        if ($key -and -not $recipientCache.ContainsKey($key)) {
            try {
                $recipients = @(Get-Recipient -Identity $key -ErrorAction Stop)
                if ($recipients.Count -ne 1) { throw 'Trustee did not resolve uniquely.' }
                $r = $recipients[0]
                $address = if ($r.UserPrincipalName) { [string]$r.UserPrincipalName } else { [string]$r.PrimarySmtpAddress }
                $recipientCache[$key] = @{Name=[string]$r.DisplayName;Address=$address;Resolved=$true}
            } catch { $recipientCache[$key] = @{Name=$key;Address='';Resolved=$false} }
        }
        $resolved = if ($key) { $recipientCache[$key] } else { @{Name='';Address='';Resolved=$true} }
        if ($Status -eq 'OK' -and -not $resolved.Resolved) { $Status = 'Trustee unresolved; original identity retained' }
        [pscustomobject]@{UserPrincipalName=$MailboxUPN;Permission=$Permission;DelegateUPN=$resolved.Address
            DelegateName=$resolved.Name;TrusteeIdentity=$key;AccessType=$AccessType;IsInherited=$Inherited;Status=$Status}
    }
    function Convert-SizeBytes($Value) {
        if ($null -eq $Value) { return $null }
        if ($Value.PSObject.Methods['ToBytes']) { return [double]$Value.ToBytes() }
        if ($Value.PSObject.Properties['Value'] -and $Value.Value.PSObject.Methods['ToBytes']) { return [double]$Value.Value.ToBytes() }
        $text = [string]$Value
        if ($text -match '\(([\d\s,.]+)\s+bytes\)') { return [double]($Matches[1] -replace '[^0-9]','') }
        if ($text -match '^([0-9]+(?:\.[0-9]+)?)\s*(B|KB|MB|GB|TB)$') {
            $factor = switch ($Matches[2]) { B {1} KB {1KB} MB {1MB} GB {1GB} TB {1TB} }
            return [double]::Parse($Matches[1],[Globalization.CultureInfo]::InvariantCulture) * $factor
        }
        if ($text -eq 'Unlimited') { return $null }
        throw "Cannot parse Exchange size: $text"
    }
    function Format-Storage($Bytes) {
        if ($null -eq $Bytes) { return 'Unknown / unlimited' }
        if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
        return ('{0:N2} MB' -f ($Bytes / 1MB))
    }
    try {
        if ($Request.Action -eq 'Connect') {
            if (-not (Get-Module -ListAvailable ExchangeOnlineManagement | Where-Object Version -ge ([version]'3.7.2'))) {
                Report 'Installing ExchangeOnlineManagement for the current user from PSGallery...'
                Install-Module ExchangeOnlineManagement -MinimumVersion 3.7.2 -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
            }
            Import-Module ExchangeOnlineManagement -MinimumVersion 3.7.2 -ErrorAction Stop
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Connect-ExchangeOnline -ShowBanner:$false -DisableWAM -ErrorAction Stop
            $connections = @(Get-ConnectionInformation | Where-Object State -eq 'Connected')
            if ($connections.Count -ne 1) { throw 'Expected one active Exchange Online connection. Reconnect in a fresh PowerShell session.' }
            Report "Connected as $($connections[0].UserPrincipalName); tenant: $($connections[0].TenantID)"
            $Queue.Enqueue([pscustomobject]@{Kind='Connected'})
            return
        }
        $allRows = [Collections.Generic.List[object]]::new()
        $index = 0
        foreach ($target in $Request.Targets) {
            $index++
            Report "[$index/$($Request.Targets.Count)] $target - $($Request.Action)"
            try {
                # Resolve every mailbox before invoking destructive commands; use
                # its GUID thereafter to avoid ambiguous aliases or display names.
                $mailboxes = @(Get-EXOMailbox -Identity $target -ErrorAction Stop)
                if ($mailboxes.Count -ne 1) { throw 'Mailbox did not resolve uniquely.' }
                $mailbox = $mailboxes[0]
                $id = [string]$mailbox.Guid
                if ([string]::IsNullOrWhiteSpace($id)) { throw 'Resolved mailbox has no GUID.' }
                Report "Resolved: $($mailbox.DisplayName) <$($mailbox.PrimarySmtpAddress)> [$id]"
                $accountUPN = if ($mailbox.UserPrincipalName) { [string]$mailbox.UserPrincipalName } else { $target }
                switch ($Request.Action) {
                    'ReadDelegation' {
                        foreach ($permission in 'Full Access','Send As','Send on Behalf') {
                            try {
                                $entries = @(switch ($permission) {
                                    'Full Access' { Get-EXOMailboxPermission -Identity $target -ResultSize Unlimited -ErrorAction Stop | Where-Object { 'FullAccess' -in $_.AccessRights } }
                                    'Send As' { Get-EXORecipientPermission -Identity $target -ResultSize Unlimited -ErrorAction Stop | Where-Object { 'SendAs' -in $_.AccessRights } }
                                    'Send on Behalf' { (Get-Mailbox -Identity $id -ErrorAction Stop).GrantSendOnBehalfTo }
                                })
                                if (-not $entries.Count) { $allRows.Add((Make-DelegateRow $accountUPN $permission '' '' $null 'No assignments returned')) }
                                foreach ($entry in $entries) {
                                    $trustee = switch ($permission) { 'Full Access' { [string]$entry.User } 'Send As' { [string]$entry.Trustee } 'Send on Behalf' { [string]$entry } }
                                    $access = if ($permission -eq 'Full Access') { if ($entry.Deny) { 'Deny' } else { 'Allow' } } elseif ($permission -eq 'Send As') { [string]$entry.AccessControlType } else { 'Allow' }
                                    $inherited = if ($permission -eq 'Send on Behalf') { $false } else { $entry.IsInherited }
                                    $allRows.Add((Make-DelegateRow $accountUPN $permission $trustee $access $inherited 'OK'))
                                }
                                Report "$permission : $($entries.Count) assignments returned."
                            } catch {
                                $allRows.Add((Make-DelegateRow $accountUPN $permission '' '' $null ("FAILED: " + $_.Exception.Message)))
                                Report "FAILED $permission : $($_.Exception.Message)"
                            }
                        }
                    }
                    'ReadStorage' {
                        $details = Get-EXOMailbox -Identity $target -Properties ProhibitSendReceiveQuota -ErrorAction Stop
                        $stats = Get-EXOMailboxStatistics -Identity $target -Properties TotalItemSize,TotalDeletedItemSize,ItemCount -ErrorAction Stop
                        $used = Convert-SizeBytes $stats.TotalItemSize
                        $quota = Convert-SizeBytes $details.ProhibitSendReceiveQuota
                        $recoverable = Convert-SizeBytes $stats.TotalDeletedItemSize
                        $percent = if ($null -ne $used -and $quota -gt 0) { [math]::Round(100 * $used / $quota,2) } else { $null }
                        $Queue.Enqueue([pscustomobject]@{Kind='StorageSummary';Row=[pscustomobject]@{
                            UserPrincipalName=$accountUPN;Mailbox=$target;Used=(Format-Storage $used);Quota=(Format-Storage $quota)
                            UsedPercent=$percent;Items=$stats.ItemCount;Recoverable=(Format-Storage $recoverable)
                            UsedBytes=$used;QuotaBytes=$quota;RecoverableBytes=$recoverable;CapturedAt=(Get-Date -Format o)
                        }})
                        $folders = @(Get-EXOMailboxFolderStatistics -Identity $target -ErrorAction Stop)
                        $rows = @(foreach ($folder in $folders) {
                            $bytes = Convert-SizeBytes $folder.FolderSize
                            $inclusive = Convert-SizeBytes $folder.FolderAndSubfolderSize
                            [pscustomobject]@{UserPrincipalName=$accountUPN;Mailbox=$target;FolderPath=[string]$folder.FolderPath;FolderType=[string]$folder.FolderType
                                SizeBytes=$bytes;SizeMB=([math]::Round($bytes / 1MB,2));Items=$folder.ItemsInFolder
                                InclusiveMB=([math]::Round($inclusive / 1MB,2));InclusiveItems=$folder.ItemsInFolderAndSubfolders}
                        })
                        $Queue.Enqueue([pscustomobject]@{Kind='StorageFolders';Rows=@($rows | Sort-Object SizeBytes -Descending)})
                        Report "Storage loaded: $(Format-Storage $used); $($folders.Count) folders."
                    }
                    'ReadRules' {
                        $found = @(Get-InboxRule -Mailbox $id -IncludeHidden:($Request.Hidden) -ErrorAction Stop)
                        foreach ($rule in $found) {
                            $allRows.Add([pscustomobject]@{Mailbox=$target;MailboxGuid=$id;Name=[string]$rule.Name;Enabled=$rule.Enabled;Priority=$rule.Priority;Description=[string]$rule.Description;RuleIdentity=[string]$rule.Identity;Raw=$rule})
                        }
                        Report "$($found.Count) rules loaded."
                    }
                    { $_ -in 'DeleteSelected','DeleteAll' } {
                        if ($Request.Action -eq 'DeleteAll') {
                            $toDelete = @(Get-InboxRule -Mailbox $id -IncludeHidden:($Request.Hidden) -ErrorAction Stop)
                        } else {
                            $toDelete = @($Request.Rows | Where-Object Mailbox -eq $target)
                            if (@($toDelete | Where-Object MailboxGuid -ne $id).Count) { throw 'Mailbox identity changed since preview. Reload rules.' }
                        }
                        Report "Attempting deletion of $($toDelete.Count) rules."
                        foreach ($rule in $toDelete) {
                            try {
                                $ruleID = if ($Request.Action -eq 'DeleteAll') { [string]$rule.Identity } else { [string]$rule.RuleIdentity }
                                if ([string]::IsNullOrWhiteSpace($ruleID)) { throw 'Missing rule identity; deletion skipped.' }
                                # A qualified Identity already contains the mailbox.
                                # Only supply Mailbox when Identity is rule-only.
                                $removeParams = @{Identity=$ruleID;Confirm=$false;Force=$true;ErrorAction='Stop'}
                                if (-not $ruleID.Contains('\')) { $removeParams.Mailbox = $id }
                                Remove-InboxRule @removeParams
                                Report "DELETED rule: $($rule.Name) [$ruleID]"
                            } catch { Report "FAILED rule '$($rule.Name)': $($_.Exception.Message)" }
                        }
                    }
                    { $_ -in 'PreviewCalendar','CancelCalendar' } {
                        $params = @{Identity=$id;CancelOrganizedMeetings=$true;QueryStartDate=$Request.Start;QueryWindowInDays=$Request.Days;Confirm=$false;ErrorAction='Stop';Verbose=$true}
                        if ($Request.Action -eq 'PreviewCalendar') { $params.PreviewOnly = $true }
                        Remove-CalendarEvents @params 4>&1 | ForEach-Object { Report ($_.ToString()) }
                        if ($Request.Action -eq 'PreviewCalendar') { Report 'PREVIEW ONLY completed. Any queued-for-cancellation wording above is preview output; no cancellations were submitted.' }
                        else { Report 'Calendar cancellation request completed. Exchange may continue processing; verify results in Outlook.' }
                    }
                }
            } catch {
                Report "FAILED $target : $($_.Exception.Message)"
                if ($Request.Action -eq 'ReadDelegation') { $allRows.Add((Make-DelegateRow $target 'Mailbox lookup' '' '' $null ("FAILED: " + $_.Exception.Message))) }
            }
        }
        if ($Request.Action -eq 'ReadDelegation') { $Queue.Enqueue([pscustomobject]@{Kind='Delegates';Rows=$allRows.ToArray()}) }
        if ($Request.Action -eq 'ReadRules') { $Queue.Enqueue([pscustomobject]@{Kind='Rules';Rows=$allRows.ToArray()}) }
    } catch { $Queue.Enqueue([pscustomobject]@{Kind='Error';Message=$_.Exception.Message}) }
}
function Start-Task($Request) {
    if ($script:busy) { return }
    if ($Request.Action -ne 'Connect' -and -not $script:connected) { throw 'Connect to Exchange Online first.' }
    if ($Request.Action -eq 'Connect') { $script:connected = $false }
    $script:busy = $true
    $ui.ReadDelegation.IsEnabled = $false; $ui.ExportDelegation.IsEnabled = $false
    $ui.ReadStorage.IsEnabled = $false; $ui.ExportStorage.IsEnabled = $false
    $ui.Controls.IsEnabled = $false; $ui.Rules.IsEnabled = $false
    $ui.Progress.IsIndeterminate = $true; $ui.Status.Text = "Working: $($Request.Action)"
    $script:lastAction = $Request.Action
    $script:worker = [PowerShell]::Create()
    $script:worker.Runspace = $script:runspace
    [void]$script:worker.AddScript($script:taskCode.ToString()).AddArgument($Request).AddArgument($script:queue)
    $script:pending = $script:worker.BeginInvoke()
}
$timer = [Windows.Threading.DispatcherTimer]::new()
$timer.Interval = [timespan]::FromMilliseconds(250)
$timer.Add_Tick({
    $message = $null
    while ($script:queue.TryDequeue([ref]$message)) {
        switch ($message.Kind) {
            'Log' { Write-Log $message.Message }
            'Error' { Write-Log "ERROR: $($message.Message)" }
            'Connected' { $script:connected = $true }
            'StorageSummary' {
                $script:storageRows += $message.Row
                $ui.StorageSummary.ItemsSource = @($script:storageRows)
                if ($ui.StorageSummary.SelectedIndex -lt 0) { $ui.StorageSummary.SelectedIndex = 0 }
            }
            'StorageFolders' { $script:folderRows += @($message.Rows); Update-StorageView }
            'Delegates' { $script:delegateRows = @($message.Rows); $ui.Delegates.ItemsSource = $script:delegateRows }
            'Rules' { $script:ruleRows = @($message.Rows); $ui.Rules.ItemsSource = $script:ruleRows }
        }
    }
    if ($script:busy -and $script:pending.IsCompleted) {
        try { [void]$script:worker.EndInvoke($script:pending) } catch { Write-Log "ERROR: $_" }
        foreach ($e in $script:worker.Streams.Error) { Write-Log "ERROR: $e" }
        foreach ($w in $script:worker.Streams.Warning) { Write-Log "WARNING: $w" }
        $script:worker.Dispose(); $script:worker = $null
        if ($script:lastAction -in 'DeleteSelected','DeleteAll') { $script:ruleRows = @(); $ui.Rules.ItemsSource = $null; Write-Log 'Reload rules to verify the remaining rules.' }
        $ui.ReadDelegation.IsEnabled = $true; $ui.ExportDelegation.IsEnabled = $true
        $ui.ReadStorage.IsEnabled = $true; $ui.ExportStorage.IsEnabled = $true
        $script:busy = $false; $ui.Controls.IsEnabled = $true; $ui.Rules.IsEnabled = $true
        $ui.Progress.IsIndeterminate = $false
        $ui.Status.Text = 'Finished. Review the log for individual successes and failures.'
    }
})
$ui.Connect.Add_Click({ try { Start-Task @{Action='Connect'} } catch { Show-Error $_ } })
$ui.Add.Add_Click({ try { Add-Target $ui.UPN.Text; $ui.UPN.Clear() } catch { Show-Error $_ } })
$ui.Import.Add_Click({
    try {
        $d = [Microsoft.Win32.OpenFileDialog]::new(); $d.Filter = 'CSV files (*.csv)|*.csv'
        if (-not $d.ShowDialog($window)) { return }
        $rows = @(Import-Csv -LiteralPath $d.FileName)
        if (-not $rows.Count -or 'UserPrincipalName' -notin $rows[0].PSObject.Properties.Name) { throw 'CSV must contain a UserPrincipalName column and at least one account.' }
        # Validate every row before adding any of them.
        foreach ($row in $rows) { if ([string]$row.UserPrincipalName -notmatch '^[^\s@,;*?]+@[^\s@,;*?]+\.[^\s@,;*?]+$') { throw "Invalid CSV UPN: $($row.UserPrincipalName)" } }
        foreach ($row in $rows) { Add-Target $row.UserPrincipalName }
        Write-Log "Imported $($rows.Count) CSV rows; targets deduplicated and imported accounts selected."
    } catch { Show-Error $_ }
})
$ui.Clear.Add_Click({ $ui.Targets.Items.Clear(); $script:ruleRows = @(); $ui.Rules.ItemsSource = $null })
$ui.ReadRules.Add_Click({ try { Start-Task @{Action='ReadRules';Targets=(Get-Targets);Hidden=[bool]$ui.Hidden.IsChecked} } catch { Show-Error $_ } })
function Request-RuleDeletion([string]$Action) {
    try {
        $targets = Get-Targets
        $rows = @()
        if ($Action -eq 'DeleteSelected') {
            $rows = @($ui.Rules.SelectedItems)
            if (-not $rows.Count) { throw 'Load rules and select the rule rows to delete.' }
            if (@($rows | Where-Object { $_.Mailbox -notin $targets }).Count) { throw 'Selected rule rows include accounts outside the selected target rows. Adjust your selections.' }
            $targets = @($rows.Mailbox | Sort-Object -Unique)
            $scope = "$($rows.Count) selected rule(s)"
        } else { $scope = 'ALL rules (including disabled rules)'; if ($ui.Hidden.IsChecked) { $scope += ', including hidden rules' } else { $scope += ', excluding hidden rules' } }
        $names = $targets -join "`n"
        $text = "Permanently delete $scope for:`n$names`n`nMicrosoft warns that changing inbox rules through Exchange PowerShell also removes Outlook client-only rules. A JSON export is a reference snapshot, not a restorable backup.`n`nProceed?"
        if (Confirm-Action $text) { Start-Task @{Action=$Action;Targets=$targets;Rows=$rows;Hidden=[bool]$ui.Hidden.IsChecked} }
    } catch { Show-Error $_ }
}
$ui.DeleteSelected.Add_Click({ Request-RuleDeletion 'DeleteSelected' })
$ui.DeleteAll.Add_Click({ Request-RuleDeletion 'DeleteAll' })
function Request-Calendar([string]$Action) {
    try {
        $targets = Get-Targets
        $days = 0
        if (-not [int]::TryParse($ui.Days.Text,[ref]$days) -or $days -lt 1 -or $days -gt 1825) { throw 'Days must be a whole number from 1 to 1825.' }
        if ($null -eq $ui.StartDate.SelectedDate -or $ui.StartDate.SelectedDate.Date -lt [datetime]::Today) { throw 'Select today or a future start date. This tool cancels future meetings.' }
        $start = $ui.StartDate.SelectedDate.Date
        if ($Action -eq 'CancelCalendar') {
            $text = "Cancel organized meetings for:`n$($targets -join "`n")`n`nStart: $($start.ToString('yyyy-MM-dd')); window: $days days.`n`nAttendees and resources will receive cancellations. If any recurring occurrence falls in this range, the ENTIRE series is cancelled. Personal appointments are excluded. Mailboxes must be enabled to send mail.`n`nRun Preview first to review the server output. Proceed with cancellations?"
            if (-not (Confirm-Action $text)) { return }
        }
        Write-Log "$Action : start $($start.ToString('yyyy-MM-dd')), window $days days."
        Start-Task @{Action=$Action;Targets=$targets;Start=$start;Days=$days}
    } catch { Show-Error $_ }
}
$ui.PreviewCalendar.Add_Click({ Request-Calendar 'PreviewCalendar' })
$ui.CancelCalendar.Add_Click({ Request-Calendar 'CancelCalendar' })
$ui.ExportRules.Add_Click({
    try {
        if (-not $script:ruleRows.Count) { throw 'Load rules before exporting.' }
        $path = Get-SavePath 'JSON files (*.json)|*.json' ('MailboxRules-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.json')
        if ($path) { ConvertTo-Json -InputObject @($script:ruleRows) -Depth 12 | Set-Content -LiteralPath $path -Encoding utf8; Write-Log "Rule reference snapshot exported: $path" }
    } catch { Show-Error $_ }
})
$ui.ExportLog.Add_Click({
    try { $path = Get-SavePath 'Text files (*.txt)|*.txt' ('MailboxCleanup-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.txt'); if ($path) { $ui.Log.Text | Set-Content -LiteralPath $path -Encoding utf8 } } catch { Show-Error $_ }
})
function Update-StorageView {
    $row = $ui.StorageSummary.SelectedItem
    if ($null -eq $row) { $ui.StorageFolders.ItemsSource = $null; return }
    $ui.StorageFolders.ItemsSource = @($script:folderRows | Where-Object Mailbox -eq $row.Mailbox | Sort-Object SizeBytes -Descending)
    $suffix = if ($null -ne $row.UsedPercent) { " ($($row.UsedPercent)%)" } else { '' }
    $ui.StorageLabel.Text = "$($row.Mailbox): $($row.Used) used / $($row.Quota)$suffix"
    $ui.StorageUsage.Value = if ($null -ne $row.UsedPercent) { [math]::Min(100,$row.UsedPercent) } else { 0 }
}
$ui.StorageSummary.Add_SelectionChanged({ Update-StorageView })
$ui.ReadStorage.Add_Click({
    try {
        $targets = Get-Targets
        if ($script:busy) { return }
        if (-not $script:connected) { throw 'Connect to Exchange Online first.' }
        $script:storageRows = @(); $script:folderRows = @()
        $ui.StorageSummary.ItemsSource = $null; $ui.StorageFolders.ItemsSource = $null
        $ui.StorageUsage.Value = 0; $ui.StorageLabel.Text = 'Loading storage...'
        Start-Task @{Action='ReadStorage';Targets=$targets}
    } catch { Show-Error $_ }
})
$ui.ExportStorage.Add_Click({
    try {
        if (-not $script:folderRows.Count) { throw 'Load folder storage before exporting. No folder data is available.' }
        $groups = @($script:folderRows | Group-Object UserPrincipalName)
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
        $directory = $null
        if ($groups.Count -gt 1) {
            $dialog = [Microsoft.Win32.OpenFolderDialog]::new()
            $dialog.Title = 'Choose a folder for the per-account folder CSV files'
            if (-not $dialog.ShowDialog($window)) { return }
            $directory = $dialog.FolderName
        }
        foreach ($group in $groups) {
            if ([string]::IsNullOrWhiteSpace($group.Name)) { throw 'Folder data is missing the account UPN. Reload storage before exporting.' }
            # Keep the UPN, including @, in the filename; sanitize only characters
            # Windows does not permit in filenames.
            $safeUPN = $group.Name -replace '[<>:"/\\|?*\x00-\x1F]', '_'
            $name = "MailboxFolders-$safeUPN-$stamp.csv"
            if ($groups.Count -eq 1) {
                $path = Get-SavePath 'CSV files (*.csv)|*.csv' $name
                if (-not $path) { return }
            } else {
                $path = [IO.Path]::Combine($directory,$name)
                if (Test-Path -LiteralPath $path) { throw "File already exists: $path. Export again to use a new timestamp." }
            }
            $group.Group | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8
            Write-Log "Exported folder storage for $($group.Name): $path"
        }
    } catch { Show-Error $_ }
})
$ui.ReadDelegation.Add_Click({
    try {
        if ($script:busy) { return }
        $targets = Get-Targets
        if (-not $script:connected) { throw 'Connect to Exchange Online first.' }
        $script:delegateRows = @(); $ui.Delegates.ItemsSource = $null
        Start-Task @{Action='ReadDelegation';Targets=$targets}
    } catch { Show-Error $_ }
})
$ui.ExportDelegation.Add_Click({
    try {
        if (-not $script:delegateRows.Count) { throw 'Load delegates before exporting.' }
        $path = Get-SavePath 'CSV files (*.csv)|*.csv' ('MailboxDelegation-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.csv')
        if ($path) { $script:delegateRows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8; Write-Log "Delegation exported: $path" }
    } catch { Show-Error $_ }
})
$window.Add_Closing({ param($sender,$e) if ($script:busy) { $e.Cancel = $true; Write-Log 'An operation is running. Wait for completion before closing; closing cannot undo submitted changes.' } })
Write-Log 'Ready. No changes occur until a deletion or cancellation is explicitly confirmed.'
$timer.Start()
try { [void]$window.ShowDialog() } finally { $timer.Stop(); $script:runspace.Close(); $script:runspace.Dispose() }
