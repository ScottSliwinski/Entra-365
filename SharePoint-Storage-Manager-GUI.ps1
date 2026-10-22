#requires -Version 7.0
<#
SharePoint Storage Manager v1.2
Run in PowerShell 7.x on Windows:
  pwsh.exe -NoProfile -STA -File .\SharePoint-Storage-Manager.ps1
Prerequisite (run once):
  Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
Sign in with a SharePoint Administrator account. Supports MFA.
The GUI runs in PowerShell 7. The SharePoint module uses Microsoft's
-UseWindowsPowerShell compatibility layer (Windows PowerShell 5.1 must be installed).
Install the module from PowerShell 7; the script imports it automatically.
Scope: active SharePoint site collections in the connected geography; excludes OneDrive and redirect sites.
Does not modify tenant storage management mode. In SharePoint admin center:
  Settings > Site storage limits > Manual is required for individual quotas.
All displayed GB are binary (1024 MB). Usage can lag 24-48 hours.
Lowering quotas does not delete files or reclaim consumed storage.
Preview is enabled by default. Updates recheck usage and quotas, then verify results.
Use Export results to retain the audit before closing. No passwords are saved.
#>
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'This WPF GUI requires PowerShell 7.x on Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Start with pwsh.exe -STA -File .\SharePoint-Storage-Manager.ps1' }
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="SharePoint Storage Manager v1.2" Width="1220" Height="820" MinWidth="1000" MinHeight="700" WindowStartupLocation="CenterScreen">
<DockPanel Margin="14">
<StackPanel DockPanel.Dock="Top">
<TextBlock Text="SharePoint Storage Manager" FontSize="25" FontWeight="Bold"/>
<TextBlock Margin="0,6" TextWrapping="Wrap" Text="Site limits are growth ceilings. Reducing a quota does not reclaim consumed storage. Usage can lag 24-48 hours. OneDrive and redirect sites are excluded; multi-geo tenants must connect to each geography separately."/>
<WrapPanel Margin="0,8">
<TextBlock Text="Admin URL:" VerticalAlignment="Center"/><TextBox Name="AdminUrl" Width="370" Margin="8,0" Text="https://YOURTENANT-admin.sharepoint.com"/>
<Button Name="Connect" Content="Connect + load sites" Padding="12,5"/><Button Name="Refresh" Content="Refresh" Padding="12,5" Margin="8,0" IsEnabled="False"/>
<Button Name="Settings" Content="Storage settings" Padding="12,5"/>
</WrapPanel>
<TextBlock Name="Summary" FontSize="14" FontWeight="SemiBold" TextWrapping="Wrap" Text="Not connected." Margin="0,8"/>
<WrapPanel Margin="0,5">
<TextBlock Text="Search title or URL:" VerticalAlignment="Center"/><TextBox Name="Search" Width="300" Margin="8,0"/>
<CheckBox Name="LowOnly" Content="Only sites below 1% used" VerticalAlignment="Center" Margin="12,0"/>
<Button Name="Export" Content="Export visible sites CSV" Padding="10,4"/>
</WrapPanel>
<TextBlock Text="Select rows using Ctrl/Shift. GB and MB are binary. A tiny percentage does not mean the site is empty." Margin="0,5"/>
</StackPanel>
<StackPanel DockPanel.Dock="Bottom">
<WrapPanel Margin="0,10,0,5">
<TextBlock Text="New limit (GB):" VerticalAlignment="Center"/><TextBox Name="Quota" Width="85" Text="100" Margin="8,0"/>
<TextBlock Text="Warning (% of limit):" VerticalAlignment="Center"/><TextBox Name="Warning" Width="60" Text="90" Margin="8,0"/>
<CheckBox Name="Preview" Content="Preview only" IsChecked="True" VerticalAlignment="Center" Margin="10,0"/>
<Button Name="Apply" Content="Review selected changes" Padding="12,5" IsEnabled="False"/>
</WrapPanel>
<CheckBox Name="Manual" Content="I verified Settings &gt; Site storage limits is set to Manual for this tenant." Margin="0,5"/>
<TextBlock Name="Status" Text="Ready. Install the SharePoint Online module before connecting." TextWrapping="Wrap" Margin="0,5"/>
<ProgressBar Name="Progress" Height="5" Visibility="Collapsed"/>
<WrapPanel Margin="0,8,0,4"><TextBlock Text="Change results / audit" FontWeight="Bold"/><Button Name="ExportAudit" Content="Export results CSV" Margin="15,0" Padding="8,3"/></WrapPanel>
<DataGrid Name="Audit" Height="135" IsReadOnly="True" AutoGenerateColumns="True" CanUserAddRows="False"/>
</StackPanel>
<DataGrid Name="Sites" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended" SelectionUnit="FullRow" CanUserAddRows="False" EnableRowVirtualization="True">
<DataGrid.Columns>
<DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="190"/>
<DataGridTextColumn Header="Used MB" Binding="{Binding UsedMB}" Width="90"/>
<DataGridTextColumn Header="Used GB" Binding="{Binding UsedGB}" Width="90"/>
<DataGridTextColumn Header="Limit GB" Binding="{Binding LimitGB}" Width="90"/>
<DataGridTextColumn Header="Used %" Binding="{Binding UsedPercent}" Width="85"/>
<DataGridTextColumn Header="Warning GB" Binding="{Binding WarningGB}" Width="100"/>
<DataGridTextColumn Header="Lock" Binding="{Binding LockState}" Width="90"/>
<DataGridTextColumn Header="URL" Binding="{Binding Url}" Width="*"/>
</DataGrid.Columns>
</DataGrid>
</DockPanel>
</Window>
'@
$window = [Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
$ui = @{}
'AdminUrl','Connect','Refresh','Settings','Summary','Search','LowOnly','Export','Quota','Warning','Preview','Apply','Manual','Status','Progress','Audit','ExportAudit','Sites' | ForEach-Object { $ui[$_] = $window.FindName($_) }
$script:rows = @(); $script:results = [Collections.ObjectModel.ObservableCollection[object]]::new()
$ui.Audit.ItemsSource = $script:results
$script:connected = $false; $script:busy = $false; $script:connectionUrl = ''
# Keep module authentication and all network calls on one persistent worker thread.
$script:runspace = [RunspaceFactory]::CreateRunspace()
$script:runspace.ApartmentState = 'STA'
$script:runspace.ThreadOptions = 'ReuseThread'
$script:runspace.Open()
$script:worker = $null
function Show-Error($message) { [void][Windows.MessageBox]::Show($window,[string]$message,'SharePoint Storage Manager','OK','Error') }
function Update-View {
    $term = $ui.Search.Text.Trim()
    $visible = @($script:rows | Where-Object {
        ((-not $term) -or $_.Title.IndexOf($term,[StringComparison]::OrdinalIgnoreCase) -ge 0 -or $_.Url.IndexOf($term,[StringComparison]::OrdinalIgnoreCase) -ge 0) -and
        ((-not $ui.LowOnly.IsChecked) -or ($null -ne $_.UsedPercent -and $_.UsedPercent -lt 1))
    })
    $ui.Sites.ItemsSource = $visible
}
function Start-Work([scriptblock]$code, [object[]]$arguments, [scriptblock]$done) {
    if ($script:busy) { return }
    $script:busy = $true; $script:done = $done
    foreach ($name in 'Connect','Refresh','Apply','AdminUrl') { $ui[$name].IsEnabled = $false }
    $ui.Progress.Visibility = 'Visible'; $ui.Progress.IsIndeterminate = $true
    $script:worker = [PowerShell]::Create(); $script:worker.Runspace = $script:runspace
    [void]$script:worker.AddScript($code.ToString())
    foreach ($arg in $arguments) { [void]$script:worker.AddArgument($arg) }
    $script:handle = $script:worker.BeginInvoke()
}
$script:loadCode = {
    param($url,$connect)
    $ErrorActionPreference = 'Stop'
    if ($connect) {
        # Global scope retains compatibility proxy commands across worker invocations.
        if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
            throw 'Missing SharePoint module. In PowerShell 7 run: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser'
        }
        Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -Scope Global -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop
        Connect-SPOService -Url $url -ErrorAction Stop
    }
    $tenant = Get-SPOTenant
    # Exclude redirects by template and by the displayed title (case-insensitive).
    $sites = @(Get-SPOSite -Limit All -IncludePersonalSite $false | Where-Object {
        ([string]$_.Template -ine 'REDIRECTSITE#0') -and
        ([string]$_.Title).Trim() -ine 'RedirectSite'
    })
    $mode = 'Unknown - verify in admin center'
    if ($tenant.PSObject.Properties['StorageQuotaType']) { $mode = [string]$tenant.StorageQuotaType }
    $rows = @($sites | ForEach-Object {
        $used = [double]$_.StorageUsageCurrent; $quota = [double]$_.StorageQuota
        [pscustomobject]@{
            Title=[string]$_.Title; Url=[string]$_.Url; UsedMB=$used
            UsedGB=[math]::Round($used/1024,4); LimitGB=[math]::Round($quota/1024,4)
            UsedPercent=$(if ($quota -gt 0) { [math]::Round(100*$used/$quota,6) } else { $null })
            WarningGB=[math]::Round([double]$_.StorageQuotaWarningLevel/1024,4)
            LockState=[string]$_.LockState; QuotaMB=$quota
        }
    })
    [pscustomobject]@{ Rows=$rows; TotalGB=[math]::Round([double]$tenant.StorageQuota/1024,2)
        AllocatedGB=[math]::Round([double]$tenant.StorageQuotaAllocated/1024,2)
        SiteUsedGB=[math]::Round([double](($rows | Measure-Object UsedMB -Sum).Sum)/1024,2); Mode=$mode }
}
$script:loaded = {
    param($data)
    $d = @($data)[-1]
    $script:rows = @($d.Rows); $script:connected = $true
    $ui.Summary.Text = "Tenant capacity: $($d.TotalGB) GB | Reported allocation: $($d.AllocatedGB) GB | Loaded sites used: $($d.SiteUsedGB) GB | Sites: $($script:rows.Count) | Mode: $($d.Mode)"
    $ui.Status.Text = 'Loaded. Reported allocation is not actual usage; site totals are scoped to this connection.'
    Update-View
}
$ui.Connect.Add_Click({
    try {
        $url = [uri]$ui.AdminUrl.Text.Trim()
        if ($url.Scheme -ne 'https' -or $url.Host -notmatch '-admin\.sharepoint\.(com|us|de|cn)$' -or $url.AbsolutePath -ne '/' -or $url.Query -or $url.Fragment -or $url.UserInfo) { throw 'Enter a valid SharePoint admin URL, such as https://contoso-admin.sharepoint.com' }
        $script:connected = $false; $script:rows = @(); Update-View
        $ui.Manual.IsChecked = $false; $script:connectionUrl = $url.AbsoluteUri.TrimEnd('/')
        $ui.Status.Text = 'Connecting and loading sites. Complete the sign-in window when prompted.'
        Start-Work $script:loadCode @($script:connectionUrl,$true) $script:loaded
    } catch { Show-Error $_.Exception.Message }
})
$ui.Refresh.Add_Click({ $ui.Status.Text = 'Refreshing site storage...'; Start-Work $script:loadCode @($script:connectionUrl,$false) $script:loaded })
$ui.Settings.Add_Click({
    if ($script:connectionUrl) { Start-Process ($script:connectionUrl + '/_layouts/15/online/AdminHome.aspx#/settings') }
    else { Show-Error 'Connect first, then open Settings > Site storage limits.' }
})
$ui.Search.Add_TextChanged({ Update-View }); $ui.LowOnly.Add_Click({ Update-View })
function Export-Data($items,$filename) {
    try {
        if (@($items).Count -eq 0) { throw 'There are no rows to export.' }
        $dialog = [Windows.Forms.SaveFileDialog]::new(); $dialog.Filter = 'CSV (*.csv)|*.csv'; $dialog.FileName = $filename
        if ($dialog.ShowDialog() -eq 'OK') { $items | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8; $ui.Status.Text = "Exported $($dialog.FileName)" }
        $dialog.Dispose()
    } catch { Show-Error $_.Exception.Message }
}
$ui.Export.Add_Click({ Export-Data @($ui.Sites.ItemsSource) 'SharePoint-Site-Storage.csv' })
$ui.ExportAudit.Add_Click({ Export-Data @($script:results) 'SharePoint-Storage-Changes.csv' })
$ui.Apply.Add_Click({
    try {
        $selected = @($ui.Sites.SelectedItems)
        if (-not $selected.Count) { throw 'Select one or more site rows first.' }
        $gb = [double]::Parse($ui.Quota.Text,[Globalization.CultureInfo]::CurrentCulture)
        $percent = [double]::Parse($ui.Warning.Text,[Globalization.CultureInfo]::CurrentCulture)
        if ([double]::IsNaN($gb) -or [double]::IsInfinity($gb) -or $gb -lt 1 -or $gb -gt 25600) { throw 'Enter a limit from 1 to 25600 GB.' }
        if ([double]::IsNaN($percent) -or $percent -lt 1 -or $percent -gt 99) { throw 'Warning percentage must be between 1 and 99.' }
        $mb = [long][math]::Floor($gb*1024); $warningMB = [long][math]::Floor($mb*$percent/100)
        $preview = [bool]$ui.Preview.IsChecked
        if (-not $preview -and -not $ui.Manual.IsChecked) { throw 'First verify Manual site storage limits in the admin center, then check the confirmation box.' }
        $plan = @($selected | ForEach-Object { [pscustomobject]@{Url=$_.Url; OldMB=$_.QuotaMB} })
        $detail = ($selected | Select-Object -First 8 | ForEach-Object { "$($_.Title): $($_.LimitGB) GB -> $($mb/1024) GB" }) -join "`n"
        if ($selected.Count -gt 8) { $detail += "`n... plus $($selected.Count-8) more sites." }
        $action = if ($preview) {'PREVIEW ONLY'} else {'APPLY LIVE CHANGES'}
        if ([Windows.MessageBox]::Show($window,"$action`n$detail`n`nWarning: $percent%. Sites at or above the new limit will be skipped. Continue?",'Review site quotas','YesNo','Warning') -ne 'Yes') { return }
        $ui.Status.Text = "$action - checking $($selected.Count) sites..."
        Start-Work {
            param($plan,$mb,$warningMB,$preview)
            $ErrorActionPreference = 'Stop'
            foreach ($p in $plan) {
                $old = $p.OldMB; $used = $null; $outcome = 'Failed'; $note = ''; $attempted = $false
                try {
                    $site = Get-SPOSite -Identity $p.Url
                    $used = [double]$site.StorageUsageCurrent
                    $old = [double]$site.StorageQuota
                    if ($old -ne $p.OldMB) { throw 'Quota changed since load. Refresh and review again.' }
                    if ($used -ge $mb) { throw 'New quota must exceed currently reported usage.' }
                    if ([string]$site.LockState -ne 'Unlock') { throw "Site lock state is $($site.LockState); skipped." }
                    if ($preview) { $outcome = 'Preview'; $note = 'Validated; no change made.' }
                    else {
                        $attempted = $true
                        Set-SPOSite -Identity $p.Url -StorageQuota $mb -StorageQuotaWarningLevel $warningMB
                        $check = Get-SPOSite -Identity $p.Url
                        if ([long]$check.StorageQuota -eq $mb -and [long]$check.StorageQuotaWarningLevel -eq $warningMB) { $outcome = 'Verified'; $note = 'Quota and warning verified.' }
                        else { $outcome = 'Check required'; $note = 'Command returned; reported settings differ. Refresh and inspect storage management mode.' }
                    }
                } catch {
                    if ($attempted) { $outcome = 'Check required'; $note = 'Update attempted; inspect current settings before retrying. ' + $_.Exception.Message }
                    else { $outcome = 'Skipped'; $note = $_.Exception.Message }
                }
                [pscustomobject]@{ Time=(Get-Date).ToString('s'); Url=$p.Url; OldLimitGB=$old/1024; NewLimitGB=$mb/1024; UsedMB=$used; WarningMB=$warningMB; Result=$outcome; Detail=$note }
            }
        } @($plan,$mb,$warningMB,$preview) {
            param($data)
            foreach ($row in $data) { $script:results.Add($row) }
            $ui.Status.Text = 'Finished. Review/export results. Refresh sites to see current limits before further edits.'
        }
    } catch { Show-Error $_.Exception.Message }
})
$timer = [Windows.Threading.DispatcherTimer]::new(); $timer.Interval = [TimeSpan]::FromMilliseconds(250)
$timer.Add_Tick({
    if ($script:busy -and $script:handle.IsCompleted) {
        try {
            $data = $script:worker.EndInvoke($script:handle)
            if ($script:worker.HadErrors) { throw (($script:worker.Streams.Error | ForEach-Object { $_.ToString() }) -join "`n") }
            & $script:done $data
        } catch { $ui.Status.Text = 'Operation failed. See error details.'; Show-Error $_.Exception.Message }
        finally {
            $script:worker.Dispose(); $script:worker = $null; $script:busy = $false
            $ui.Progress.Visibility = 'Collapsed'; $ui.Connect.IsEnabled = $true; $ui.AdminUrl.IsEnabled = $true
            $ui.Refresh.IsEnabled = $script:connected; $ui.Apply.IsEnabled = $script:connected
        }
    }
})
$window.Add_Closing({ param($sender,$e)
    if ($script:busy) { $e.Cancel = $true; Show-Error 'An operation is running. Wait for it to finish before closing.' }
})
$timer.Start()
try { [void]$window.ShowDialog() }
finally { $timer.Stop(); if ($script:worker) { $script:worker.Dispose() }; $script:runspace.Dispose() }
