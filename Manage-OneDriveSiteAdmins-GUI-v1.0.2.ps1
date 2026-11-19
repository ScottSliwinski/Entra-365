#requires -Version 7.0
# Owner: Scott Sliwinski
# Creator: Scott Sliwinski
# Author: Scott Sliwinski
# Website: https://www.caeitsolutions.com/
# LinkedIn: https://www.linkedin.com/in/scottsliwinski/
# Revision: 1.0.2
# Baseline: Current reference v1.0.2; awaiting tenant validation.
# Changes: Fixed UPN validation syntax and removed duplicated event-handler tail.
<#
Windows PowerShell 7 WPF GUI. Launch in a fresh session:
  pwsh.exe -STA -File .\Manage-OneDriveSiteAdmins-GUI-v1.0.2.ps1
Requires SharePoint Administrator role and Graph User.Read.All consent.
Graph and SharePoint sign-ins MUST use the same tenant. Enter its tenant ID,
SharePoint admin URL, and the account used for BOTH sign-ins (optional UPN).
Reuse current connection uses a delegated Graph context visible to this process;
it cannot attach to a separate PowerShell window or inherit browser cookies.
Sign-in window uses Microsoft's interactive Graph authentication (Windows may
show the Windows account picker). SharePoint browser sign-in uses the system
browser, where an existing Microsoft login may be reused subject to policy.
Missing modules can be installed for CurrentUser after confirmation.
SharePoint uses the Windows PowerShell compatibility session (Windows required).
Operations run in one persistent background runspace; the UI stays responsive.
No OneDrive is provisioned, and no URL is guessed. Lookup uses the site Owner.
An optional cached inventory handles renamed UPNs (select the verified site).
Granting yourself access is persistent and explicit; remove it when finished.
A copied site URL is NOT an anonymous or permission-granting sharing link.
#>
[CmdletBinding()]param()
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'This WPF tool requires Windows and PowerShell 7.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Launch with: pwsh.exe -STA -File "Manage-OneDriveSiteAdmins-GUI-v1.0.2.ps1"'
}
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
function Confirm-Action([string]$Text) {
    [Windows.MessageBox]::Show($Text,'Confirm change','YesNo','Warning') -eq 'Yes'
}
$worker = [runspacefactory]::CreateRunspace()
$worker.ApartmentState = 'STA'
$worker.ThreadOptions = 'ReuseThread'
$worker.Open()
$bootstrap = @'
$ErrorActionPreference = 'Stop'
$script:Connected = $false
$script:Inventory = @()
function Find-People([string]$Text) {
    $t = $Text.Trim()
    if ($t.Length -lt 2) { throw 'Enter at least two characters.' }
    if ($t.Contains('@')) {
        try { return @(Get-MgUser -UserId $t -Property Id,DisplayName,UserPrincipalName,Mail -ErrorAction Stop | Select-Object Id,DisplayName,UserPrincipalName,Mail) }
        catch { if ($_.Exception.Message -notmatch '404|Request_ResourceNotFound|does not exist') { throw } }
    }
    $q = $t.Replace("'", "''")
    @(Get-MgUser -Filter "startsWith(displayName,'$q') or startsWith(userPrincipalName,'$q')" -Top 100 -Property Id,DisplayName,UserPrincipalName,Mail | Select-Object Id,DisplayName,UserPrincipalName,Mail)
}
function Read-Admins([string]$Url) {
    @(Get-SPOUser -Site $Url -Limit All -ErrorAction Stop | Where-Object IsSiteAdmin | Select-Object DisplayName,LoginName,IsSiteAdmin)
}
function Invoke-Task($Operation,$Data) {
    switch ($Operation) {
        'Connect' {
            $script:Connected = $false
            $script:Inventory = @()
            foreach ($m in @('Microsoft.Graph.Authentication','Microsoft.Graph.Users','Microsoft.Online.SharePoint.PowerShell')) {
                if (-not (Get-Module -ListAvailable -Name $m)) {
                    if (-not $Data.Install) { throw "Missing module: $m. Enable installation and reconnect." }
                    Install-Module -Name $m -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                }
            }
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
            $ctx = Get-MgContext
            if ($Data.Reuse) {
                if (-not $ctx -or $ctx.AuthType -ne 'Delegated') {
                    throw 'No current delegated Graph connection is available in this process. Choose Microsoft sign-in window.'
                }
                if ($ctx.TenantId -ine $Data.Tenant) {
                    throw "Current connection is for tenant $($ctx.TenantId). Choose Microsoft sign-in window to connect to the entered tenant."
                }
            } else {
                $connect = @{TenantId=$Data.Tenant; Scopes=@('User.Read.All'); ContextScope='Process'; NoWelcome=$true; ErrorAction='Stop'}
                # Newer SDKs support LoginHint. Do not require it on older versions.
                if ($Data.Operator -and (Get-Command Connect-MgGraph).Parameters.ContainsKey('LoginHint')) {
                    $connect.LoginHint = $Data.Operator
                }
                # Interactive Microsoft sign-in; no forced sign-out or device-code flow.
                Connect-MgGraph @connect | Out-Null
                $ctx = Get-MgContext
            }
            if (-not $ctx -or $ctx.AuthType -ne 'Delegated' -or $ctx.TenantId -ine $Data.Tenant) {
                throw 'A delegated connection to the specified tenant is required.'
            }
            if ($Data.Operator -and $ctx.Account -ine $Data.Operator) {
                throw "Signed in as $($ctx.Account). Enter that UPN or reconnect with the intended administrator."
            }
            if (-not $ctx.Account) { throw 'The current connection has no signed-in user.' }
            # Validate the token and directory-read capability before enabling the GUI.
            Get-MgUser -Top 1 -Property Id -ErrorAction Stop | Out-Null
            $me = Get-MgUser -UserId $ctx.Account -Property Id,UserPrincipalName -ErrorAction Stop
            $script:Operator = $me.UserPrincipalName
            $spoConnect = @{Url=$Data.AdminUrl; ErrorAction='Stop'}
            $spoCommand = Get-Command Connect-SPOService
            if ($spoCommand.Parameters.ContainsKey('UseSystemBrowser')) {
                $spoConnect.UseSystemBrowser = [bool]$Data.Browser
            } elseif ($Data.Browser) {
                throw 'This SharePoint module does not support system-browser sign-in. Update Microsoft.Online.SharePoint.PowerShell, restart PowerShell, and try again.'
            }
            Connect-SPOService @spoConnect
            Get-SPOTenant -ErrorAction Stop | Out-Null
            $script:Connected = $true
            return "Connected as $script:Operator. Graph tenant: $($ctx.TenantId). Confirm SharePoint sign-in used this same tenant/account."
        }
        default { if (-not $script:Connected) { throw 'Connect first.' } }
    }
    switch ($Operation) {
        'People' { return Find-People $Data.Text }
        'Sites' {
            $q = $Data.UPN.Replace("'", "''")
            return @(Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Owner -eq '$q'" | Where-Object { $_.Template -like 'SPSPERS*' } | Select-Object Url,Owner,Title,LockState)
        }
        'Inventory' {
            $script:Inventory = @(Get-SPOSite -IncludePersonalSite $true -Limit All | Where-Object { $_.Template -like 'SPSPERS*' } | Select-Object Url,Owner,Title,LockState)
            return $script:Inventory
        }
        'FilterInventory' {
            $q = $Data.Text
            return @($script:Inventory | Where-Object { $_.Owner -like "*$q*" -or $_.Url -like "*$q*" -or $_.Title -like "*$q*" })
        }
        'Admins' { return Read-Admins $Data.Url }
        'Add' {
            Set-SPOUser -Site $Data.Url -LoginName $Data.UPN -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
            $check = Get-SPOUser -Site $Data.Url -LoginName $Data.UPN -ErrorAction Stop
            if (-not $check.IsSiteAdmin) { throw 'The server did not confirm administrator access.' }
            return "Verified administrator access: $($Data.UPN). Refresh the administrator list."
        }
        'Self' {
            Set-SPOUser -Site $Data.Url -LoginName $script:Operator -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
            return "Granted $script:Operator administrator access. This persists until you remove it. Refresh the list."
        }
        'Remove' {
            $site = Get-SPOSite -Identity $Data.Url -ErrorAction Stop
            $owner = ($site.Owner -split '\|')[-1]
            $target = ($Data.Login -split '\|')[-1]
            if ($owner -ieq $target) { throw 'The primary OneDrive owner is protected from removal.' }
            $admins = @(Read-Admins $Data.Url)
            if ($admins.Count -le 1) { throw 'Cannot remove the last site administrator.' }
            Set-SPOUser -Site $Data.Url -LoginName $Data.Login -IsSiteCollectionAdmin $false -ErrorAction Stop | Out-Null
            return "Removed administrator status for $($Data.Login). Other permissions may still allow access. Refresh the list."
        }
    }
}
'@
$p = [powershell]::Create(); $p.Runspace = $worker
[void]$p.AddScript($bootstrap); $null = $p.Invoke()
if ($p.HadErrors) { throw ($p.Streams.Error | Out-String) }; $p.Dispose()
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="OneDrive Site Administrators | Scott Sliwinski | v1.0.2" Width="1120" Height="900" MinWidth="950" MinHeight="760" WindowStartupLocation="CenterScreen" Background="#F3F5F8">
<DockPanel Margin="18">
<TextBlock DockPanel.Dock="Top" Text="OneDrive • Site Collection Administrators" FontSize="25" FontWeight="Bold" Margin="0,0,0,12"/>
<TextBox x:Name="Status" DockPanel.Dock="Bottom" Height="82" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Margin="0,10,0,0" Text="Enter the tenant ID and SharePoint admin URL. Choose an existing connection or Microsoft sign-in window."/>
<ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel>
<GroupBox Header="1. Connect — use the same tenant and account for both sign-ins" Padding="10"><StackPanel>
<TextBlock Text="Entra tenant ID (GUID)"/><TextBox x:Name="Tenant" Margin="0,3,0,7"/>
<TextBlock Text="SharePoint admin URL (https://yourtenant-admin.sharepoint.com)"/><TextBox x:Name="AdminUrl" Margin="0,3,0,7"/>
<TextBlock Text="Administrator sign-in UPN (optional; checks the account if entered)"/><TextBox x:Name="Operator" Margin="0,3,0,7"/>
<TextBlock Text="Graph connection" Margin="0,2,0,3"/>
<ComboBox x:Name="AuthMode" SelectedIndex="0" Margin="0,0,0,7">
<ComboBoxItem Content="Microsoft sign-in window (current or another account)"/>
<ComboBoxItem Content="Reuse current Graph connection in this PowerShell process"/>
</ComboBox>
<CheckBox x:Name="Browser" Content="Open SharePoint sign-in in my default browser (can reuse its login session)" IsChecked="True" Margin="0,0,0,7"/>
<TextBlock Text="Graph may use the Windows account picker. Browser sessions and Graph connections are separate; Microsoft may still request MFA." TextWrapping="Wrap" Foreground="#586475" Margin="0,0,0,7"/>
<WrapPanel><CheckBox x:Name="Install" Content="Install missing modules for my Windows account" VerticalAlignment="Center" Margin="0,0,18,0"/><Button x:Name="Connect" Content="Connect" Padding="18,5"/></WrapPanel>
</StackPanel></GroupBox>
<GroupBox Header="2. Find the OneDrive owner" Padding="10" Margin="0,10,0,0"><StackPanel>
<DockPanel><Button x:Name="SearchOwner" DockPanel.Dock="Right" Content="Search user" Padding="12,4"/><TextBox x:Name="OwnerQuery" Margin="0,0,8,0" ToolTip="Full UPN or start of display name"/></DockPanel>
<DataGrid x:Name="Owners" Height="105" Margin="0,7,0,7" AutoGenerateColumns="True" IsReadOnly="True" SelectionMode="Single"/>
<WrapPanel><Button x:Name="FindSite" Content="Find selected user's OneDrive" Padding="12,4" Margin="0,0,8,0"/><Button x:Name="Inventory" Content="Load all OneDrives (may be slow)" Padding="12,4"/></WrapPanel>
<DockPanel Margin="0,7,0,0"><Button x:Name="FilterSites" DockPanel.Dock="Right" Content="Filter cached sites" Padding="12,4"/><TextBox x:Name="SiteQuery" Margin="0,0,8,0" ToolTip="Search cached inventory by owner, URL or title; useful after a UPN rename"/></DockPanel>
<DataGrid x:Name="Sites" Height="100" Margin="0,7,0,7" AutoGenerateColumns="True" IsReadOnly="True" SelectionMode="Single"/>
<DockPanel><Button x:Name="Copy" DockPanel.Dock="Right" Content="Copy link" Padding="12,4"/><Button x:Name="Open" DockPanel.Dock="Right" Content="Open" Padding="12,4" Margin="0,0,6,0"/><TextBox x:Name="Link" IsReadOnly="True" Margin="0,0,8,0"/></DockPanel>
<TextBlock Text="This link requires existing permission. It does not grant access." Foreground="#586475" Margin="0,4,0,0"/>
</StackPanel></GroupBox>
<GroupBox Header="3. Manage administrators on the selected OneDrive" Padding="10" Margin="0,10,0,0"><StackPanel>
<WrapPanel><Button x:Name="Refresh" Content="Show / refresh administrators" Padding="12,4" Margin="0,0,8,0"/><Button x:Name="Self" Content="Grant my admin account access" Padding="12,4" Margin="0,0,8,0"/><Button x:Name="Remove" Content="Remove selected administrator" Padding="12,4"/></WrapPanel>
<DataGrid x:Name="Admins" Height="120" Margin="0,7,0,7" AutoGenerateColumns="True" IsReadOnly="True" SelectionMode="Single"/>
<DockPanel><Button x:Name="SearchAdmin" DockPanel.Dock="Right" Content="Search administrator to add" Padding="12,4"/><TextBox x:Name="AdminQuery" Margin="0,0,8,0" ToolTip="Full UPN or start of display name"/></DockPanel>
<DataGrid x:Name="Candidates" Height="100" Margin="0,7,0,7" AutoGenerateColumns="True" IsReadOnly="True" SelectionMode="Single"/>
<Button x:Name="Add" Content="Add selected user as site administrator" Padding="12,5" HorizontalAlignment="Left"/>
</StackPanel></GroupBox>
</StackPanel></ScrollViewer></DockPanel></Window>
'@
$window = [Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
$ui = @{}
foreach ($n in @('Status','Tenant','AdminUrl','Operator','AuthMode','Browser','Install','Connect','OwnerQuery','SearchOwner','Owners','FindSite','Inventory','SiteQuery','FilterSites','Sites','Link','Copy','Open','Refresh','Self','Remove','Admins','AdminQuery','SearchAdmin','Candidates','Add')) { $ui[$n] = $window.FindName($n) }
$script:Pending = $null
$script:Connected = $false
$script:Site = $null
$buttons = @('Connect','SearchOwner','FindSite','Inventory','FilterSites','Refresh','Self','Remove','SearchAdmin','Add')
function Start-Task([string]$Operation,[hashtable]$Data,[string]$Target) {
    if ($script:Pending) { return }
    if ($Operation -ne 'Connect' -and -not $script:Connected) { throw 'Connect first.' }
    foreach ($b in $buttons) { $ui[$b].IsEnabled = $false }
    $ui.Sites.IsEnabled = $false
    $ui.Status.Text = "Working: $Operation ... Complete the Microsoft sign-in window or browser prompt if shown."
    $ps = [powershell]::Create(); $ps.Runspace = $worker
    [void]$ps.AddScript('param($op,$data) Invoke-Task $op $data').AddArgument($Operation).AddArgument($Data)
    $script:Pending = @{ PowerShell=$ps; Handle=$ps.BeginInvoke(); Target=$Target; Operation=$Operation }
}
function Selected-Site {
    if (-not $script:Site) { throw 'Select a OneDrive site first.' }
    [string]$script:Site.Url
}
function Guard([scriptblock]$Action) {
    try { & $Action } catch { $ui.Status.Text = $_.Exception.Message }
}
$timer = [Windows.Threading.DispatcherTimer]::new()
$timer.Interval = [timespan]::FromMilliseconds(250)
$timer.Add_Tick({
    if (-not $script:Pending -or -not $script:Pending.Handle.IsCompleted) { return }
    $task = $script:Pending
    try {
        $result = @($task.PowerShell.EndInvoke($task.Handle))
        if ($task.PowerShell.HadErrors) { throw ($task.PowerShell.Streams.Error | Out-String) }
        if ($task.Operation -eq 'Connect') { $script:Connected = $true }
        if ($task.Target) {
            $ui[$task.Target].ItemsSource = $result
            $ui.Status.Text = "Completed: $($result.Count) result(s)."
            if ($task.Target -eq 'Sites' -and $result.Count -eq 0) {
                $ui.Status.Text = 'No existing OneDrive found by primary owner. It may be unprovisioned, renamed, in another geo, or have a changed owner. Load all OneDrives to locate it manually.'
            }
            if ($task.Target -in @('Owners','Candidates') -and $result.Count -ge 100) { $ui.Status.Text += ' First 100 users shown; narrow the search.' }
        } else { $ui.Status.Text = $result -join [Environment]::NewLine }
    } catch {
        $ui.Status.Text = $_.Exception.Message
        if ($task.Operation -eq 'Admins') { $ui.Status.Text += "`r`nIf site access is denied, explicitly grant your admin account access, then refresh. Access will persist until removed." }
    } finally {
        $task.PowerShell.Dispose(); $script:Pending = $null
        foreach ($b in $buttons) { $ui[$b].IsEnabled = $true }
        $ui.Sites.IsEnabled = $true
    }
})
$ui.Connect.Add_Click({ Guard {
    if ($ui.Tenant.Text.Trim() -notmatch '^[0-9a-fA-F-]{36}$') { throw 'Enter the Entra tenant GUID.' }
    $uri = $null
    if (-not [uri]::TryCreate($ui.AdminUrl.Text.Trim(),[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.Host -notmatch '^[a-zA-Z0-9-]+-admin\.sharepoint\.com$' -or $uri.AbsolutePath -ne '/' -or $uri.Query -or $uri.UserInfo) { throw 'Enter a commercial-cloud SharePoint admin URL, for example https://contoso-admin.sharepoint.com.' }
    if ($ui.Operator.Text.Trim() -and $ui.Operator.Text.Trim() -notmatch '^\S+@\S+\.\S+$') { throw 'Enter your administrator sign-in UPN.' }
    if ($ui.Install.IsChecked -and -not (Confirm-Action 'Install any missing Microsoft Graph / SharePoint modules from PSGallery for your Windows account?')) { return }
    $script:Connected = $false; $script:Site = $null
    foreach ($n in @('Owners','Sites','Admins','Candidates')) { $ui[$n].ItemsSource = @() }; $ui.Link.Text = ''
    Start-Task Connect @{Tenant=$ui.Tenant.Text.Trim();AdminUrl=$uri.AbsoluteUri.TrimEnd('/');Operator=$ui.Operator.Text.Trim();Install=[bool]$ui.Install.IsChecked;Reuse=($ui.AuthMode.SelectedIndex -eq 1);Browser=[bool]$ui.Browser.IsChecked} ''
} })
$ui.SearchOwner.Add_Click({ Guard { Start-Task People @{Text=$ui.OwnerQuery.Text} Owners } })
$ui.SearchAdmin.Add_Click({ Guard { Start-Task People @{Text=$ui.AdminQuery.Text} Candidates } })
function Clear-Site {
    $script:Site = $null; $ui.Link.Text = ''; $ui.Admins.ItemsSource = @(); $ui.Sites.ItemsSource = @()
}
$ui.FindSite.Add_Click({ Guard {
    $u = $ui.Owners.SelectedItem; if (-not $u) { throw 'Select an owner search result.' }
    Clear-Site; Start-Task Sites @{UPN=$u.UserPrincipalName} Sites
} })
$ui.Inventory.Add_Click({ Guard { Clear-Site; Start-Task Inventory @{} Sites } })
$ui.FilterSites.Add_Click({ Guard { Clear-Site; Start-Task FilterInventory @{Text=$ui.SiteQuery.Text.Trim()} Sites } })
$ui.Sites.Add_SelectionChanged({
    $script:Site = $ui.Sites.SelectedItem; $ui.Admins.ItemsSource = @()
    $ui.Link.Text = if ($script:Site) { [string]$script:Site.Url } else { '' }
})
$ui.Refresh.Add_Click({ Guard { $url = Selected-Site; $ui.Admins.ItemsSource = @(); Start-Task Admins @{Url=$url} Admins } })
$ui.Self.Add_Click({ Guard {
    $url = Selected-Site
    if (Confirm-Action "Grant the connected administrator account FULL site administrator access to:`n$url`n`nThis persists until explicitly removed. Use the same administrator for both sign-ins.") { Start-Task Self @{Url=$url} '' }
} })
$ui.Add.Add_Click({ Guard {
    $url = Selected-Site; $u = $ui.Candidates.SelectedItem
    if (-not $u) { throw 'Search for and select an administrator to add.' }
    if (Confirm-Action "Grant FULL OneDrive site administrator access to:`n$($u.DisplayName) <$($u.UserPrincipalName)>`n`nSite: $url") { Start-Task Add @{Url=$url;UPN=$u.UserPrincipalName} '' }
} })
$ui.Remove.Add_Click({ Guard {
    $url = Selected-Site; $a = $ui.Admins.SelectedItem
    if (-not $a) { throw 'Select an administrator from the current list.' }
    if (Confirm-Action "Remove site administrator status for:`n$($a.LoginName)`n`nSite: $url`n`nRemoving yourself may prevent further listing. Other permissions are unaffected.") {
        $ui.Admins.ItemsSource = @(); Start-Task Remove @{Url=$url;Login=$a.LoginName} ''
    }
} })
$ui.Copy.Add_Click({ Guard { $url = Selected-Site; [Windows.Clipboard]::SetText($url); $ui.Status.Text = 'OneDrive URL copied. The recipient must already have permission.' } })
$ui.Open.Add_Click({ Guard { Start-Process (Selected-Site) } })
$window.Add_Closing({ param($sender,$eventArgs)
    if ($script:Pending) { $eventArgs.Cancel = $true; $ui.Status.Text = 'Wait for the current operation to finish before closing.' }
})
$timer.Start()
try { $null = $window.ShowDialog() } finally { $timer.Stop(); $worker.Close(); $worker.Dispose() }

