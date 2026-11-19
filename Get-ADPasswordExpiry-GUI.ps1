#Requires -Version 7.0
# Owner: Scott Sliwinski
# Creator: Scott Sliwinski
# Author: Scott Sliwinski
# Website: https://www.caeitsolutions.com/
# LinkedIn: https://www.linkedin.com/in/scottsliwinski/
# Revision: 1.1.0

if (-not $IsWindows) { throw 'This GUI requires Windows.' }
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    throw 'Launch with pwsh -STA -File Get-ADPasswordExpiry-GUI.ps1.'
}

Add-Type -AssemblyName PresentationFramework -ErrorAction Stop

[xml]$layout = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="AD Password Expiration — v1.1.0" Width="560" MinWidth="430"
        SizeToContent="Height" WindowStartupLocation="CenterScreen">
    <StackPanel Margin="22">
        <Label Content="User name or UPN" Target="{Binding ElementName=Upn}" Padding="0,0,0,5" />
        <DockPanel>
            <Button Name="Search" DockPanel.Dock="Right" Content="Search AD"
                    Padding="12,5" Margin="8,0,0,0" />
            <TextBox Name="Upn" Padding="6" />
        </DockPanel>
        <Label Content="Matching users" Target="{Binding ElementName=Matches}" Padding="0,8,0,5" />
        <ComboBox Name="Matches" IsEnabled="False" MaxDropDownHeight="240"
                  DisplayMemberPath="Label" />
        <Label Content="AD domain or domain controller (optional)" Target="{Binding ElementName=Server}" Padding="0,12,0,5" />
        <TextBox Name="Server" Padding="6" />
        <Button Name="Check" Content="Check password expiration"
                HorizontalAlignment="Left" Padding="12,6" Margin="0,14,0,14" />
        <TextBox Name="Result" IsReadOnly="True" TextWrapping="Wrap" Padding="9"
                 MinHeight="125" MaxHeight="250" VerticalScrollBarVisibility="Auto" />
    </StackPanel>
</Window>
'@
$window = [Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($layout))
$upnBox = $window.FindName('Upn')
$serverBox = $window.FindName('Server')
$resultBox = $window.FindName('Result')
$checkButton = $window.FindName('Check')

$searchButton = $window.FindName('Search')
$matchesBox = $window.FindName('Matches')

function Import-LookupModule {
    if (-not (Get-Module ActiveDirectory)) {
        if (-not (Get-Module -ListAvailable ActiveDirectory)) {
            throw 'Install the Active Directory PowerShell module through Windows RSAT, then retry.'
        }
        Import-Module ActiveDirectory -ErrorAction Stop
    }
}

$searchButton.Add_Click({
    $matchesBox.ItemsSource = $null
    $matchesBox.IsEnabled = $false
    $term = $upnBox.Text.Trim()
    if ($term.Length -lt 2) {
        $resultBox.Text = 'Enter at least two characters from a name or UPN.'
        return
    }

    $resultBox.Clear()
    $window.Cursor = [Windows.Input.Cursors]::Wait
    try {
        Import-LookupModule

        # Escape LDAP filter metacharacters so the search text is treated literally.
        $term = $term.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace([string][char]0, '\00')
        $query = @{
            LDAPFilter = "(&(userPrincipalName=*)(|(displayName=*$term*)(name=*$term*)(userPrincipalName=*$term*)))"
            Properties = 'DisplayName'
            ResultSetSize = 51
            ErrorAction = 'Stop'
        }
        if ($serverBox.Text.Trim()) { $query.Server = $serverBox.Text.Trim() }

        $users = @(Get-ADUser @query)
        if ($users.Count -eq 0) {
            $resultBox.Text = 'No matching users in the queried AD domain.'
            return
        }

        $matchesBox.ItemsSource = @(
            $users | Sort-Object DisplayName, UserPrincipalName | Select-Object -First 50 | ForEach-Object {
                [pscustomobject]@{
                    Label = "$($_.DisplayName) <$($_.UserPrincipalName)>"
                    Upn = $_.UserPrincipalName
                }
            }
        )
        $matchesBox.IsEnabled = $true
        $resultBox.Text = if ($users.Count -gt 50) {
            'Showing up to 50 matches. Select a user or enter a more specific search.'
        } else {
            'Select a matching user, then click Check password expiration.'
        }
        $matchesBox.IsDropDownOpen = $true
    } catch {
        $resultBox.Text = $_.Exception.Message
    } finally {
        $window.Cursor = $null
    }
})

$matchesBox.Add_SelectionChanged({
    if ($null -ne $matchesBox.SelectedItem) {
        $upnBox.Text = $matchesBox.SelectedItem.Upn
    }
})

$serverBox.Add_TextChanged({
    $matchesBox.ItemsSource = $null
    $matchesBox.IsEnabled = $false
    $resultBox.Clear()
})

$checkButton.Add_Click({
    $upn = $upnBox.Text.Trim()
    if ($upn -notmatch '^[^\s@]+@[^\s@]+$') {
        $resultBox.Text = 'Enter a full UPN such as user@example.com.'
        [void]$upnBox.Focus()
        return
    }

    $resultBox.Clear()
    $window.Cursor = [Windows.Input.Cursors]::Wait
    try {
        Import-LookupModule

        $query = @{
            Filter = 'UserPrincipalName -eq $upn'
            Properties = @('DisplayName', 'pwdLastSet', 'msDS-UserPasswordExpiryTimeComputed')
            ResultSetSize = 2
            ErrorAction = 'Stop'
        }
        if ($serverBox.Text.Trim()) { $query.Server = $serverBox.Text.Trim() }

        $users = @(Get-ADUser @query)
        if ($users.Count -eq 0) {
            throw 'UPN was not found in the queried AD domain. Specify the AD domain or domain controller if needed.'
        }
        if ($users.Count -gt 1) { throw 'Multiple AD accounts have this UPN.' }

        $user = $users[0]
        $expiry = $user.'msDS-UserPasswordExpiryTimeComputed'
        $lastSet = 'Unavailable'
        $expiryText = 'Unavailable'
        $status = 'Expiration could not be determined'

        if ([long]$user.pwdLastSet -gt 0) {
            $lastSet = [DateTimeOffset]::FromFileTime([long]$user.pwdLastSet).ToString('yyyy-MM-dd HH:mm:ss zzz')
        } elseif ($null -ne $user.pwdLastSet) {
            $lastSet = 'No timestamp recorded'
        }

        if ($null -ne $expiry) {
            if ([long]$expiry -eq [long]::MaxValue) {
                $status = 'No scheduled password expiration'
                $expiryText = 'None'
            } elseif ([long]$expiry -eq 0) {
                $status = 'Password change required at next sign-in'
                $expiryText = 'No expiration timestamp recorded'
            } else {
                $expiryDate = [DateTimeOffset]::FromFileTime([long]$expiry)
                $expiryText = $expiryDate.ToString('yyyy-MM-dd HH:mm:ss zzz')
                $status = if ($expiryDate -le [DateTimeOffset]::Now) { 'Expired' } else { 'Not expired' }
            }
        }

        $resultBox.Text = @(
            "User: $($user.DisplayName) <$($user.UserPrincipalName)>"
            "Status: $status"
            "Password last set: $lastSet"
            "Password expiration: $expiryText"
        ) -join [Environment]::NewLine
    } catch {
        $resultBox.Text = $_.Exception.Message
    } finally {
        $window.Cursor = $null
    }
})

$window.Add_ContentRendered({ [void]$upnBox.Focus() })
[void]$window.ShowDialog()
