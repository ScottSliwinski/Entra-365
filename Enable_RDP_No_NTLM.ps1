#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'

try {
    $terminalServer = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
    $rdpListener    = "$terminalServer\WinStations\RDP-Tcp"
    $policyPath     = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'

    if (-not (Test-Path $rdpListener)) {
        throw 'The RDP listener was not found. Verify this Windows edition supports hosting Remote Desktop.'
    }

    # Enable incoming Remote Desktop connections.
    New-ItemProperty -Path $terminalServer `
        -Name 'fDenyTSConnections' -Value 0 `
        -PropertyType DWord -Force | Out-Null

    # Disable Network Level Authentication.
    New-ItemProperty -Path $rdpListener `
        -Name 'UserAuthentication' -Value 0 `
        -PropertyType DWord -Force | Out-Null

    # Update matching local policy registry settings.
    # Domain Group Policy or device management can overwrite these.
    New-Item -Path $policyPath -Force | Out-Null

    New-ItemProperty -Path $policyPath `
        -Name 'fDenyTSConnections' -Value 0 `
        -PropertyType DWord -Force | Out-Null

    New-ItemProperty -Path $policyPath `
        -Name 'UserAuthentication' -Value 0 `
        -PropertyType DWord -Force | Out-Null

    # Enable built-in RDP rules using language-independent rule names.
    $rules = @(Get-NetFirewallRule -Name 'RemoteDesktop-UserMode-In-*')

    if ($rules.Count -eq 0) {
        throw 'No built-in Remote Desktop firewall rules were found.'
    }

    $rules | Enable-NetFirewallRule

    # Start the service without restarting active RDP sessions.
    Set-Service -Name 'TermService' -StartupType Automatic
    Start-Service -Name 'TermService'

    Write-Host "`nRDP configuration completed." -ForegroundColor Green

    [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        RDPEnabled   = (
            (Get-ItemProperty $terminalServer).fDenyTSConnections -eq 0
        )
        NLARequired  = (
            (Get-ItemProperty $rdpListener).UserAuthentication -eq 1
        )
        ServiceState = (Get-Service TermService).Status
    } | Format-List
}
catch {
    Write-Error "RDP configuration failed: $($_.Exception.Message)"
}