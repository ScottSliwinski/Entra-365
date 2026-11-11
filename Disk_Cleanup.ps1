#requires -RunAsAdministrator

<#
.SYNOPSIS
    Performs an aggressive Windows disk cleanup.

.DESCRIPTION
    Cleans:
      - Windows temporary files
      - User temporary files
      - Windows Update download cache
      - Delivery Optimization cache
      - Recycle Bin
      - Previous Windows installation files (Windows.old) via CleanMgr
      - Superseded Windows component files
      - Windows upgrade/setup leftovers

    Run PowerShell as Administrator.

.WARNING
    DISM /ResetBase prevents currently superseded Windows updates
    from being uninstalled afterward.
#>

$ErrorActionPreference = "SilentlyContinue"

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "      WINDOWS DISK CLEANUP SCRIPT" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

# ------------------------------------------------------------
# Verify Administrator Rights
# ------------------------------------------------------------

$CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()

$Principal = New-Object Security.Principal.WindowsPrincipal($CurrentUser)

if (-not $Principal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit 1
}


# ------------------------------------------------------------
# Free Space BEFORE Cleanup
# ------------------------------------------------------------

$DriveBefore = Get-CimInstance Win32_LogicalDisk `
    -Filter "DeviceID='C:'"

$FreeBefore = $DriveBefore.FreeSpace

Write-Host "Free space before cleanup:" -ForegroundColor Yellow

"{0:N2} GB" -f ($FreeBefore / 1GB)

Write-Host ""


# ------------------------------------------------------------
# Stop Windows Update Services
# ------------------------------------------------------------

Write-Host "[1/8] Stopping Windows Update services..." -ForegroundColor Cyan

Stop-Service wuauserv -Force
Stop-Service bits -Force
Stop-Service dosvc -Force


# ------------------------------------------------------------
# Clean Windows Temp
# ------------------------------------------------------------

Write-Host "[2/8] Cleaning Windows temporary files..." -ForegroundColor Cyan

$WindowsTemp = "$env:SystemRoot\Temp"

if (Test-Path $WindowsTemp) {

    Get-ChildItem $WindowsTemp -Force |
        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
}


# ------------------------------------------------------------
# Clean Current User Temp
# ------------------------------------------------------------

Write-Host "[3/8] Cleaning user temporary files..." -ForegroundColor Cyan

if (Test-Path $env:TEMP) {

    Get-ChildItem $env:TEMP -Force |
        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
}


# ------------------------------------------------------------
# Clean Windows Update Download Cache
# ------------------------------------------------------------

Write-Host "[4/8] Cleaning Windows Update download cache..." -ForegroundColor Cyan

$UpdateCache = "$env:SystemRoot\SoftwareDistribution\Download"

if (Test-Path $UpdateCache) {

    Get-ChildItem $UpdateCache -Force |
        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
}


# ------------------------------------------------------------
# Clean Delivery Optimization Cache
# ------------------------------------------------------------

Write-Host "[5/8] Cleaning Delivery Optimization cache..." -ForegroundColor Cyan

try {

    Delete-DeliveryOptimizationCache -Force

}
catch {

    Write-Host "Delivery Optimization cleanup skipped." -ForegroundColor DarkYellow
}


# ------------------------------------------------------------
# Empty Recycle Bin
# ------------------------------------------------------------

Write-Host "[6/8] Emptying Recycle Bin..." -ForegroundColor Cyan

Clear-RecycleBin -Force -ErrorAction SilentlyContinue


# ------------------------------------------------------------
# Windows Disk Cleanup
# ------------------------------------------------------------

Write-Host "[7/8] Running Windows Disk Cleanup..." -ForegroundColor Cyan

#
# /AUTOCLEAN specifically cleans files left behind
# after Windows upgrades, including previous Windows
# installation files when applicable.
#

Start-Process cleanmgr.exe `
    -ArgumentList "/AUTOCLEAN" `
    -Wait `
    -WindowStyle Hidden


# ------------------------------------------------------------
# DISM Component Store Cleanup
# ------------------------------------------------------------

Write-Host "[8/8] Cleaning Windows component store..." -ForegroundColor Cyan

Write-Host ""
Write-Host "Removing superseded Windows components..." -ForegroundColor Yellow

DISM.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase


# ------------------------------------------------------------
# Additional Windows Setup Cleanup
# ------------------------------------------------------------

$CleanupFolders = @(

    "C:\`$WINDOWS.~BT",
    "C:\`$WINDOWS.~WS",
    "C:\Windows.old"

)

foreach ($Folder in $CleanupFolders) {

    if (Test-Path $Folder) {

        Write-Host "Found Windows upgrade folder: $Folder" -ForegroundColor Yellow

        try {

            takeown.exe /F "$Folder" /R /D Y | Out-Null

            icacls.exe "$Folder" `
                /grant Administrators:F `
                /T `
                /C `
                /Q | Out-Null

            Remove-Item `
                -Path $Folder `
                -Recurse `
                -Force `
                -ErrorAction SilentlyContinue

        }
        catch {

            Write-Host "Unable to completely remove $Folder" `
                -ForegroundColor DarkYellow
        }
    }
}


# ------------------------------------------------------------
# Restart Windows Update Services
# ------------------------------------------------------------

Write-Host ""
Write-Host "Restarting Windows Update services..." -ForegroundColor Cyan

Start-Service bits
Start-Service wuauserv
Start-Service dosvc


# ------------------------------------------------------------
# Free Space AFTER Cleanup
# ------------------------------------------------------------

Start-Sleep -Seconds 2

$DriveAfter = Get-CimInstance Win32_LogicalDisk `
    -Filter "DeviceID='C:'"

$FreeAfter = $DriveAfter.FreeSpace

$Recovered = $FreeAfter - $FreeBefore

Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
Write-Host "             CLEANUP COMPLETE" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Green
Write-Host ""

Write-Host "Free space before:" -ForegroundColor White -NoNewline
Write-Host (" {0:N2} GB" -f ($FreeBefore / 1GB)) -ForegroundColor Yellow

Write-Host "Free space after: " -ForegroundColor White -NoNewline
Write-Host (" {0:N2} GB" -f ($FreeAfter / 1GB)) -ForegroundColor Green

Write-Host "Space recovered:  " -ForegroundColor White -NoNewline
Write-Host (" {0:N2} GB" -f ($Recovered / 1GB)) -ForegroundColor Green

Write-Host ""