Connect-ExchangeOnline

# ==========================================
# SETTINGS
# ==========================================

$CalendarName = "Room Scheduling"

$Domains = @(
    "@awmcharis.com",
    "@awmi.net"
)

# ==========================================
# GET MATCHING MAILBOXES
# ==========================================

Write-Host "Getting matching mailboxes..." -ForegroundColor Cyan

$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox |
    Where-Object {
        $UPN = $_.UserPrincipalName

        $UPN -and (
            $UPN.EndsWith("@awmcharis.com", [System.StringComparison]::OrdinalIgnoreCase) -or
            $UPN.EndsWith("@awmi.net", [System.StringComparison]::OrdinalIgnoreCase)
        )
    }

Write-Host ""
Write-Host "Mailboxes to search: $($Mailboxes.Count)" -ForegroundColor Yellow
Write-Host ""

# ==========================================
# SEARCH CALENDARS
# ==========================================

$Results = @()

$Count = 0
$Total = $Mailboxes.Count

foreach ($Mailbox in $Mailboxes) {

    $Count++

    Write-Progress `
        -Activity "Searching Office 365 secondary calendars" `
        -Status "[$Count of $Total] $($Mailbox.UserPrincipalName)" `
        -PercentComplete (($Count / $Total) * 100)

    try {

        $Folders = Get-MailboxFolderStatistics `
            -Identity $Mailbox.PrimarySmtpAddress `
            -FolderScope Calendar `
            -ErrorAction Stop

        foreach ($Folder in $Folders) {

            # Partial-name search
            if ($Folder.Name -like "*$CalendarName*") {

                $Result = [PSCustomObject]@{
                    OwnerDisplayName = $Mailbox.DisplayName
                    OwnerUPN         = $Mailbox.UserPrincipalName
                    OwnerEmail       = $Mailbox.PrimarySmtpAddress
                    CalendarName     = $Folder.Name
                    FolderPath       = $Folder.FolderPath
                    FolderType       = $Folder.FolderType
                    ItemCount        = $Folder.ItemsInFolder
                }

                $Results += $Result

                Write-Host ""
                Write-Host "======================================" -ForegroundColor Green
                Write-Host " CALENDAR FOUND!" -ForegroundColor Green
                Write-Host "======================================" -ForegroundColor Green
                Write-Host "Owner:    $($Mailbox.DisplayName)"
                Write-Host "UPN:      $($Mailbox.UserPrincipalName)"
                Write-Host "Email:    $($Mailbox.PrimarySmtpAddress)"
                Write-Host "Calendar: $($Folder.Name)"
                Write-Host "Path:     $($Folder.FolderPath)"
                Write-Host "Items:    $($Folder.ItemsInFolder)"
                Write-Host "======================================" -ForegroundColor Green
                Write-Host ""
            }
        }
    }
    catch {

        Write-Warning "Unable to scan $($Mailbox.UserPrincipalName): $($_.Exception.Message)"

    }
}

Write-Progress -Activity "Searching Office 365 secondary calendars" -Completed

# ==========================================
# RESULTS
# ==========================================

Write-Host ""
Write-Host "SEARCH COMPLETE" -ForegroundColor Cyan
Write-Host ""

if ($Results.Count -gt 0) {

    $Results |
        Format-Table `
            OwnerDisplayName,
            OwnerUPN,
            CalendarName,
            FolderPath,
            ItemCount `
            -AutoSize

    $ExportPath = ".\CalendarOwnerSearch.csv"

    $Results |
        Export-Csv `
            -Path $ExportPath `
            -NoTypeInformation `
            -Encoding UTF8

    Write-Host ""
    Write-Host "Found $($Results.Count) matching calendar(s)." -ForegroundColor Green
    Write-Host "Results exported to: $ExportPath" -ForegroundColor Green

}
else {

    Write-Host "No calendars matching '$CalendarName' were found." -ForegroundColor Yellow

}

Disconnect-ExchangeOnline -Confirm:$false