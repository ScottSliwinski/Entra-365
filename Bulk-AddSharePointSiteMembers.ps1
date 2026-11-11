# CSV header: UPN
# Run in PowerShell after installing Microsoft.Online.SharePoint.PowerShell.
param(
    [string]$SiteUrl,
    [string]$CsvPath,
    [string]$MemberGroup,
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable Microsoft.Online.SharePoint.PowerShell)) {
    throw 'Install the SharePoint Online module first: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser'
}
Import-Module Microsoft.Online.SharePoint.PowerShell

if (-not $SiteUrl) { $SiteUrl = Read-Host 'SharePoint site URL (https://tenant.sharepoint.com/sites/name)' }
$siteUri = $null
if (-not [uri]::TryCreate($SiteUrl, [uriKind]::Absolute, [ref]$siteUri) -or
    $siteUri.Scheme -ne 'https' -or $siteUri.Host -notlike '*.sharepoint.com') {
    throw 'Enter a valid SharePoint Online site URL.'
}
if (-not $CsvPath) { $CsvPath = Read-Host 'Full path to CSV (header: UPN)' }
$CsvPath = (Resolve-Path -LiteralPath $CsvPath).Path
$rows = @(Import-Csv -LiteralPath $CsvPath)
if (-not $rows.Count -or 'UPN' -notin $rows[0].PSObject.Properties.Name) {
    throw 'CSV must contain at least one row and a UPN column.'
}

$adminUrl = "https://$($siteUri.Host.Split('.')[0])-admin.sharepoint.com"
Connect-SPOService -Url $adminUrl

if (-not $MemberGroup) {
    Write-Host 'Groups on this site:'
    Get-SPOSiteGroup -Site $SiteUrl | Select-Object Title | Format-Table -AutoSize
    $MemberGroup = Read-Host 'Enter the exact site Members group name'
}
$groups = @(Get-SPOSiteGroup -Site $SiteUrl)
if ($MemberGroup -notin @($groups | ForEach-Object Title)) {
    throw "Group '$MemberGroup' was not found on this site."
}

if (-not $ReportPath) {
    $ReportPath = Join-Path (Split-Path -Parent $CsvPath) ("SharePoint-Member-Results-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
$results = [System.Collections.Generic.List[object]]::new()
$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
for ($i = 0; $i -lt $rows.Count; $i++) {
    $upn = ([string]$rows[$i].UPN).Trim()
    Write-Progress -Activity 'Adding SharePoint site members' -Status "$($i + 1) of $($rows.Count): $upn" -PercentComplete ([int](100 * ($i + 1) / $rows.Count))
    $status = 'Added'; $detail = ''
    if ([string]::IsNullOrWhiteSpace($upn)) {
        $status = 'Skipped'; $detail = 'Blank UPN'
    } elseif (-not $seen.Add($upn)) {
        $status = 'Skipped'; $detail = 'Duplicate UPN in CSV'
    } else {
        try {
            Add-SPOUser -Site $SiteUrl -Group $MemberGroup -LoginName $upn -ErrorAction Stop
        } catch {
            $status = 'Failed'; $detail = $_.Exception.Message
        }
    }
    Write-Host "$upn : $status"
    $results.Add([pscustomobject]@{ UPN = $upn; Status = $status; Detail = $detail; Site = $SiteUrl; Group = $MemberGroup })
}
Write-Progress -Activity 'Adding SharePoint site members' -Completed
$results | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "Results saved to: $ReportPath"
