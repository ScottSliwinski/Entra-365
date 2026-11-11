#requires -Version 5.1
<#
.SYNOPSIS
Exports shared mailboxes with no Full Access, Send As, or Send on Behalf delegates.
.EXAMPLE
.\Export-SharedMailboxesWithoutDelegates.ps1
.EXAMPLE
.\Export-SharedMailboxesWithoutDelegates.ps1 -OutputPath C:\Reports\NoDelegates.csv
.NOTES
Install-Module ExchangeOnlineManagement -Scope CurrentUser (if needed).
Requires permission to read shared mailboxes and their delegation settings.
This checks mailbox-level delegation; it does not inspect folder permissions.
#>
param(
    [string]$OutputPath = (Join-Path (Get-Location).Path 'SharedMailboxes-WithoutDelegates.csv')
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw 'ExchangeOnlineManagement is missing. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
}
Import-Module ExchangeOnlineManagement

try {
    Connect-ExchangeOnline -ShowBanner:$false
    $mailboxes = @(Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties PrimarySmtpAddress,DisplayName,ExternalDirectoryObjectId)
    Write-Host "Found $($mailboxes.Count) shared mailboxes."
    $results = [System.Collections.Generic.List[object]]::new()
    $failures = [System.Collections.Generic.List[string]]::new()

    for ($i = 0; $i -lt $mailboxes.Count; $i++) {
        $mailbox = $mailboxes[$i]
        $identity = [string]$mailbox.PrimarySmtpAddress
        Write-Progress -Activity 'Checking shared mailbox delegation' -Status "$($i + 1) of $($mailboxes.Count): $identity" -PercentComplete ((($i + 1) / $mailboxes.Count) * 100)

        try {
            # Ignore inherited entries, Deny entries, and the mailbox's own SELF entry.
            $fullAccess = @(Get-EXOMailboxPermission -Identity $identity | Where-Object {
                -not $_.IsInherited -and -not $_.Deny -and
                $_.AccessRights -contains 'FullAccess' -and
                [string]$_.User -notmatch '^(NT AUTHORITY\\SELF|SELF)$'
            })
            $sendAs = @(Get-EXORecipientPermission -Identity $identity | Where-Object {
                -not $_.Deny -and $_.AccessRights -contains 'SendAs' -and
                [string]$_.Trustee -notmatch '^(NT AUTHORITY\\SELF|SELF)$'
            })
            # Read this property from Get-Mailbox; it is not a supported Get-EXOMailbox -Properties value.
            $details = Get-Mailbox -Identity $identity
            $sendOnBehalf = @($details.GrantSendOnBehalfTo | Where-Object { $null -ne $_ })

            if ($fullAccess.Count -eq 0 -and $sendAs.Count -eq 0 -and $sendOnBehalf.Count -eq 0) {
                $results.Add([pscustomobject]@{
                    DisplayName = $mailbox.DisplayName
                    PrimarySmtpAddress = $identity
                    ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                })
            }
        }
        catch {
            $failures.Add("${identity}: $($_.Exception.Message)")
            Write-Warning "Skipped $identity because a permission check failed: $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity 'Checking shared mailbox delegation' -Completed

    $parent = Split-Path -Parent $OutputPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    if ($results.Count -gt 0) {
        $results | Sort-Object DisplayName | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    }
    else {
        '"DisplayName","PrimarySmtpAddress","ExternalDirectoryObjectId"' | Set-Content -LiteralPath $OutputPath -Encoding UTF8
    }
    Write-Host "Exported $($results.Count) shared mailboxes without delegates to: $OutputPath"
    if ($failures.Count -gt 0) {
        Write-Warning "$($failures.Count) mailboxes could not be checked and were excluded from the CSV. Review the warnings above."
    }
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
