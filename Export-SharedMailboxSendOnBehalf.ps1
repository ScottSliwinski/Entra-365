#requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
Exports Send on Behalf delegates on shared mailboxes to a CSV.
.EXAMPLE
.\Export-SharedMailboxSendOnBehalf.ps1
.EXAMPLE
.\Export-SharedMailboxSendOnBehalf.ps1 -OutputPath C:\Reports\SharedMailbox-SendOnBehalf.csv
#>
param(
    [string]$OutputPath = (Join-Path (Get-Location) 'SharedMailbox-SendOnBehalf.csv')
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
    throw 'Install the ExchangeOnlineManagement module first: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' })) {
    Connect-ExchangeOnline -ShowBanner:$false
}

$folder = Split-Path -Parent $OutputPath
if ($folder -and -not (Test-Path -LiteralPath $folder)) {
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
}

Write-Host 'Loading shared mailboxes...'
$mailboxes = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited)
$rows = [System.Collections.Generic.List[object]]::new()
$index = 0

foreach ($mailbox in $mailboxes) {
    $index++
    Write-Progress -Activity 'Checking Send on Behalf delegates' -Status "$index of $($mailboxes.Count): $($mailbox.DisplayName)" -PercentComplete ([int](100 * $index / [Math]::Max(1, $mailboxes.Count)))

    foreach ($delegate in $mailbox.GrantSendOnBehalfTo) {
        $delegateIdentity = [string]$delegate
        $recipient = $null
        try {
            $recipient = Get-Recipient -Identity $delegateIdentity -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not resolve delegate '$delegateIdentity' on '$($mailbox.PrimarySmtpAddress)': $($_.Exception.Message)"
        }

        $rows.Add([pscustomobject]@{
            SharedMailboxName  = $mailbox.DisplayName
            SharedMailboxEmail = [string]$mailbox.PrimarySmtpAddress
            DelegateName       = if ($recipient) { $recipient.DisplayName } else { $delegateIdentity }
            DelegateUPN        = if ($recipient) { [string]$recipient.UserPrincipalName } else { '' }
            DelegateEmail      = if ($recipient) { [string]$recipient.PrimarySmtpAddress } else { '' }
            DelegateIdentity   = $delegateIdentity
            Permission         = 'Send on Behalf'
        })
    }
}

Write-Progress -Activity 'Checking Send on Behalf delegates' -Completed

if ($rows.Count -eq 0) {
    Write-Host 'No Send on Behalf delegates were found on shared mailboxes. No CSV was created.'
    return
}

$rows | Sort-Object SharedMailboxEmail, DelegateName | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported $($rows.Count) delegate assignment(s) across $(($rows | Select-Object -ExpandProperty SharedMailboxEmail -Unique).Count) shared mailbox(es) to: $OutputPath"
