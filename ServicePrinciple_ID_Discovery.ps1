Connect-MgGraph -Scopes "Application.Read.All"

$Id = Read-Host "Enter the Service Principal Object ID or Application ID"

# First, try it as the service principal Object ID
$ServicePrincipal = Get-MgServicePrincipal `
    -ServicePrincipalId $Id `
    -Property Id,AppId,DisplayName,ServicePrincipalType,AccountEnabled `
    -ErrorAction SilentlyContinue

# If not found, try it as the Application (client) ID
if (-not $ServicePrincipal) {
    $ServicePrincipal = Get-MgServicePrincipal `
        -Filter "appId eq '$Id'" `
        -Property Id,AppId,DisplayName,ServicePrincipalType,AccountEnabled `
        -ErrorAction SilentlyContinue
}

if ($ServicePrincipal) {
    $ServicePrincipal | Select-Object `
        DisplayName,
        Id,
        AppId,
        ServicePrincipalType,
        AccountEnabled | Format-List
}
else {
    Write-Warning "No service principal was found using ID: $Id"
}