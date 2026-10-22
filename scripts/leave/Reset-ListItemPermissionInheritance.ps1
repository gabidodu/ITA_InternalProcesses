<#
.SYNOPSIS
    Resets item-level permission inheritance on every item of a SharePoint
    list, undoing the legacy app's per-item sharing (each employee shared
    only their own rows) so that access is governed solely by the list's own
    permissions (the "IT Assist" group Contribute grant) — the same list,
    visible the same way, to everyone who has access to the list itself.

.DESCRIPTION
    The legacy Power-Automate-driven app shared each LR_LeaveRequests item
    individually with the employee it concerned, which is why every employee
    ended up with an automatic "Limited Access" role assignment directly on
    the list (see Remove-IndividualListPermissions.ps1) — a side effect of
    unique, item-level permissions existing underneath. Removing the list-
    level artifact alone doesn't undo the real cause: the items themselves
    still have their own unique ACLs.

    This resets every item back to inheriting the list's own permissions
    (`ResetRoleInheritance`), which is safe/idempotent to call even on an item
    that's already inheriting. Processes in batches (one Invoke-PnPQuery per
    batch, not per item) so it doesn't take one round trip per row on a list
    with 1000+ items.

.PARAMETER SiteUrl
    URL of the SharePoint site.

.PARAMETER ListName
    Display name of the list to fix.

.PARAMETER ClientId
    Entra ID application (client) ID used for authentication.

.PARAMETER Thumbprint
    Thumbprint of the certificate (installed in the local CurrentUser
    certificate store) whose public key has been uploaded to the app
    registration.

.PARAMETER Tenant
    Tenant's initial/verified domain, e.g. "contoso.onmicrosoft.com".

.PARAMETER BatchSize
    How many items to reset per round trip. Defaults to 200.

.EXAMPLE
    ./Reset-ListItemPermissionInheritance.ps1 `
        -SiteUrl "https://itassist2.sharepoint.com/sites/LeaveRequest" `
        -ListName "LR_LeaveRequests" `
        -ClientId "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "itassist2.onmicrosoft.com"

.NOTES
    Run Remove-IndividualListPermissions.ps1 afterwards (or before — order
    doesn't matter) to also clean up the now-redundant list-level "Limited
    Access" entries this per-item sharing left behind.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ListName = "LR_LeaveRequests",

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $true)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 200
)

$ErrorActionPreference = "Stop"

Write-Host "Connecting to: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

Write-Host "Loading items from $ListName (this can take a while on a large list)..." -ForegroundColor Cyan
$items = Get-PnPListItem -List $ListName -PageSize 500
Write-Host "Loaded $($items.Count) items." -ForegroundColor Cyan

$total = $items.Count
$processed = 0
$batchCount = 0

for ($i = 0; $i -lt $total; $i += $BatchSize) {
    $batch = $items[$i..([Math]::Min($i + $BatchSize - 1, $total - 1))]
    foreach ($item in $batch) {
        $item.ResetRoleInheritance()
    }
    Invoke-PnPQuery
    $processed += $batch.Count
    $batchCount++
    Write-Host "  Batch $batchCount done — $processed / $total items reset." -ForegroundColor DarkGray
}

Disconnect-PnPOnline
Write-Host ""
Write-Host "Done. All $total items on '$ListName' now inherit permissions from the list — nobody has a leftover per-item grant anymore." -ForegroundColor Green
