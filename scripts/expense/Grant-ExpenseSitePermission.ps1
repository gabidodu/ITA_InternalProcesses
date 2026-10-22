<#
.SYNOPSIS
    Grants the app registration a Sites.Selected "write" permission on a
    destination site, so its delegated (Teams SSO + On-Behalf-Of) calls can
    read/write this app's SharePoint lists and document library there.

.DESCRIPTION
    Thin wrapper over Grant-PnPAzureADAppSitePermission. Needed once per
    destination site collection — this project's app registration
    (AccessLeaveRequest) already has broad tenant-wide Application
    permissions (Sites.FullControl.All etc.), which is what lets the
    connection below succeed even on a site it doesn't yet have a
    Sites.Selected grant on; but the deployed app itself calls Graph with a
    per-user DELEGATED token requesting Sites.Selected, which independently
    requires this per-site grant to actually resolve to that site. Skip this
    entirely if the destination is AVI_Site itself — it already has this
    grant from the sibling LeaveRequestForTeams app's setup.

.PARAMETER DestinationSiteUrl
    URL of the SharePoint site to grant permission on.

.PARAMETER ClientId
    Entra ID application (client) ID to grant. Defaults to the shared
    AccessLeaveRequest app registration this project reuses.

.PARAMETER AppDisplayName
    Display name recorded alongside the grant (cosmetic, shown in the site's
    permission listing). Defaults to "AccessLeaveRequest".

.PARAMETER Thumbprint
    Thumbprint of the certificate (installed in the local CurrentUser
    certificate store) whose public key has been uploaded to the app
    registration.

.PARAMETER Tenant
    Tenant's initial/verified domain, e.g. "contoso.onmicrosoft.com". For a
    genuinely different CLIENT tenant (not just a different site in the same
    tenant), a Global/Application admin in that tenant must first grant
    tenant-wide admin consent to this app registration — see the project
    plan's cross-tenant deployment note.

.EXAMPLE
    ./Grant-ExpenseSitePermission.ps1 `
        -DestinationSiteUrl "https://contoso.sharepoint.com/sites/Expenses" `
        -ClientId "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "contoso.onmicrosoft.com"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DestinationSiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "359906f4-0bbf-4810-97f4-0d1b28aa5f1a",

    [Parameter(Mandatory = $false)]
    [string]$AppDisplayName = "AccessLeaveRequest",

    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $true)]
    [string]$Tenant
)

$ErrorActionPreference = "Stop"

Write-Host "Connecting to destination site: $DestinationSiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

Write-Host "Granting Sites.Selected 'write' to $AppDisplayName ($ClientId) on $DestinationSiteUrl" -ForegroundColor Cyan
Grant-PnPAzureADAppSitePermission -AppId $ClientId -DisplayName $AppDisplayName -Site $DestinationSiteUrl -Permissions Write

Write-Host "Done. Verify via the site's API access settings, or a subsequent delegated Graph call, before relying on it." -ForegroundColor Green
