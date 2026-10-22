<#
.SYNOPSIS
    Orchestrates a full SharePoint infrastructure deploy for this app onto a
    destination site (same tenant or a different one entirely) in one call:
    grants site permission, creates the 5 lists, creates the attachments
    document library, and seeds the default expense-category catalog.

.DESCRIPTION
    Chains, in order: Grant-ExpenseSitePermission.ps1 (skippable — pass
    -SkipGrantPermission if the destination site already has this grant, e.g.
    AVI_Site via the sibling LeaveRequestForTeams app) -> New-ExpenseListSchema.ps1
    -> New-ExpenseAttachmentLibrary.ps1 -> Copy-ExpenseCategorySeedData.ps1.

    Every step is independently safe to re-run (existing lists/fields/rows
    are left alone), so this whole script is safe to re-run too if it fails
    partway through — just fix the underlying issue and run it again.

.PARAMETER DestinationSiteUrl
    URL of the SharePoint site to provision (existing site — this script does
    not create the site collection itself).

.PARAMETER ClientId
    Entra ID application (client) ID used for authentication. Defaults to the
    shared AccessLeaveRequest app registration this project reuses.

.PARAMETER Thumbprint
    Thumbprint of the certificate (installed in the local CurrentUser
    certificate store) whose public key has been uploaded to the app
    registration.

.PARAMETER Tenant
    Tenant's initial/verified domain, e.g. "contoso.onmicrosoft.com".

.PARAMETER SkipGrantPermission
    Skip the Sites.Selected permission grant step — use this when the
    destination site already has it (e.g. AVI_Site, already granted for the
    sibling LeaveRequestForTeams app).

.EXAMPLE
    First deploy, onto AVI_Site (already has the site permission grant):
    ./Deploy-ExpenseInfrastructure.ps1 `
        -DestinationSiteUrl "https://itassist2.sharepoint.com/sites/AVI_Site" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "itassist2.onmicrosoft.com" `
        -SkipGrantPermission

.EXAMPLE
    A brand-new client site/tenant (needs the permission grant — and, for a
    genuinely different tenant, that tenant's admin consent must already be
    granted for this app registration; see the project plan):
    ./Deploy-ExpenseInfrastructure.ps1 `
        -DestinationSiteUrl "https://contoso.sharepoint.com/sites/Expenses" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "contoso.onmicrosoft.com"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DestinationSiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ClientId = "359906f4-0bbf-4810-97f4-0d1b28aa5f1a",

    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $true)]
    [string]$Tenant,

    [switch]$SkipGrantPermission
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

if (-not $SkipGrantPermission) {
    Write-Host "=== Step 1/4: Grant site permission ===" -ForegroundColor Magenta
    & (Join-Path $scriptDir "Grant-ExpenseSitePermission.ps1") `
        -DestinationSiteUrl $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant
} else {
    Write-Host "=== Step 1/4: Grant site permission (skipped) ===" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "=== Step 2/4: Create lists ===" -ForegroundColor Magenta
& (Join-Path $scriptDir "New-ExpenseListSchema.ps1") `
    -DestinationSiteUrl $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

Write-Host ""
Write-Host "=== Step 3/4: Create attachments document library ===" -ForegroundColor Magenta
& (Join-Path $scriptDir "New-ExpenseAttachmentLibrary.ps1") `
    -DestinationSiteUrl $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

Write-Host ""
Write-Host "=== Step 4/4: Seed expense-category catalog ===" -ForegroundColor Magenta
& (Join-Path $scriptDir "Copy-ExpenseCategorySeedData.ps1") `
    -DestinationSiteUrl $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Green
Write-Host "Next steps:"
Write-Host "  1. Seed DE_Setup with at least one Admin row, one Finance row, the"
Write-Host "     allEmployeesGroupID row, and Setting/approvalFlow — no in-app UI for"
Write-Host "     this by design, add the rows directly in SharePoint (same convention"
Write-Host "     as the sibling LeaveRequestForTeams app's LR_Setup)."
Write-Host "  2. Paste the list IDs printed above into api/local.settings.json and,"
Write-Host "     once deployed, the Azure Static Web App's app settings."
Write-Host "  3. If this is a genuinely different SharePoint site ID than AVI_Site,"
Write-Host "     update SHAREPOINT_SITE_ID too."
