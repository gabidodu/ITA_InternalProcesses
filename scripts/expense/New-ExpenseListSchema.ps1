<#
.SYNOPSIS
    Creates the 5 SharePoint lists this app needs (DE_ExpenseReports,
    DE_ExpenseLines, DE_ExpenseCategory, DE_Setup, DE_AuditLog) on a
    destination site, using an Entra ID app-only registration (certificate
    auth) via PnP.PowerShell.

.DESCRIPTION
    Unlike the sibling LeaveRequestForTeams app's Copy-SharePointListSchema.ps1
    (which copies schema FROM a legacy source site), there is no legacy
    "expense reports" site to clone — the schema below is authored directly
    in this script as the source of truth. Fields are added one at a time via
    Add-PnPField (not the PnP site-template engine), following the same
    already-validated approach and the same reason: Get-PnPSiteTemplate /
    Invoke-PnPSiteTemplate exhibited a reproducible bug on this tenant where
    fields from one list ended up attached to a different, previously
    processed list.

    Safe to re-run: existing lists/fields are left alone, only missing ones
    are added.

.PARAMETER DestinationSiteUrl
    URL of the SharePoint site where the lists should be created (e.g.
    https://itassist2.sharepoint.com/sites/AVI_Site, or a different site/
    tenant entirely for a later client deployment).

.PARAMETER ClientId
    Entra ID application (client) ID used for authentication. Defaults to the
    shared AccessLeaveRequest app registration this project reuses.

.PARAMETER Thumbprint
    Thumbprint of the certificate (installed in the local CurrentUser
    certificate store) whose public key has been uploaded to the app
    registration.

.PARAMETER Tenant
    Tenant's initial/verified domain, e.g. "contoso.onmicrosoft.com".

.EXAMPLE
    ./New-ExpenseListSchema.ps1 `
        -DestinationSiteUrl "https://itassist2.sharepoint.com/sites/AVI_Site" `
        -ClientId "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "itassist2.onmicrosoft.com"

.NOTES
    Required app permissions: the app registration needs Application
    permission Sites.Manage.All (or Sites.FullControl.All) on SharePoint with
    admin consent, OR Sites.Selected with explicit
    Grant-PnPAzureADAppSitePermission on the destination site (see
    Grant-ExpenseSitePermission.ps1).
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
    [string]$Tenant
)

$ErrorActionPreference = "Stop"

# Field shape mirrors Copy-SharePointListSchema.ps1's captured schema JSON:
# Title / InternalName / Type / Required / DefaultValue.
$script:ListSchemas = [ordered]@{
    "DE_ExpenseReports" = @(
        @{ Title = "EmployeeCode"; InternalName = "EmployeeCode"; Type = "Text"; Required = $true }
        @{ Title = "RequestSubject"; InternalName = "RequestSubject"; Type = "Text"; Required = $true }
        @{ Title = "RequestDate"; InternalName = "RequestDate"; Type = "DateTime"; Required = $false }
        @{ Title = "TotalAmount"; InternalName = "TotalAmount"; Type = "Number"; Required = $false }
        @{ Title = "Currency"; InternalName = "Currency"; Type = "Text"; Required = $false; DefaultValue = "RON" }
        @{ Title = "Status"; InternalName = "Status"; Type = "Text"; Required = $false; DefaultValue = "PENDING" }
        @{ Title = "CurrentUser"; InternalName = "CurrentUser"; Type = "Text"; Required = $false }
        @{ Title = "ApprovalFlow"; InternalName = "ApprovalFlow"; Type = "Text"; Required = $false }
        @{ Title = "Comments"; InternalName = "Comments"; Type = "Text"; Required = $false }
    )
    "DE_ExpenseLines" = @(
        @{ Title = "ExpenseReportId"; InternalName = "ExpenseReportId"; Type = "Number"; Required = $true }
        @{ Title = "LineNumber"; InternalName = "LineNumber"; Type = "Number"; Required = $false }
        @{ Title = "ExpenseDate"; InternalName = "ExpenseDate"; Type = "DateTime"; Required = $false }
        @{ Title = "CategoryId"; InternalName = "CategoryId"; Type = "Number"; Required = $false }
        @{ Title = "Vendor"; InternalName = "Vendor"; Type = "Text"; Required = $false }
        @{ Title = "Description"; InternalName = "Description"; Type = "Note"; Required = $false }
        @{ Title = "Amount"; InternalName = "Amount"; Type = "Number"; Required = $false }
        @{ Title = "Currency"; InternalName = "Currency"; Type = "Text"; Required = $false }
        @{ Title = "ExchangeRate"; InternalName = "ExchangeRate"; Type = "Number"; Required = $false }
    )
    "DE_ExpenseCategory" = @(
        @{ Title = "CategoryName"; InternalName = "CategoryName"; Type = "Text"; Required = $true }
        @{ Title = "CategoryNameRO"; InternalName = "CategoryNameRO"; Type = "Text"; Required = $false }
        @{ Title = "IDCategory"; InternalName = "IDCategory"; Type = "Number"; Required = $true }
        @{ Title = "CategoryCode"; InternalName = "CategoryCode"; Type = "Text"; Required = $false }
        @{ Title = "CategoryDescription"; InternalName = "CategoryDescription"; Type = "Note"; Required = $false }
    )
    "DE_Setup" = @(
        @{ Title = "SettingKey"; InternalName = "SettingKey"; Type = "Text"; Required = $true }
        @{ Title = "SettingValue"; InternalName = "SettingValue"; Type = "Text"; Required = $false }
        @{ Title = "Category"; InternalName = "Category"; Type = "Text"; Required = $true }
    )
    "DE_AuditLog" = @(
        @{ Title = "Timestamp"; InternalName = "Timestamp"; Type = "DateTime"; Required = $false }
        @{ Title = "Actor"; InternalName = "Actor"; Type = "Text"; Required = $false }
        @{ Title = "Action"; InternalName = "Action"; Type = "Text"; Required = $false }
        @{ Title = "EmployeeCode"; InternalName = "EmployeeCode"; Type = "Text"; Required = $false }
        @{ Title = "RequestId"; InternalName = "RequestId"; Type = "Text"; Required = $false }
        @{ Title = "Details"; InternalName = "Details"; Type = "Note"; Required = $false }
    )
}

Write-Host "Connecting to destination site: $DestinationSiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

$script:CreatedListIds = @{}

foreach ($listName in $script:ListSchemas.Keys) {
    $list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "Creating list: $listName" -ForegroundColor Cyan
        $list = New-PnPList -Title $listName -Template GenericList -OnQuickLaunch
    } else {
        Write-Host "List already exists, adding missing fields: $listName" -ForegroundColor Yellow
    }

    $existingInternalNames = (Get-PnPField -List $listName).InternalName

    foreach ($f in $script:ListSchemas[$listName]) {
        if ($existingInternalNames -contains $f.InternalName) {
            Write-Host "  Skipping (already exists): $($f.InternalName)" -ForegroundColor DarkGray
            continue
        }

        Write-Host "  Adding field: $($f.InternalName) ($($f.Type))"
        Add-PnPField -List $listName -DisplayName $f.Title -InternalName $f.InternalName `
            -Type $f.Type -Required:$f.Required -AddToDefaultView | Out-Null

        if ($f.DefaultValue) {
            Set-PnPField -List $listName -Identity $f.InternalName -Values @{ DefaultValue = $f.DefaultValue } | Out-Null
        }
    }

    $refreshedList = Get-PnPList -Identity $listName
    $script:CreatedListIds[$listName] = $refreshedList.Id.Guid
}

Write-Host ""
Write-Host "Done. List IDs (paste into api/local.settings.json / SWA app settings):" -ForegroundColor Green
foreach ($listName in $script:CreatedListIds.Keys) {
    Write-Host "  $listName -> $($script:CreatedListIds[$listName])"
}
Write-Host "Verify the lists in $DestinationSiteUrl before relying on them." -ForegroundColor Yellow
