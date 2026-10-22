<#
.SYNOPSIS
    Seeds DE_ExpenseCategory with the default expense-category catalog on a
    destination site, using an Entra ID app-only registration (certificate
    auth) via PnP.PowerShell.

.DESCRIPTION
    There's no legacy site to clone this data from (unlike, e.g., the sibling
    LeaveRequestForTeams app's LR_ActionType, which was cloned row-for-row
    from a real legacy list) — these are a sensible default set, editable
    afterwards directly in SharePoint like any other catalog list in this
    project. Safe to re-run: rows are matched by IDCategory and skipped if
    already present, never duplicated.

.PARAMETER DestinationSiteUrl
    URL of the SharePoint site whose DE_ExpenseCategory list should be seeded.
    Must already exist (run New-ExpenseListSchema.ps1 first).

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
    ./Copy-ExpenseCategorySeedData.ps1 `
        -DestinationSiteUrl "https://itassist2.sharepoint.com/sites/AVI_Site" `
        -ClientId "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "itassist2.onmicrosoft.com"
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

$script:ListName = "DE_ExpenseCategory"
$script:DefaultCategories = @(
    @{ Id = 1; Name = "Transport"; NameRO = "Transport"; Code = "TRA" }
    @{ Id = 2; Name = "Accommodation"; NameRO = "Cazare"; Code = "CAZ" }
    @{ Id = 3; Name = "Meals"; NameRO = "Masă"; Code = "MAS" }
    @{ Id = 4; Name = "Fuel"; NameRO = "Combustibil"; Code = "COMB" }
    @{ Id = 5; Name = "Supplies"; NameRO = "Consumabile"; Code = "CONS" }
    @{ Id = 6; Name = "Other"; NameRO = "Diverse"; Code = "DIV" }
    @{ Id = 7; Name = "Cash Advance"; NameRO = "Avans numerar"; Code = "ADV" }
)

Write-Host "Connecting to destination site: $DestinationSiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

$existingItems = Get-PnPListItem -List $script:ListName -Fields "IDCategory"
$existingIds = $existingItems | ForEach-Object { [int]$_["IDCategory"] }

foreach ($category in $script:DefaultCategories) {
    if ($existingIds -contains $category.Id) {
        Write-Host "Skipping (already seeded): $($category.NameRO) (IDCategory=$($category.Id))" -ForegroundColor DarkGray
        continue
    }

    Write-Host "Adding category: $($category.NameRO) (IDCategory=$($category.Id))" -ForegroundColor Cyan
    Add-PnPListItem -List $script:ListName -Values @{
        Title               = $category.NameRO
        CategoryName        = $category.Name
        CategoryNameRO      = $category.NameRO
        IDCategory          = $category.Id
        CategoryCode        = $category.Code
        CategoryDescription = ""
    } | Out-Null
}

Write-Host "Done. Verify $script:ListName in $DestinationSiteUrl before relying on it." -ForegroundColor Green
