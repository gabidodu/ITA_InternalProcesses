<#
.SYNOPSIS
    Turns on the expense module for a client tenant that's already registered
    (scripts/leave/Register-ClientTenant.ps1), by adding the DE_* list ids to
    its existing TenantConfigs row.

.DESCRIPTION
    Run AFTER Deploy-ExpenseInfrastructure.ps1 has created the DE_* lists and
    the DE_Attachments library on the client's site. Resolves their ids and
    MERGES them into the tenant's row — leave columns are left untouched.
    Once every expense column is present, the home screen shows the expense
    tile for that tenant (see api/src/shared/lib/tenantConfig.ts,
    getTenantModules); until then the tenant simply doesn't see it.

    The DE_* lists may live on a different site than the leave lists — pass
    that site as -SiteUrl and its Graph id is stored as expenseSiteId.
    Otherwise (same site, the IT ASSIST setup) expenseSiteId is still written,
    it's just equal to siteId.

    Requires PnP.PowerShell (certificate auth, same as the other scripts) and
    an `az` login that can read the storage account's connection string.

.EXAMPLE
    ./Enable-ExpenseModule.ps1 `
        -ClientTenantId "30b69b00-a156-460e-9654-1b28d0ef5522" `
        -SiteUrl "https://itassist2.sharepoint.com/sites/AVI_Site" `
        -ClientId "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" `
        -Thumbprint "<cert thumbprint>" `
        -Tenant "itassist2.onmicrosoft.com" `
        -WhatIf

.NOTES
    -WhatIf prints the columns it would write and stops before touching the
    table. Idempotent: re-running just rewrites the same columns.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientTenantId,

    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $true)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [string]$StorageAccountName = "stleaverequestforteams",

    [Parameter(Mandatory = $false)]
    [string]$StorageResourceGroup = "rg-internal-processes"
)

$ErrorActionPreference = "Stop"

# Keep in sync with EXPENSE_REQUIRED_FIELDS in api/src/shared/lib/tenantConfig.ts.
$script:ListNameToConfigField = [ordered]@{
    "DE_ExpenseReports"  = "expenseReportsListId"
    "DE_ExpenseLines"    = "expenseLinesListId"
    "DE_ExpenseCategory" = "expenseCategoryListId"
    "DE_Attachments"     = "expenseAttachmentsListId"
    "DE_Setup"           = "expenseSetupListId"
    "DE_AuditLog"        = "expenseAuditLogListId"
}

Write-Host "Connecting to site: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

$uri = [Uri]$SiteUrl
$siteInfo = Invoke-PnPGraphMethod -Method GET -Url "sites/$($uri.Host):$($uri.AbsolutePath.TrimEnd('/'))"
if (-not $siteInfo.id) {
    throw "Could not resolve a Graph site id for $SiteUrl."
}

$configFields = [ordered]@{ "expenseSiteId" = $siteInfo.id }
foreach ($name in $script:ListNameToConfigField.Keys) {
    $list = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
    if (-not $list) {
        throw "List '$name' was not found on $SiteUrl. Run Deploy-ExpenseInfrastructure.ps1 first."
    }
    $configFields[$script:ListNameToConfigField[$name]] = $list.Id.Guid
}
Disconnect-PnPOnline

$configFields.GetEnumerator() | ForEach-Object { Write-Host ("  {0,-26} {1}" -f $_.Key, $_.Value) -ForegroundColor DarkGray }

if (-not $PSCmdlet.ShouldProcess("TenantConfigs row $ClientTenantId", "Merge expense columns")) {
    return
}

$connectionString = az storage account show-connection-string `
    --name $StorageAccountName `
    --resource-group $StorageResourceGroup `
    --query connectionString -o tsv
if (-not $connectionString) {
    throw "Could not obtain a connection string for storage account $StorageAccountName — check `az` login/permissions."
}

$entityArgs = @("PartitionKey=$ClientTenantId", "RowKey=$ClientTenantId")
foreach ($key in $configFields.Keys) {
    $entityArgs += "$key=$($configFields[$key])"
}

# merge = add/overwrite only these columns, keep the leave ones.
az storage entity merge --connection-string $connectionString --table-name TenantConfigs --entity $entityArgs | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Writing the TenantConfigs row failed (az exit code $LASTEXITCODE)."
}

Write-Host "Done. Expense module enabled for tenant $ClientTenantId." -ForegroundColor Green
