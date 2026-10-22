<#
.SYNOPSIS
    Registers a new client tenant in the shared app's TenantConfigs table, so
    the multi-tenant deployment knows which SharePoint site/lists to use for
    that client's requests.

.DESCRIPTION
    Part of the per-client onboarding checklist (see the project plan / memory
    for the full sequence). Run this AFTER:
      1. The client's admin has granted tenant-wide admin consent for the
         shared Entra app, and granted it Sites.Selected write access on
         their chosen SharePoint site (Grant-PnPAzureADAppSitePermission,
         same mechanism already used for the existing tenant).
      2. Copy-SharePointListSchema.ps1 has been run against the client's site
         to clone the LR_* list schema.

    Resolves the client site's Graph-style site id (hostname,siteCollectionId,
    webId) and each LR_* list's id via Invoke-PnPGraphMethod (connected with
    this app's own certificate, against the CLIENT's tenant/site), then writes
    one row into the shared TenantConfigs Azure Table (PartitionKey = RowKey =
    the client's Entra tenant id) — this is the only step CreateLeaveRequest /
    RespondToLeaveRequest / etc. need to start correctly resolving that
    tenant's own SharePoint data at request time (see api/src/shared/lib/tenantConfig.ts).

    Run New-LeaveApprovalsLibrary.ps1 against the client's site BEFORE this
    script — LR_Approvals is in the default $ListNames set below, and this
    script throws if any listed list/library isn't found yet.

    Requires the Azure CLI (az) to be logged in with access to write to the
    TenantConfigs storage account — this script shells out to `az storage
    entity insert` rather than needing an extra PowerShell module, consistent
    with how this project already uses `az` for anything outside SharePoint
    itself.

.PARAMETER ClientTenantId
    The client's own Entra ID tenant id (GUID) — becomes the TenantConfigs
    row's PartitionKey/RowKey, and is what every request's OBO/Actionable
    Message token resolves to at runtime.

.PARAMETER SiteUrl
    URL of the client's SharePoint site holding the cloned LR_* lists.

.PARAMETER MailSenderAddress
    A real mailbox INSIDE the client's own tenant that the app has app-only
    Mail.Send + Calendars.ReadWrite access to — notification emails and the
    shared "Leave Requests" calendar for this client are sent/created from
    here. Must belong to the client's tenant, not IT ASSIST's — an app-only
    token minted for the client's tenant cannot send as a mailbox in a
    different tenant.

.PARAMETER ClientId
    This app's own Entra application (client) id — same one used everywhere
    else in this project (AAD_APP_CLIENT_ID).

.PARAMETER Thumbprint
    Thumbprint of the certificate (installed in the local CurrentUser
    certificate store) whose public key is uploaded to the app registration —
    same credential Copy-SharePointListSchema.ps1 uses.

.PARAMETER Tenant
    The CLIENT's own initial/verified domain (e.g. "contoso.onmicrosoft.com")
    — PnP connects to their site using their tenant context, not IT ASSIST's.

.PARAMETER ListNames
    Display names of the LR_* lists to resolve ids for. Defaults to the full
    standard set.

.PARAMETER StorageAccountName
    The shared TenantConfigs storage account. Defaults to the one this
    project already provisioned (stleaverequestforteams).

.PARAMETER StorageResourceGroup
    Resource group of the storage account above.

.EXAMPLE
    ./Register-ClientTenant.ps1 `
        -ClientTenantId "11111111-2222-3333-4444-555555555555" `
        -SiteUrl "https://contoso.sharepoint.com/sites/LeaveRequests" `
        -MailSenderAddress "leaverequests@contoso.com" `
        -ClientId "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "contoso.onmicrosoft.com"

.NOTES
    Re-running for the same ClientTenantId overwrites the existing row
    (idempotent) — safe to use to fix a mistake or update the mailbox/site
    without needing a separate "update" script.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientTenantId,

    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$MailSenderAddress,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $true)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [string[]]$ListNames = @(
        "LR_LeaveRequests", "LR_LeaveStatus", "LR_LeaveSetup",
        "LR_ActionType", "LR_Setup", "LR_AuditLog", "LR_HolidaySetup", "LR_Approvals"
    ),

    [Parameter(Mandatory = $false)]
    [string]$StorageAccountName = "stleaverequestforteams",

    [Parameter(Mandatory = $false)]
    [string]$StorageResourceGroup = "rg-internal-processes"
)

$ErrorActionPreference = "Stop"

# Maps each LR_* display name to the TenantConfig field it fills — keep in
# sync with api/src/shared/lib/tenantConfig.ts's TenantConfig type.
$script:ListNameToConfigField = @{
    "LR_LeaveRequests"  = "leaveRequestsListId"
    "LR_LeaveStatus"    = "leaveStatusListId"
    "LR_LeaveSetup"     = "leaveSetupListId"
    "LR_ActionType"     = "actionTypeListId"
    "LR_Setup"          = "setupListId"
    "LR_AuditLog"       = "auditLogListId"
    "LR_HolidaySetup"   = "holidaySetupListId"
    "LR_Approvals"      = "approvalsLibraryId"
}

Write-Host "Connecting to client site: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

$uri = [Uri]$SiteUrl
$hostName = $uri.Host
$sitePath = $uri.AbsolutePath.TrimEnd('/')
Write-Host "Resolving Graph site id for ${hostName}:${sitePath}" -ForegroundColor Cyan
$siteInfo = Invoke-PnPGraphMethod -Method GET -Url "sites/${hostName}:${sitePath}"
$siteId = $siteInfo.id
if (-not $siteId) {
    throw "Could not resolve a Graph site id for $SiteUrl — check the URL and that the app has been granted Sites.Selected on this site."
}
Write-Host "  siteId: $siteId" -ForegroundColor DarkGray

$configFields = @{}
foreach ($name in $ListNames) {
    $list = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
    if (-not $list) {
        throw "List '$name' was not found on $SiteUrl. Run Copy-SharePointListSchema.ps1 first."
    }
    $fieldName = $script:ListNameToConfigField[$name]
    if (-not $fieldName) {
        throw "No TenantConfig field mapping for list '$name' — add one to `$script:ListNameToConfigField above."
    }
    $configFields[$fieldName] = $list.Id.Guid
    Write-Host "  $name -> $fieldName = $($list.Id.Guid)" -ForegroundColor DarkGray
}

Disconnect-PnPOnline

$missingFields = @("leaveRequestsListId", "leaveStatusListId", "leaveSetupListId", "actionTypeListId", "setupListId", "auditLogListId", "holidaySetupListId", "approvalsLibraryId") |
    Where-Object { -not $configFields[$_] }
if ($missingFields) {
    throw "Missing TenantConfig fields (no matching list resolved): $($missingFields -join ', ')"
}

Write-Host "Writing TenantConfigs row for tenant $ClientTenantId..." -ForegroundColor Cyan
$connectionString = az storage account show-connection-string `
    --name $StorageAccountName `
    --resource-group $StorageResourceGroup `
    --query connectionString -o tsv
if (-not $connectionString) {
    throw "Could not obtain a connection string for storage account $StorageAccountName — check `az` login/permissions."
}

$entityArgs = @(
    "PartitionKey=$ClientTenantId",
    "RowKey=$ClientTenantId",
    "siteId=$siteId",
    "mailSenderAddress=$MailSenderAddress"
)
foreach ($key in $configFields.Keys) {
    $entityArgs += "$key=$($configFields[$key])"
}

# "merge", not "replace": the same row also carries other modules' columns
# (expense — see expense/Enable-ExpenseModule.ps1), which a replace would wipe.
az storage entity insert --connection-string $connectionString --table-name TenantConfigs --if-exists merge --entity $entityArgs | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Writing the TenantConfigs row failed (az exit code $LASTEXITCODE)."
}

Write-Host "Done. Tenant $ClientTenantId is now registered." -ForegroundColor Green
Write-Host "Next: have the client sideload/install the Teams app package — the first sign-in will resolve straight to this config." -ForegroundColor Yellow
