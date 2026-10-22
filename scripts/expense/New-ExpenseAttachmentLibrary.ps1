<#
.SYNOPSIS
    Creates the DE_Attachments document library (for receipt/invoice files)
    on a destination site, with its metadata columns, using an Entra ID
    app-only registration (certificate auth) via PnP.PowerShell.

.DESCRIPTION
    Receipts/invoices are stored as real files in a document library (not as
    classic SharePoint list-item Attachments) so each file can carry its own
    metadata — which expense report/line it belongs to, its document type,
    who uploaded it, and when — and be found independently via Microsoft
    Graph (see api/src/functions/UploadExpenseAttachment.ts /
    GetMyExpenseContext.ts). Safe to re-run: if the library or a field
    already exists, it's left alone.

.PARAMETER DestinationSiteUrl
    URL of the SharePoint site where the library should be created.

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
    ./New-ExpenseAttachmentLibrary.ps1 `
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

$script:LibraryName = "DE_Attachments"
$script:Fields = @(
    @{ Title = "ExpenseReportId"; InternalName = "ExpenseReportId"; Type = "Number"; Required = $true }
    @{ Title = "ExpenseLineId"; InternalName = "ExpenseLineId"; Type = "Number"; Required = $false }
    @{ Title = "DocumentType"; InternalName = "DocumentType"; Type = "Text"; Required = $false; DefaultValue = "Receipt" }
    @{ Title = "UploadedBy"; InternalName = "UploadedBy"; Type = "Text"; Required = $false }
    @{ Title = "UploadedDate"; InternalName = "UploadedDate"; Type = "DateTime"; Required = $false }
)

Write-Host "Connecting to destination site: $DestinationSiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

$list = Get-PnPList -Identity $script:LibraryName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "Creating document library: $script:LibraryName" -ForegroundColor Cyan
    $list = New-PnPList -Title $script:LibraryName -Template DocumentLibrary -OnQuickLaunch
} else {
    Write-Host "Library already exists, adding missing fields: $script:LibraryName" -ForegroundColor Yellow
}

$existingInternalNames = (Get-PnPField -List $script:LibraryName).InternalName

foreach ($f in $script:Fields) {
    if ($existingInternalNames -contains $f.InternalName) {
        Write-Host "  Skipping (already exists): $($f.InternalName)" -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Adding field: $($f.InternalName) ($($f.Type))"
    Add-PnPField -List $script:LibraryName -DisplayName $f.Title -InternalName $f.InternalName `
        -Type $f.Type -Required:$f.Required -AddToDefaultView | Out-Null

    if ($f.DefaultValue) {
        Set-PnPField -List $script:LibraryName -Identity $f.InternalName -Values @{ DefaultValue = $f.DefaultValue } | Out-Null
    }
}

$refreshedList = Get-PnPList -Identity $script:LibraryName
Write-Host ""
Write-Host "Done. DE_Attachments list ID (paste into api/local.settings.json / SWA app settings as SHAREPOINT_EXPENSE_ATTACHMENTS_LIST_ID):" -ForegroundColor Green
Write-Host "  $($refreshedList.Id.Guid)"
Write-Host "Verify the library in $DestinationSiteUrl before relying on it." -ForegroundColor Yellow
