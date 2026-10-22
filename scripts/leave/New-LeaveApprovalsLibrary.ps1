<#
.SYNOPSIS
    Creates the LR_Approvals document library (archived copies of the
    approved-leave-request PDF) on a destination site, with its metadata
    columns, using an Entra ID app-only registration (certificate auth) via
    PnP.PowerShell.

.DESCRIPTION
    On final approval, RespondToLeaveRequest.ts / RespondToLeaveRequestAction.ts
    already email the generated PDF — this library gives it a permanent,
    browsable home too, so an approved request survives even if the
    notification email is deleted. Files land in a per-employee subfolder
    (e.g. "Popescu Ion/LeaveRequest_....pdf") — see
    api/src/lib/approvalArchive.ts. Safe to re-run: if the library or a field
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
    ./New-LeaveApprovalsLibrary.ps1 `
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

$script:LibraryName = "LR_Approvals"
$script:Fields = @(
    @{ Title = "EmployeeCode"; InternalName = "EmployeeCode"; Type = "Text"; Required = $false }
    @{ Title = "RequestId"; InternalName = "RequestId"; Type = "Text"; Required = $false }
    @{ Title = "ApprovedDate"; InternalName = "ApprovedDate"; Type = "DateTime"; Required = $false }
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
}

$refreshedList = Get-PnPList -Identity $script:LibraryName
Write-Host ""
Write-Host "Done. LR_Approvals list ID (this is the tenant config's 'approvalsLibraryId'):" -ForegroundColor Green
Write-Host "  $($refreshedList.Id.Guid)"
Write-Host "Verify the library in $DestinationSiteUrl before relying on it." -ForegroundColor Yellow
