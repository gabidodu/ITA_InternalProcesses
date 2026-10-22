<#
.SYNOPSIS
    Copies the schema (custom columns) of one or more SharePoint lists from a
    source site to a destination site, using an Entra ID app-only registration
    (certificate auth) via PnP.PowerShell.

.DESCRIPTION
    Reads each list's custom fields (display name, internal name, type, required,
    default value) from the source site, then recreates the list (if missing) and
    adds each field individually on the destination site.

    Fields are added one at a time via Add-PnPField rather than through PnP's site
    template engine (Get-PnPSiteTemplate / Invoke-PnPSiteTemplate). The template
    engine was tried first but exhibited a reproducible bug on this tenant: fields
    from one list ended up attached to a different, previously-processed list, even
    with a single list per template and a fresh connection per apply. Adding fields
    directly against an explicitly-resolved target list avoids that failure mode.

    Only column schema is recreated, not classic list views (custom CAML views are
    fragile to reproduce this way and aren't needed by a Teams-native front end that
    reads/writes lists via REST/Graph). Each copied field is added to the default
    view so the list is still usable from the SharePoint UI.

    Designed to be reusable for future client-tenant deployments: source/destination,
    tenant and credentials are all parameters, nothing is hardcoded.

.PARAMETER SourceSiteUrl
    URL of the SharePoint site containing the original lists.

.PARAMETER DestinationSiteUrl
    URL of the SharePoint site where the lists should be recreated.

.PARAMETER ListNames
    Display names of the lists to copy.

.PARAMETER ClientId
    Entra ID application (client) ID used for authentication.

.PARAMETER Thumbprint
    Thumbprint of the certificate (installed in the local CurrentUser certificate
    store) whose public key has been uploaded to the app registration.

.PARAMETER Tenant
    Tenant's initial/verified domain, e.g. "contoso.onmicrosoft.com".

.PARAMETER SchemaExportPath
    Where to save the captured schema as JSON (for reference/audit, and so the
    same schema could later be re-applied without needing access to the source
    site again). Defaults to a "schema" folder next to this script.

.EXAMPLE
    ./Copy-SharePointListSchema.ps1 `
        -SourceSiteUrl "https://itassist2.sharepoint.com/sites/LeaveRequest" `
        -DestinationSiteUrl "https://itassist2.sharepoint.com/sites/AVI_Site" `
        -ListNames "LR_LeaveStatus","LR_LeaveSetup","LR_LeaveRequests","LR_HolidaySetup" `
        -ClientId "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "itassist2.onmicrosoft.com"

.NOTES
    Required app permissions: the app registration needs Application permission
    Sites.Manage.All (or Sites.FullControl.All) on SharePoint with admin consent,
    OR Sites.Selected with explicit Grant-PnPAzureADAppSitePermission on both site
    collections (write on destination, read is enough on source).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceSiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$DestinationSiteUrl,

    [Parameter(Mandatory = $true)]
    [string[]]$ListNames,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $true)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [string]$SchemaExportPath = (Join-Path $PSScriptRoot "schema")
)

$ErrorActionPreference = "Stop"

# Built-in / system fields present on every list — never copied as "custom" columns.
$script:BuiltInFieldNames = @(
    'ID','Title','LinkTitle','LinkTitleNoMenu','ContentType','Created','Modified',
    'Author','Editor','Attachments','Edit','DocIcon','_UIVersionString',
    'ItemChildCount','FolderChildCount','_ComplianceFlags','_ComplianceTag',
    '_ComplianceTagWrittenTime','_ComplianceTagUserId','_IsRecord','AppAuthor',
    'AppEditor','_ColorTag','ComplianceAssetId','GUID','FileLeafRef','FileRef',
    'FSObjType','owshiddenversion','WorkflowVersion'
)

New-Item -ItemType Directory -Force -Path $SchemaExportPath | Out-Null

Write-Host "Connecting to source site: $SourceSiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SourceSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

$listSchemas = @{}
foreach ($name in $ListNames) {
    $list = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
    if (-not $list) {
        throw "List '$name' was not found on $SourceSiteUrl. Check the exact display name."
    }
    Write-Host "Reading fields for: $name (ItemCount: $($list.ItemCount))" -ForegroundColor Cyan

    $fields = Get-PnPField -List $name |
        Where-Object { $script:BuiltInFieldNames -notcontains $_.InternalName } |
        ForEach-Object {
            [PSCustomObject]@{
                Title        = $_.Title
                InternalName = $_.InternalName
                Type         = $_.TypeAsString
                Required     = $_.Required
                DefaultValue = $_.DefaultValue
            }
        }

    foreach ($f in $fields) {
        Write-Host "  $($f.InternalName) ($($f.Type))"
    }
    $listSchemas[$name] = $fields

    $exportFile = Join-Path $SchemaExportPath ("$($name -replace '[^\w\-]', '_').json")
    $fields | ConvertTo-Json -Depth 5 | Set-Content -Path $exportFile
}

Disconnect-PnPOnline

foreach ($name in $ListNames) {
    Write-Host "Connecting to destination site: $DestinationSiteUrl" -ForegroundColor Cyan
    Connect-PnPOnline -Url $DestinationSiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

    $destList = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
    if (-not $destList) {
        Write-Host "Creating list: $name" -ForegroundColor Cyan
        $destList = New-PnPList -Title $name -Template GenericList -OnQuickLaunch
    }
    else {
        Write-Host "List already exists, adding missing fields: $name" -ForegroundColor Yellow
    }

    $existingInternalNames = (Get-PnPField -List $name).InternalName

    foreach ($f in $listSchemas[$name]) {
        if ($existingInternalNames -contains $f.InternalName) {
            Write-Host "  Skipping (already exists): $($f.InternalName)" -ForegroundColor DarkGray
            continue
        }

        Write-Host "  Adding field: $($f.InternalName) ($($f.Type))"
        Add-PnPField -List $name -DisplayName $f.Title -InternalName $f.InternalName `
            -Type $f.Type -Required:$f.Required -AddToDefaultView | Out-Null

        if ($f.DefaultValue) {
            Set-PnPField -List $name -Identity $f.InternalName -Values @{ DefaultValue = $f.DefaultValue } | Out-Null
        }
    }

    Disconnect-PnPOnline
}

Write-Host "Done. Captured schema saved under: $SchemaExportPath" -ForegroundColor Green
Write-Host "Verify the lists in $DestinationSiteUrl before relying on them." -ForegroundColor Yellow
