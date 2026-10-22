<#
.SYNOPSIS
    Removes individual (per-user) "Limited Access" role assignments on one or
    more SharePoint lists, left over from the legacy app's per-item sharing
    model. Leaves group-based and any non-"Limited Access" assignments alone.

.DESCRIPTION
    The legacy Power-Automate-driven app shared each LR_LeaveRequests item
    individually with the employees who could see it, which left every one of
    them with an automatic "Limited Access" role assignment directly on the
    list itself — enough to navigate to items already shared with them, but
    never enough to create new items (see the "IT Assist" group Contribute
    fix, which is the actual replacement for this). Once that group-based fix
    is in place, these hundreds of individual "Limited Access" entries are
    pure cruft.

    Only removes a role assignment when BOTH are true:
      - the principal is an individual user (not a SharePoint/AAD group)
      - "Limited Access" is that user's ONLY role on the list

    Anyone with a real, higher-level individual grant (Full Control, Edit,
    Contribute, ...) is left untouched — this only cleans up the specific
    artifact described above, never a deliberate elevated permission.

    Prints every candidate before removing it, and a final summary of what
    was removed vs. kept, so the effect is visible in the console output.

.PARAMETER SiteUrl
    URL of the SharePoint site.

.PARAMETER ListNames
    Display names of the lists to clean up.

.PARAMETER ClientId
    Entra ID application (client) ID used for authentication.

.PARAMETER Thumbprint
    Thumbprint of the certificate (installed in the local CurrentUser
    certificate store) whose public key has been uploaded to the app
    registration.

.PARAMETER Tenant
    Tenant's initial/verified domain, e.g. "contoso.onmicrosoft.com".

.PARAMETER WhatIf
    List what would be removed without actually removing anything.

.EXAMPLE
    ./Remove-IndividualListPermissions.ps1 `
        -SiteUrl "https://itassist2.sharepoint.com/sites/LeaveRequest" `
        -ListNames "LR_LeaveRequests" `
        -ClientId "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" `
        -Thumbprint "9B0CFE16D7A505725109F90264E81A487AC6B41D" `
        -Tenant "itassist2.onmicrosoft.com" `
        -WhatIf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string[]]$ListNames = @("LR_LeaveRequests"),

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $true)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"

Write-Host "Connecting to: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant

foreach ($listName in $ListNames) {
    Write-Host ""
    Write-Host "=== $listName ===" -ForegroundColor Cyan

    $list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "  List not found, skipping." -ForegroundColor Yellow
        continue
    }

    $roleAssignments = Get-PnPProperty -ClientObject $list -Property RoleAssignments

    $toRemove = @()
    $kept = @()
    foreach ($ra in $roleAssignments) {
        $member = Get-PnPProperty -ClientObject $ra -Property Member
        Get-PnPProperty -ClientObject $member -Property PrincipalType, Title, LoginName | Out-Null
        $bindings = Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings
        $roleNames = @($bindings | ForEach-Object { Get-PnPProperty -ClientObject $_ -Property Name | Out-Null; $_.Name })

        $isIndividual = $member.PrincipalType -eq "User"
        $isOnlyLimitedAccess = ($roleNames.Count -eq 1) -and ($roleNames[0] -eq "Limited Access")

        if ($isIndividual -and $isOnlyLimitedAccess) {
            $toRemove += [PSCustomObject]@{ Title = $member.Title; LoginName = $member.LoginName }
        } else {
            $kept += "$($member.Title) [$($roleNames -join ',')]"
        }
    }

    Write-Host "  Individual 'Limited Access' grants to remove: $($toRemove.Count)" -ForegroundColor Yellow
    $toRemove | ForEach-Object { Write-Host "    - $($_.Title)" }
    Write-Host "  Other assignments kept as-is: $($kept.Count)" -ForegroundColor DarkGray
    $kept | ForEach-Object { Write-Host "    - $_" -ForegroundColor DarkGray }

    if ($WhatIf) {
        Write-Host "  -WhatIf specified, not removing anything." -ForegroundColor Magenta
        continue
    }

    foreach ($u in $toRemove) {
        Set-PnPListPermission -Identity $listName -User $u.LoginName -RemoveRole "Limited Access" | Out-Null
    }
    Write-Host "  Removed $($toRemove.Count) individual grant(s) from $listName." -ForegroundColor Green
}

Disconnect-PnPOnline
Write-Host ""
Write-Host "Done." -ForegroundColor Green
