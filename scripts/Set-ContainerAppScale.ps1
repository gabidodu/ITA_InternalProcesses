<#
.SYNOPSIS
    Applies a scaling profile from infra/ to the Container Apps, and reports drift.

.DESCRIPTION
    The scaling policy: minReplicas = 0 (scale to zero when idle), plus a cron
    rule that keeps ONE replica running Monday-Friday 08:00-20:00 Romania time,
    so the business-hours cold start never hits users. Outside those hours the
    app scales to zero; the first request after that starts it from cold.

    Source of truth. Without -Apply the script only reads
    each app's current scale and says whether it matches; nothing is written.

.PARAMETER Profile
    "production" (default) or "testing".

.PARAMETER Apply
    Write the profile to the apps whose scale differs.

.PARAMETER Apps
    Apps to check/update. Defaults to the three apps in the shared environment.

.EXAMPLE
    ./Set-ContainerAppScale.ps1 -Profile testing              # check only
    ./Set-ContainerAppScale.ps1 -Profile testing -Apply       # scale to zero, no schedule
    ./Set-ContainerAppScale.ps1 -Profile production -Apply    # going live: add the 8-20 rule
#>
[CmdletBinding()]
param(
    [ValidateSet("production", "testing")]
    [string]$Profile = "production",
    [switch]$Apply,
    [string[]]$Apps = @("ca-leaverequestforteams", "ca-expensereportforteams", "ca-internal-processes"),
    [string]$ResourceGroup = "rg-internal-processes",
    [string]$Subscription = "2c0c5038-ac67-487a-bc6e-3fbc5753ebee"
)

$ErrorActionPreference = "Stop"
$desiredPath = Join-Path $PSScriptRoot "..\infra\scale.$Profile.json"
$desired = Get-Content $desiredPath -Raw | ConvertFrom-Json

function Invoke-Az {
    $out = & az @args
    if ($LASTEXITCODE -ne 0) { throw "az $($args -join ' ') failed (exit $LASTEXITCODE)" }
    return $out
}

# Order-insensitive, compares only the fields infra/scale.json defines.
function Get-Canonical($scale) {
    $rules = foreach ($r in $scale.rules) {
        $kind = if ($r.http) { "http" } else { "custom" }
        $meta = if ($r.http) { $r.http.metadata } else { $r.custom.metadata }
        $type = if ($r.custom) { $r.custom.type } else { "" }
        $pairs = ($meta.PSObject.Properties | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ";"
        "$($r.name)|$kind|$type|$pairs"
    }
    "min=$($scale.minReplicas);max=$($scale.maxReplicas);cool=$($scale.cooldownPeriod);poll=$($scale.pollingInterval);" + (($rules | Sort-Object) -join "||")
}

$want = Get-Canonical $desired
foreach ($app in $Apps) {
    $current = (Invoke-Az containerapp show -n $app -g $ResourceGroup --subscription $Subscription --query "properties.template.scale" -o json) -join "`n" | ConvertFrom-Json
    if ((Get-Canonical $current) -eq $want) {
        Write-Host "[ok]    $app matches the $Profile profile" -ForegroundColor Green
        continue
    }
    Write-Host "[drift] $app differs from the $Profile profile" -ForegroundColor Yellow
    Write-Host "        current: min $($current.minReplicas), max $($current.maxReplicas), rules: $((($current.rules | ForEach-Object { $_.name }) -join ', '))"
    if (-not $Apply) { continue }

    $yaml = @{ properties = @{ template = @{ scale = $desired } } }
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("scale-$app-" + [guid]::NewGuid().ToString("N") + ".json")
    try {
        # No BOM: Windows PowerShell 5.1's "-Encoding utf8" adds one, which `az --yaml` rejects.
        [IO.File]::WriteAllText($tmp, ($yaml | ConvertTo-Json -Depth 20), (New-Object Text.UTF8Encoding $false))
        Invoke-Az containerapp update -n $app -g $ResourceGroup --subscription $Subscription --yaml $tmp --query "properties.provisioningState" -o tsv | Out-Null
        Write-Host "[fixed] $app" -ForegroundColor Green
    }
    finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}
