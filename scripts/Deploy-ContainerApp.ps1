<#
.SYNOPSIS
    Creates the ca-internal-processes Container App in the existing
    "internal-processes" environment, using ca-leaverequestforteams as the
    template (ingress, registry login, probes, scale rules) and copying its
    secrets across.

.DESCRIPTION
    Nothing here is hard-coded secret material: the four secrets
    (ghcr-pat, aad-client-secret, tenant-config-conn, azure-webjobs-storage)
    are read from the source app with `az containerapp secret show`, kept in
    memory, written to a temporary YAML for `az containerapp create --yaml`,
    and that file is deleted in a finally block. Nothing is printed.

    The new app gets the unified image and the new Actionable Messages
    originator (e1414a49…); everything else mirrors leave. ca-leaverequestforteams
    and ca-expensereportforteams are NOT touched — they keep serving the emails
    already sent.

    Run it from a terminal where `az` is signed in to the IT ASSIST tenant.
    -WhatIf prints what it would create (no secret values) and stops.

.PARAMETER Image
    Image to run. Defaults to the :latest tag the build workflow pushes.

.EXAMPLE
    ./Deploy-ContainerApp.ps1 -WhatIf
    ./Deploy-ContainerApp.ps1
    ./Deploy-ContainerApp.ps1 -Image ghcr.io/gabidodu/ita-internalprocesses:330e8aa23936a58e6d14b7f3d3b99b80316bb0db
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Image = "ghcr.io/gabidodu/ita-internalprocesses:latest",
    [string]$SourceApp = "ca-leaverequestforteams",
    [string]$AppName = "ca-internal-processes",
    [string]$ResourceGroup = "rg-internal-processes",
    [string]$Subscription = "2c0c5038-ac67-487a-bc6e-3fbc5753ebee",
    [string]$OriginatorId = "e1414a49-1a3e-4bba-94b4-e27134577476"
)

$ErrorActionPreference = "Stop"

function Invoke-Az {
    $out = & az @args
    if ($LASTEXITCODE -ne 0) { throw "az $($args -join ' ') failed (exit $LASTEXITCODE)" }
    return $out
}

$src = (Invoke-Az containerapp show -n $SourceApp -g $ResourceGroup --subscription $Subscription -o json) -join "`n" | ConvertFrom-Json
$p = $src.properties
$container = $p.template.containers[0]

$domain = ($p.configuration.ingress.fqdn -split "\.", 2)[1]
$hostName = "$AppName.$domain"
$audience = "api://auth-am-$OriginatorId/359906f4-0bbf-4810-97f4-0d1b28aa5f1a"

$env = @(
    @{ name = "FUNCTIONS_WORKER_RUNTIME"; value = "node" },
    @{ name = "AzureWebJobsStorage"; secretRef = "azure-webjobs-storage" },
    @{ name = "AAD_APP_CLIENT_ID"; value = "359906f4-0bbf-4810-97f4-0d1b28aa5f1a" },
    @{ name = "AAD_APP_CLIENT_SECRET"; secretRef = "aad-client-secret" },
    @{ name = "TENANT_CONFIG_STORAGE_CONNECTION_STRING"; secretRef = "tenant-config-conn" },
    @{ name = "TAB_BASE_URL"; value = "https://$hostName" },
    @{ name = "TEAMS_APP_ID"; value = "661a64f0-7ed1-413c-a709-4ea98ef49fd7" },
    @{ name = "ACTIONABLE_MESSAGE_ORIGINATOR_ID"; value = $OriginatorId },
    @{ name = "ACTIONABLE_MESSAGE_AUDIENCE"; value = $audience },
    @{ name = "CALENDAR_NAME"; value = "CalendarConcedii" },
    @{ name = "TZ"; value = "Europe/Bucharest" }
)

Write-Host "Will create $AppName in $($p.environmentId.Split('/')[-1]) ($ResourceGroup)" -ForegroundColor Cyan
Write-Host "  image:  $Image"
Write-Host "  host:   https://$hostName"
Write-Host "  scale:  min $($p.template.scale.minReplicas), max $($p.template.scale.maxReplicas), $($p.template.scale.rules.Count) rules (copied from $SourceApp)"
$env | ForEach-Object { Write-Host ("  env {0,-42} {1}" -f $_.name, $(if ($_.secretRef) { "secret:$($_.secretRef)" } else { $_.value })) -ForegroundColor DarkGray }

if (-not $PSCmdlet.ShouldProcess("$AppName in $ResourceGroup", "Create Container App")) { return }

# Same ingress / registry login as the source, so the new app pulls from ghcr the same way.
$secrets = foreach ($name in "ghcr-pat", "aad-client-secret", "tenant-config-conn", "azure-webjobs-storage") {
    $value = (Invoke-Az containerapp secret show -n $SourceApp -g $ResourceGroup --subscription $Subscription --secret-name $name --query value -o tsv)
    if (-not $value) { throw "Secret '$name' came back empty from $SourceApp." }
    @{ name = $name; value = "$value".Trim() }
}

$config = @{
    location   = $src.location
    properties = @{
        environmentId = $p.environmentId
        configuration = @{
            activeRevisionsMode = "Single"
            ingress             = @{ external = $true; targetPort = 8080; transport = "Auto"; allowInsecure = $false; traffic = @(@{ latestRevision = $true; weight = 100 }) }
            registries          = @(@{ server = "ghcr.io"; username = $p.configuration.registries[0].username; passwordSecretRef = "ghcr-pat" })
            secrets             = @($secrets)
        }
        template      = @{
            containers = @(@{ name = "app"; image = $Image; env = $env; resources = @{ cpu = 0.25; memory = "0.5Gi" }; probes = $container.probes })
            scale      = $p.template.scale
        }
    }
}

$tmp = Join-Path ([IO.Path]::GetTempPath()) ("$AppName-" + [guid]::NewGuid().ToString("N") + ".json")
try {
    # No BOM: Windows PowerShell 5.1's "-Encoding utf8" adds one, which `az --yaml` rejects.
    [IO.File]::WriteAllText($tmp, ($config | ConvertTo-Json -Depth 20), (New-Object Text.UTF8Encoding $false))
    Invoke-Az containerapp create -n $AppName -g $ResourceGroup --subscription $Subscription --yaml $tmp --query "properties.configuration.ingress.fqdn" -o tsv
}
finally {
    Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
}

Write-Host "Created. Check the revision: az containerapp revision list -n $AppName -g $ResourceGroup --subscription $Subscription -o table" -ForegroundColor Green
Write-Host "Health: https://$hostName/healthz" -ForegroundColor Green
