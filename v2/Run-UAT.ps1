param([ValidateSet('work','home')][string]$MachineRole='work')
$ErrorActionPreference='Stop'
$packageRoot=$PSScriptRoot
$manifest=Get-Content -LiteralPath (Join-Path $packageRoot 'manifest.json') -Raw | ConvertFrom-Json
foreach ($entry in $manifest.files) {
    $candidate=[IO.Path]::GetFullPath((Join-Path $packageRoot $entry.path))
    if (-not $candidate.StartsWith($packageRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe manifest path' }
    $item=Get-Item -LiteralPath $candidate
    if ($item.Length -ne $entry.bytes -or (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.sha256) { throw "Package verification failed: $($entry.path)" }
}
$evidenceRoot=Join-Path $env:LOCALAPPDATA ('NMDC-V2-UAT/'+$MachineRole+'-'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Path $evidenceRoot | Out-Null
$workspace=Join-Path $evidenceRoot 'workspace'
Copy-Item -LiteralPath (Join-Path $packageRoot 'ACCEPTANCE.csv') -Destination (Join-Path $evidenceRoot 'results.csv')
@{ machine_role=$MachineRole; build=$manifest.commit; executable_sha256=(Get-FileHash -LiteralPath (Join-Path $packageRoot 'Launch.exe') -Algorithm SHA256).Hash.ToLowerInvariant(); os=[Environment]::OSVersion.VersionString; started_utc=[DateTime]::UtcNow.ToString('o'); status='NOT RUN - fill results.csv using the guide'; package_verified=$true } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $evidenceRoot 'environment.json') -Encoding utf8
Write-Output "Package verified. Isolated workspace and evidence: $evidenceRoot"
Write-Output 'Follow UAT-GUIDE.md. Keep evidence local; share only sanitized results. This runner does not approve any acceptance case.'
Start-Process -FilePath (Join-Path $packageRoot 'Launch.exe') -ArgumentList @('--workspace',('"'+$workspace+'"')) -WindowStyle Hidden
