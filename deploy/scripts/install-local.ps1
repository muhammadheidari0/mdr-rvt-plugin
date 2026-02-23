param(
    [string]$RevitVersion = "2026",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [switch]$SkipBuild,
    [switch]$ForceConfig,
    [string]$GoogleClientId = $env:MDR_GOOGLE_CLIENT_ID,
    [string]$GoogleClientSecret = $env:MDR_GOOGLE_CLIENT_SECRET,
    [string]$GoogleRefreshToken = $env:MDR_GOOGLE_REFRESH_TOKEN
)

$ErrorActionPreference = "Stop"

function Resolve-RepoRoot {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDir)) {
        $scriptDir = Split-Path -Parent $PSCommandPath
    }
    return (Resolve-Path (Join-Path $scriptDir "..\..")).Path
}

function Resolve-Dotnet {
    param([string]$RepoRoot)
    $localDotnet = Join-Path $RepoRoot ".dotnet\dotnet.exe"
    if (Test-Path $localDotnet) {
        return $localDotnet
    }
    return "dotnet"
}

function Ensure-Directory {
    param([string]$PathValue)
    if (-not (Test-Path $PathValue)) {
        New-Item -ItemType Directory -Path $PathValue | Out-Null
    }
}

function Assert-RevitNotRunning {
    $running = Get-Process -Name "Revit" -ErrorAction SilentlyContinue
    if ($null -eq $running) {
        return
    }

    $ids = ($running | Select-Object -ExpandProperty Id) -join ", "
    throw "Autodesk Revit is running (PID: $ids). Close Revit and re-run install-local.ps1."
}

function Write-Manifest {
    param(
        [string]$ManifestPath,
        [string]$AssemblyPath
    )

    $escapedAssembly = [System.Security.SecurityElement]::Escape($AssemblyPath)
    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>MDR Revit Plugin</Name>
    <Assembly>$escapedAssembly</Assembly>
    <AddInId>8D83C886-B739-4ACD-A9DB-1BC78F315B3A</AddInId>
    <FullClassName>Mdr.Revit.Addin.RevitExternalApplication</FullClassName>
    <VendorId>MDR</VendorId>
    <VendorDescription>MDR BIM Integration</VendorDescription>
  </AddIn>
</RevitAddIns>
"@

    Set-Content -Path $ManifestPath -Value $xml -Encoding UTF8
}

function Apply-ConfigSecrets {
    param(
        [string]$ConfigPath,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$RefreshToken
    )

    if (-not (Test-Path $ConfigPath)) {
        return
    }

    $hasGoogleValues = (-not [string]::IsNullOrWhiteSpace($ClientId)) -or
        (-not [string]::IsNullOrWhiteSpace($ClientSecret)) -or
        (-not [string]::IsNullOrWhiteSpace($RefreshToken))
    if (-not $hasGoogleValues) {
        return
    }

    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    if (-not ($config.PSObject.Properties.Name -contains "google")) {
        $config | Add-Member -NotePropertyName "google" -NotePropertyValue ([pscustomobject]@{})
    }

    $google = $config.google
    if (-not ($google.PSObject.Properties.Name -contains "clientId")) {
        $google | Add-Member -NotePropertyName "clientId" -NotePropertyValue ""
    }
    if (-not ($google.PSObject.Properties.Name -contains "clientSecret")) {
        $google | Add-Member -NotePropertyName "clientSecret" -NotePropertyValue ""
    }
    if (-not ($google.PSObject.Properties.Name -contains "refreshToken")) {
        $google | Add-Member -NotePropertyName "refreshToken" -NotePropertyValue ""
    }

    if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
        $google.clientId = $ClientId
    }
    if (-not [string]::IsNullOrWhiteSpace($ClientSecret)) {
        $google.clientSecret = $ClientSecret
    }
    if (-not [string]::IsNullOrWhiteSpace($RefreshToken)) {
        $google.refreshToken = $RefreshToken
    }

    $config | ConvertTo-Json -Depth 20 | Set-Content -Path $ConfigPath -Encoding UTF8
    Write-Host "==> Applied Google OAuth values to runtime config"
}

$repoRoot = Resolve-RepoRoot
$dotnet = Resolve-Dotnet -RepoRoot $repoRoot
$solutionPath = Join-Path $repoRoot "Mdr.RevitPlugin.sln"
$addinOutput = Join-Path $repoRoot "src\Mdr.Revit.Addin\bin\$Configuration\net48"

$pluginInstallDir = Join-Path $env:LocalAppData "MDR\RevitPlugin\addin\$RevitVersion"
$runtimeConfigDir = Join-Path $env:LocalAppData "MDR\RevitPlugin"
$runtimeConfigPath = Join-Path $runtimeConfigDir "config.json"
$templateConfigPath = Join-Path $repoRoot "config\appsettings.template.json"
$revitAddinsDir = Join-Path $env:AppData "Autodesk\Revit\Addins\$RevitVersion"
$manifestPath = Join-Path $revitAddinsDir "Mdr.Revit.addin"

Write-Host "==> MDR Revit local install"
Write-Host "Repo root: $repoRoot"
Write-Host "Revit version: $RevitVersion"
Write-Host "Configuration: $Configuration"

Assert-RevitNotRunning

if (-not $SkipBuild) {
    Write-Host "==> Building solution"
    & $dotnet build $solutionPath -c $Configuration
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet build failed with exit code $LASTEXITCODE."
    }
}

if (-not (Test-Path $addinOutput)) {
    throw "Addin output not found: $addinOutput"
}

Ensure-Directory -PathValue $pluginInstallDir
Ensure-Directory -PathValue $runtimeConfigDir
Ensure-Directory -PathValue $revitAddinsDir

Write-Host "==> Copying plugin binaries to $pluginInstallDir"
Copy-Item -Path (Join-Path $addinOutput "*") -Destination $pluginInstallDir -Recurse -Force

if ($ForceConfig -or -not (Test-Path $runtimeConfigPath)) {
    if (-not (Test-Path $templateConfigPath)) {
        throw "Template config not found: $templateConfigPath"
    }
    Write-Host "==> Writing runtime config to $runtimeConfigPath"
    Copy-Item -Path $templateConfigPath -Destination $runtimeConfigPath -Force
} else {
    Write-Host "==> Runtime config already exists: $runtimeConfigPath"
}

Apply-ConfigSecrets -ConfigPath $runtimeConfigPath -ClientId $GoogleClientId -ClientSecret $GoogleClientSecret -RefreshToken $GoogleRefreshToken

$assemblyPath = Join-Path $pluginInstallDir "Mdr.Revit.Addin.dll"
if (-not (Test-Path $assemblyPath)) {
    throw "Main addin assembly not found after copy: $assemblyPath"
}

Write-Host "==> Writing Revit addin manifest: $manifestPath"
Write-Manifest -ManifestPath $manifestPath -AssemblyPath $assemblyPath

Write-Host ""
Write-Host "Install completed."
Write-Host "Assembly: $assemblyPath"
Write-Host "Manifest: $manifestPath"
Write-Host "Config:   $runtimeConfigPath"
