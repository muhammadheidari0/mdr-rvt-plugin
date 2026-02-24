param(
    [Parameter(Mandatory = $true)]
    [string]$PackagePath,
    [string]$RevitVersion = "2026",
    [switch]$ForceConfig,
    [string]$ConfigTemplatePath = "",
    [switch]$KeepExtractedPackage,
    [string]$GoogleClientId = $env:MDR_GOOGLE_CLIENT_ID,
    [string]$GoogleClientSecret = $env:MDR_GOOGLE_CLIENT_SECRET,
    [string]$GoogleRefreshToken = $env:MDR_GOOGLE_REFRESH_TOKEN
)

$ErrorActionPreference = "Stop"

function Ensure-Directory {
    param([string]$PathValue)

    if (-not (Test-Path -Path $PathValue)) {
        New-Item -ItemType Directory -Path $PathValue | Out-Null
    }
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

    if (-not (Test-Path -Path $ConfigPath)) {
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

function Resolve-PackageRoot {
    param([string]$PathValue)

    if (-not (Test-Path -LiteralPath $PathValue)) {
        throw "Package path not found: $PathValue"
    }

    $resolved = (Resolve-Path -LiteralPath $PathValue).Path
    $item = Get-Item -LiteralPath $resolved
    if ($item.PSIsContainer) {
        return @{
            RootPath = $resolved
            TempPath = ""
        }
    }

    if ([System.IO.Path]::GetExtension($resolved).Equals(".zip", [System.StringComparison]::OrdinalIgnoreCase)) {
        $extractPath = Join-Path $env:TEMP ("mdr-revit-plugin-" + [guid]::NewGuid().ToString("N"))
        Ensure-Directory -PathValue $extractPath
        Write-Host "==> Extracting package zip to $extractPath"
        Expand-Archive -Path $resolved -DestinationPath $extractPath -Force
        return @{
            RootPath = $extractPath
            TempPath = $extractPath
        }
    }

    throw "PackagePath must be a directory or .zip file: $resolved"
}

function Get-BinaryCandidateScore {
    param([System.IO.FileInfo]$File)

    $path = $File.FullName.ToLowerInvariant()
    $score = 1000
    if ($path -like "*\src\mdr.revit.addin\bin\release\net8.0-windows\mdr.revit.addin.dll") {
        $score = 0
    } elseif ($path -like "*\bin\release\net8.0-windows\mdr.revit.addin.dll") {
        $score = 10
    } elseif ($path -like "*\src\mdr.revit.addin\bin\release\net8.0\mdr.revit.addin.dll") {
        $score = 15
    } elseif ($path -like "*\bin\release\net8.0\mdr.revit.addin.dll") {
        $score = 20
    } elseif ($path -like "*\bin\release\net48\mdr.revit.addin.dll") {
        $score = 30
    } elseif ($path -like "*\bin\release\*") {
        $score = 40
    } elseif ($path -like "*\src\mdr.revit.addin\bin\debug\net8.0-windows\mdr.revit.addin.dll") {
        $score = 50
    } elseif ($path -like "*\bin\debug\net8.0-windows\mdr.revit.addin.dll") {
        $score = 60
    } elseif ($path -like "*\src\mdr.revit.addin\bin\debug\net8.0\mdr.revit.addin.dll") {
        $score = 70
    } elseif ($path -like "*\bin\debug\net8.0\mdr.revit.addin.dll") {
        $score = 80
    } elseif ($path -like "*\src\mdr.revit.addin\bin\debug\net48\mdr.revit.addin.dll") {
        $score = 90
    } elseif ($path -like "*\bin\debug\net48\mdr.revit.addin.dll") {
        $score = 100
    } elseif ($path -like "*\bin\debug\*") {
        $score = 110
    } elseif ($path -like "*\net8.0-windows\*") {
        $score = 120
    } elseif ($path -like "*\net8.0\*") {
        $score = 130
    } elseif ($path -like "*\net48\*") {
        $score = 140
    }

    return $score
}

function Resolve-BinarySourceDirectory {
    param([string]$PackageRoot)

    $candidates = Get-ChildItem -Path $PackageRoot -Recurse -File -Filter "Mdr.Revit.Addin.dll" |
        Where-Object {
            $_.FullName -notlike "*\obj\*" -and
            $_.FullName -notlike "*\ref\*"
        }
    if ($null -eq $candidates -or $candidates.Count -eq 0) {
        throw "Could not find Mdr.Revit.Addin.dll in package."
    }

    $selected = $candidates |
        Sort-Object @{ Expression = { Get-BinaryCandidateScore -File $_ } }, @{ Expression = { $_.FullName.Length } } |
        Select-Object -First 1
    return (Split-Path -Parent $selected.FullName)
}

function Resolve-TemplateConfigPath {
    param(
        [string]$PackageRoot,
        [string]$ExplicitPath
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        if (-not (Test-Path -LiteralPath $ExplicitPath -PathType Leaf)) {
            throw "Config template path not found: $ExplicitPath"
        }
        return (Resolve-Path -LiteralPath $ExplicitPath).Path
    }

    $defaultPath = Join-Path $PackageRoot "config\appsettings.template.json"
    if (Test-Path -LiteralPath $defaultPath -PathType Leaf) {
        return $defaultPath
    }

    $matches = Get-ChildItem -Path $PackageRoot -Recurse -File -Filter "appsettings.template.json" |
        Sort-Object FullName
    if ($null -eq $matches -or $matches.Count -eq 0) {
        return ""
    }

    return $matches[0].FullName
}

$pluginInstallDir = Join-Path $env:LocalAppData "MDR\RevitPlugin\addin\$RevitVersion"
$runtimeConfigDir = Join-Path $env:LocalAppData "MDR\RevitPlugin"
$runtimeConfigPath = Join-Path $runtimeConfigDir "config.json"
$revitAddinsDir = Join-Path $env:AppData "Autodesk\Revit\Addins\$RevitVersion"
$manifestPath = Join-Path $revitAddinsDir "Mdr.Revit.addin"

$packageInfo = Resolve-PackageRoot -PathValue $PackagePath
$packageRoot = $packageInfo.RootPath
$tempExtractPath = $packageInfo.TempPath

Write-Host "==> MDR Revit install from package"
Write-Host "Package root:  $packageRoot"
Write-Host "Revit version: $RevitVersion"

try {
    $binarySourceDir = Resolve-BinarySourceDirectory -PackageRoot $packageRoot
    Write-Host "==> Binary source: $binarySourceDir"

    Ensure-Directory -PathValue $pluginInstallDir
    Ensure-Directory -PathValue $runtimeConfigDir
    Ensure-Directory -PathValue $revitAddinsDir

    Write-Host "==> Copying binaries to $pluginInstallDir"
    Get-ChildItem -Path $pluginInstallDir -Force | Remove-Item -Recurse -Force
    Copy-Item -Path (Join-Path $binarySourceDir "*") -Destination $pluginInstallDir -Recurse -Force

    $templateConfigPath = Resolve-TemplateConfigPath -PackageRoot $packageRoot -ExplicitPath $ConfigTemplatePath
    if ($ForceConfig -or -not (Test-Path -LiteralPath $runtimeConfigPath)) {
        if ([string]::IsNullOrWhiteSpace($templateConfigPath)) {
            throw "Runtime config does not exist and no appsettings.template.json was found in package. Use -ConfigTemplatePath."
        }

        Write-Host "==> Writing runtime config to $runtimeConfigPath"
        Copy-Item -Path $templateConfigPath -Destination $runtimeConfigPath -Force
    } else {
        Write-Host "==> Runtime config already exists: $runtimeConfigPath"
    }

    Apply-ConfigSecrets -ConfigPath $runtimeConfigPath -ClientId $GoogleClientId -ClientSecret $GoogleClientSecret -RefreshToken $GoogleRefreshToken

    $assemblyPath = Join-Path $pluginInstallDir "Mdr.Revit.Addin.dll"
    if (-not (Test-Path -LiteralPath $assemblyPath)) {
        throw "Main addin assembly not found after copy: $assemblyPath"
    }

    Write-Host "==> Writing Revit addin manifest: $manifestPath"
    Write-Manifest -ManifestPath $manifestPath -AssemblyPath $assemblyPath

    Write-Host ""
    Write-Host "Install completed."
    Write-Host "Assembly: $assemblyPath"
    Write-Host "Manifest: $manifestPath"
    Write-Host "Config:   $runtimeConfigPath"
}
finally {
    if (-not [string]::IsNullOrWhiteSpace($tempExtractPath) -and
        (Test-Path -LiteralPath $tempExtractPath) -and
        -not $KeepExtractedPackage) {
        Remove-Item -Path $tempExtractPath -Recurse -Force
    }
}
