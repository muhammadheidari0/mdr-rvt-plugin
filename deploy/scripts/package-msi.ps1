param(
    [string]$RevitVersion = "2026",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [string]$Version = "",
    [string]$OutputDirectory = "",
    [switch]$SkipBuild,
    [switch]$SkipZip,
    [switch]$SkipMsi,
    [switch]$IncludePdb,
    [switch]$KeepStaging,
    [string]$WixPath = "",
    [string]$ProductName = "MDR Revit Plugin",
    [string]$Manufacturer = "MDR BIM Integration",
    [string]$UpgradeCode = "{D80D7022-8046-41D7-9B4C-C64D79F37AF5}"
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

function Resolve-PluginVersion {
    param(
        [string]$RepoRoot,
        [string]$RequestedVersion
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedVersion)) {
        return $RequestedVersion.Trim()
    }

    $templateConfig = Join-Path $RepoRoot "config\appsettings.template.json"
    if (-not (Test-Path $templateConfig)) {
        return "0.3.0"
    }

    try {
        $config = Get-Content -Path $templateConfig -Raw | ConvertFrom-Json
        if ($null -ne $config -and -not [string]::IsNullOrWhiteSpace($config.pluginVersion)) {
            return $config.pluginVersion.Trim()
        }
    } catch {
        return "0.3.0"
    }

    return "0.3.0"
}

function Convert-ToMsiVersion {
    param([string]$VersionText)

    $core = if ([string]::IsNullOrWhiteSpace($VersionText)) { "0.3.0" } else { $VersionText.Trim() }
    $dashIndex = $core.IndexOf("-")
    if ($dashIndex -gt 0) {
        $core = $core.Substring(0, $dashIndex)
    }

    $parts = $core.Split(".")
    $major = 0
    $minor = 1
    $build = 0

    if ($parts.Length -gt 0) {
        [void][int]::TryParse($parts[0], [ref]$major)
    }
    if ($parts.Length -gt 1) {
        [void][int]::TryParse($parts[1], [ref]$minor)
    }
    if ($parts.Length -gt 2) {
        [void][int]::TryParse($parts[2], [ref]$build)
    }

    $major = [Math]::Min([Math]::Max($major, 0), 65535)
    $minor = [Math]::Min([Math]::Max($minor, 0), 65535)
    $build = [Math]::Min([Math]::Max($build, 0), 65535)

    return "$major.$minor.$build"
}

function Resolve-WixExe {
    param([string]$RequestedPath)

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        if (-not (Test-Path -Path $RequestedPath -PathType Leaf)) {
            throw "WixPath was provided but does not exist: $RequestedPath"
        }
        return (Resolve-Path $RequestedPath).Path
    }

    $cmd = Get-Command wix -ErrorAction SilentlyContinue
    if ($null -ne $cmd) {
        return "wix"
    }

    $cmdExe = Get-Command wix.exe -ErrorAction SilentlyContinue
    if ($null -ne $cmdExe) {
        return "wix.exe"
    }

    return ""
}

function Get-WixInstallGuidance {
    return @"
WiX CLI (v4) is required for MSI packaging.

Install options:
1) .NET tool (recommended):
   dotnet tool install --global wix
   # or update:
   dotnet tool update --global wix

2) Official WiX installer:
   https://wixtoolset.org

Then re-run this script, or pass -SkipMsi to create ZIP only.
"@
}

function Copy-PackageFiles {
    param(
        [string]$RepoRoot,
        [string]$AddinOutputDir,
        [string]$PackageDir,
        [bool]$IncludeDebugSymbols
    )

    Ensure-Directory -PathValue $PackageDir

    $binDir = Join-Path $PackageDir "bin"
    Ensure-Directory -PathValue $binDir

    Write-Host "==> Copying addin binaries"
    $files = Get-ChildItem -Path $AddinOutputDir -File
    foreach ($file in $files) {
        if (-not $IncludeDebugSymbols -and $file.Extension.Equals(".pdb", [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        Copy-Item -Path $file.FullName -Destination (Join-Path $binDir $file.Name) -Force
    }

    $configSource = Join-Path $RepoRoot "config\appsettings.template.json"
    if (Test-Path $configSource) {
        $configDir = Join-Path $PackageDir "config"
        Ensure-Directory -PathValue $configDir
        Copy-Item -Path $configSource -Destination (Join-Path $configDir "appsettings.template.json") -Force
    }

    $scriptsDir = Join-Path $PackageDir "deploy\scripts"
    Ensure-Directory -PathValue $scriptsDir
    Copy-Item -Path (Join-Path $RepoRoot "deploy\scripts\install-from-package.ps1") -Destination (Join-Path $scriptsDir "install-from-package.ps1") -Force
    Copy-Item -Path (Join-Path $RepoRoot "deploy\scripts\uninstall-local.ps1") -Destination (Join-Path $scriptsDir "uninstall-local.ps1") -Force

    $docsSource = Join-Path $RepoRoot "docs\deployment.md"
    if (Test-Path $docsSource) {
        $docsDir = Join-Path $PackageDir "docs"
        Ensure-Directory -PathValue $docsDir
        Copy-Item -Path $docsSource -Destination (Join-Path $docsDir "deployment.md") -Force
    }
}

function Write-MsiManifestTemplate {
    param(
        [string]$ManifestPath,
        [string]$RevitVersionValue
    )

    # Revit reads the manifest at runtime; using %LocalAppData% keeps path user-specific.
    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>MDR Revit Plugin</Name>
    <Assembly>%LocalAppData%\MDR\RevitPlugin\addin\$RevitVersionValue\Mdr.Revit.Addin.dll</Assembly>
    <AddInId>8D83C886-B739-4ACD-A9DB-1BC78F315B3A</AddInId>
    <FullClassName>Mdr.Revit.Addin.RevitExternalApplication</FullClassName>
    <VendorId>MDR</VendorId>
    <VendorDescription>MDR BIM Integration</VendorDescription>
  </AddIn>
</RevitAddIns>
"@

    Set-Content -Path $ManifestPath -Value $xml -Encoding UTF8
}

function Escape-XmlAttribute {
    param([string]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape($Value)
}

function Write-WixSource {
    param(
        [string]$TemplatePath,
        [string]$WxsPath,
        [string]$ProductNameValue,
        [string]$ManufacturerValue,
        [string]$MsiVersion,
        [string]$UpgradeCodeValue,
        [string]$RevitVersionValue,
        [string[]]$BinaryFiles
    )

    if (-not (Test-Path -Path $TemplatePath -PathType Leaf)) {
        throw "WiX template file was not found: $TemplatePath"
    }

    $componentLines = New-Object System.Collections.Generic.List[string]
    $componentRefLines = New-Object System.Collections.Generic.List[string]
    $fileId = 1
    foreach ($fileName in $BinaryFiles) {
        $safeFileId = "FBin$fileId"
        $safeComponentId = "CmpBin$fileId"
        $escaped = Escape-XmlAttribute -Value $fileName
        [void]$componentLines.Add("      <Component Id=""$safeComponentId"" Guid=""*"">")
        [void]$componentLines.Add("        <File Id=""$safeFileId"" Source=""`$(var.PackageRoot)\bin\$escaped"" KeyPath=""yes"" />")
        [void]$componentLines.Add("      </Component>")
        [void]$componentRefLines.Add("      <ComponentRef Id=""$safeComponentId"" />")
        $fileId++
    }

    $binaryComponents = $componentLines -join [Environment]::NewLine
    $binaryComponentRefs = $componentRefLines -join [Environment]::NewLine
    $wxs = Get-Content -Path $TemplatePath -Raw
    $wxs = $wxs.Replace("__PRODUCT_NAME__", (Escape-XmlAttribute -Value $ProductNameValue))
    $wxs = $wxs.Replace("__MANUFACTURER__", (Escape-XmlAttribute -Value $ManufacturerValue))
    $wxs = $wxs.Replace("__MSI_VERSION__", (Escape-XmlAttribute -Value $MsiVersion))
    $wxs = $wxs.Replace("__UPGRADE_CODE__", (Escape-XmlAttribute -Value $UpgradeCodeValue))
    $wxs = $wxs.Replace("__REVIT_VERSION__", (Escape-XmlAttribute -Value $RevitVersionValue))
    $wxs = $wxs.Replace("__BINARY_COMPONENTS__", $binaryComponents)
    $wxs = $wxs.Replace("__BINARY_COMPONENT_REFS__", $binaryComponentRefs)

    Set-Content -Path $WxsPath -Value $wxs -Encoding UTF8
}

$repoRoot = Resolve-RepoRoot
$dotnet = Resolve-Dotnet -RepoRoot $repoRoot
$solutionPath = Join-Path $repoRoot "Mdr.RevitPlugin.sln"
$addinOutput = Join-Path $repoRoot "src\Mdr.Revit.Addin\bin\$Configuration\net48"
$resolvedVersion = Resolve-PluginVersion -RepoRoot $repoRoot -RequestedVersion $Version
$msiVersion = Convert-ToMsiVersion -VersionText $resolvedVersion

$outputRoot = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    Join-Path $repoRoot "artifacts\release"
} else {
    $OutputDirectory
}

Ensure-Directory -PathValue $outputRoot

$packageName = "mdr-rvt-plugin-$resolvedVersion-rvt$RevitVersion"
$stagingRoot = Join-Path $outputRoot ("staging\" + $packageName)
$packageRoot = Join-Path $stagingRoot "package"
$msiWorkDir = Join-Path $stagingRoot "msi"
$wixTemplatePath = Join-Path $repoRoot "deploy\wix\Product.wxs.template"

if (Test-Path $stagingRoot) {
    Remove-Item -Path $stagingRoot -Recurse -Force
}

Ensure-Directory -PathValue $packageRoot
Ensure-Directory -PathValue $msiWorkDir

Write-Host "==> Packaging MDR Revit Plugin"
Write-Host "Repo root:      $repoRoot"
Write-Host "Configuration:  $Configuration"
Write-Host "Plugin version: $resolvedVersion"
Write-Host "MSI version:    $msiVersion"
Write-Host "Revit version:  $RevitVersion"
Write-Host "Output root:    $outputRoot"

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

Copy-PackageFiles `
    -RepoRoot $repoRoot `
    -AddinOutputDir $addinOutput `
    -PackageDir $packageRoot `
    -IncludeDebugSymbols:$IncludePdb

$zipPath = Join-Path $outputRoot ($packageName + ".zip")
if (-not $SkipZip) {
    Write-Host "==> Creating zip package: $zipPath"
    if (Test-Path $zipPath) {
        Remove-Item -Path $zipPath -Force
    }

    Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath -Force
}

$msiPath = Join-Path $outputRoot ($packageName + ".msi")
if (-not $SkipMsi) {
    $wixExe = Resolve-WixExe -RequestedPath $WixPath
    if ([string]::IsNullOrWhiteSpace($wixExe)) {
        $guidance = Get-WixInstallGuidance
        throw "WiX CLI was not found.`n`n$guidance"
    }

    $manifestOutputDir = Join-Path $packageRoot "msi"
    Ensure-Directory -PathValue $manifestOutputDir
    $manifestTemplate = Join-Path $manifestOutputDir "Mdr.Revit.addin"
    Write-MsiManifestTemplate -ManifestPath $manifestTemplate -RevitVersionValue $RevitVersion

    $binaryDir = Join-Path $packageRoot "bin"
    $binaryFiles = Get-ChildItem -Path $binaryDir -File | Select-Object -ExpandProperty Name
    if ($null -eq $binaryFiles -or $binaryFiles.Count -eq 0) {
        throw "No plugin binaries were found in $binaryDir"
    }

    $wxsPath = Join-Path $msiWorkDir "Product.wxs"
    Write-WixSource `
        -TemplatePath $wixTemplatePath `
        -WxsPath $wxsPath `
        -ProductNameValue $ProductName `
        -ManufacturerValue $Manufacturer `
        -MsiVersion $msiVersion `
        -UpgradeCodeValue $UpgradeCode `
        -RevitVersionValue $RevitVersion `
        -BinaryFiles $binaryFiles

    Write-Host "==> Building MSI via WiX: $msiPath"
    if (Test-Path $msiPath) {
        Remove-Item -Path $msiPath -Force
    }

    & $wixExe build `
        -arch x64 `
        -out $msiPath `
        $wxsPath `
        -d PackageRoot=$packageRoot
    if ($LASTEXITCODE -ne 0) {
        throw "wix build failed with exit code $LASTEXITCODE."
    }
}

if (-not $KeepStaging) {
    if (Test-Path $stagingRoot) {
        Remove-Item -Path $stagingRoot -Recurse -Force
    }
}

Write-Host ""
Write-Host "Packaging completed."
if (-not $SkipZip) {
    Write-Host "ZIP: $zipPath"
}
if (-not $SkipMsi) {
    Write-Host "MSI: $msiPath"
}
