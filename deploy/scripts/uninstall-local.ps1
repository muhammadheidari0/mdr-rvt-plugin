param(
    [string]$RevitVersion = "2026",
    [switch]$RemoveConfig
)

$ErrorActionPreference = "Stop"

$pluginInstallDir = Join-Path $env:LocalAppData "MDR\RevitPlugin\addin\$RevitVersion"
$manifestPath = Join-Path $env:AppData "Autodesk\Revit\Addins\$RevitVersion\Mdr.Revit.addin"
$runtimeConfigPath = Join-Path $env:LocalAppData "MDR\RevitPlugin\config.json"

Write-Host "==> MDR Revit local uninstall"
Write-Host "Revit version: $RevitVersion"

if (Test-Path $manifestPath) {
    Remove-Item -Path $manifestPath -Force
    Write-Host "Removed manifest: $manifestPath"
} else {
    Write-Host "Manifest not found: $manifestPath"
}

if (Test-Path $pluginInstallDir) {
    Remove-Item -Path $pluginInstallDir -Recurse -Force
    Write-Host "Removed plugin directory: $pluginInstallDir"
} else {
    Write-Host "Plugin directory not found: $pluginInstallDir"
}

if ($RemoveConfig -and (Test-Path $runtimeConfigPath)) {
    Remove-Item -Path $runtimeConfigPath -Force
    Write-Host "Removed runtime config: $runtimeConfigPath"
}

Write-Host "Uninstall completed."
