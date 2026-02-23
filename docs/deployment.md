# Deployment Notes

## Minimum environment

- Revit 2026+
- .NET Framework 4.8 runtime
- MDR backend reachable over HTTPS/HTTP

## Local config

1. Copy `config/appsettings.template.json` to local runtime config path.
2. Set:
   - `apiBaseUrl`
   - `projectCode`
   - `pluginVersion`
   - `publishOutputDirectory`
3. Keep `retryFailedItems=true` for EDMS MVP.

## Addin manifest

- Template file: `deploy/addin/Mdr.Revit.addin`
- Update `<Assembly>` to deployed `Mdr.Revit.Addin.dll` absolute path.

## Install options

### Local developer install (build from source)

```powershell
.\deploy\scripts\install-local.ps1 -RevitVersion 2026 -Configuration Release
```

### Install from packaged binaries (no SDK/source required)

```powershell
.\deploy\scripts\install-from-package.ps1 -PackagePath "C:\drop\mdr-rvt-plugin.zip" -RevitVersion 2026 -ForceConfig
```

- `PackagePath` accepts a folder or `.zip`.
- Script auto-detects `Mdr.Revit.Addin.dll` and writes the `.addin` manifest.
- Google OAuth values can be injected via env vars:
  - `MDR_GOOGLE_CLIENT_ID`
  - `MDR_GOOGLE_CLIENT_SECRET`
  - `MDR_GOOGLE_REFRESH_TOKEN`

## Release packaging

### Build ZIP + MSI (official release)

```powershell
.\deploy\scripts\package-msi.ps1 -Configuration Release -RevitVersion 2026
```

- Outputs are created under `artifacts\release`.
- ZIP can be installed with `install-from-package.ps1`.
- MSI build requires WiX v4 (`wix.exe` on PATH) or explicit `-WixPath`.
- WiX source template lives at `deploy/wix/Product.wxs.template` and is materialized during packaging.
- If WiX is missing, script prints installation guidance (`dotnet tool install --global wix`) and exits with a clear error.

### ZIP-only packaging (when WiX is not installed)

```powershell
.\deploy\scripts\package-msi.ps1 -Configuration Release -RevitVersion 2026 -SkipMsi
```

## Validation checklist

- Clean machine MSI validation checklist:
  - `docs/clean-machine-msi-checklist.md`

## Log locations

- Runtime logs: `%LocalAppData%/MDR/RevitPlugin/logs`
- Exported publish artifacts (default): `%LocalAppData%/MDR/RevitPlugin/publish`
