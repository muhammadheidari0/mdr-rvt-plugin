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

## Log locations

- Runtime logs: `%LocalAppData%/MDR/RevitPlugin/logs`
- Exported publish artifacts (default): `%LocalAppData%/MDR/RevitPlugin/publish`
