# Clean Machine MSI Checklist

Use this checklist on a clean Windows machine to validate first-time installation.

## Build Under Test (v0.3.6)

- ZIP: `mdr-rvt-plugin-0.3.6-rvt2026.zip`
- MSI: `mdr-rvt-plugin-0.3.6-rvt2026.msi`
- ZIP SHA-256: `5598A39A26CB05AE3FA7FAD114E2428788651F6E143555B5B5E9ED56578D1F76`
- MSI SHA-256: `8A753893BE1239473AE8A84E52BC1CE97901C2D2E9F4E10200637BD0032422AA`

## 1) Preconditions

- Windows user has local install permission.
- Revit 2026 is installed and has been opened once.
- No old MDR plugin is installed:
  - `%AppData%\Autodesk\Revit\Addins\2026\Mdr.Revit.addin` must not exist.
  - `%LocalAppData%\MDR\RevitPlugin\addin\2026\` must not exist.

## 2) Install MSI

- Run MSI (double-click) or silent install:

```powershell
msiexec /i "mdr-rvt-plugin-0.3.6-rvt2026.msi" /qn /norestart
```

- Verify installer exits successfully (`ExitCode 0`).

## 3) File Placement Validation

- Confirm plugin binaries:
  - `%LocalAppData%\MDR\RevitPlugin\addin\2026\Mdr.Revit.Addin.dll`
- Confirm manifest:
  - `%AppData%\Autodesk\Revit\Addins\2026\Mdr.Revit.addin`
- Open manifest and verify:
  - `FullClassName` = `Mdr.Revit.Addin.RevitExternalApplication`
  - `Assembly` points to `%LocalAppData%\MDR\RevitPlugin\addin\2026\Mdr.Revit.Addin.dll`

## 4) Revit Smoke Test

- Start Revit 2026 and open a test model.
- Verify `MDR` tab appears.
- Verify buttons appear at least:
  - `Google Sheets Sync`
  - `Smart Numbering`
- Click each command:
  - Window opens without crash.
  - Close works.

## 5) Smart Numbering Quick Test

- Select a few elements with known parameters (`Mark` or `Comments`).
- Open `Smart Numbering`.
- Use formula: `{Mark}-{Sequence:5}`.
- Verify preview loads and `Apply` is enabled only when no errors exist.
- Apply once and confirm values are written.
- Undo once in Revit and confirm rollback behavior is clean.

## 6) Uninstall Validation

- Uninstall from Apps & Features or:

```powershell
msiexec /x "mdr-rvt-plugin-0.3.6-rvt2026.msi" /qn /norestart
```

- Verify removed:
  - `%AppData%\Autodesk\Revit\Addins\2026\Mdr.Revit.addin`
  - `%LocalAppData%\MDR\RevitPlugin\addin\2026\`

## 7) Pass/Fail Criteria

- Pass: install, startup, ribbon visibility, command launch, quick apply, uninstall all succeed.
- Fail: any missing file, Revit load error, command crash, or uninstall residue.
