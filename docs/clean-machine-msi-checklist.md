# Clean Machine MSI Checklist

Use this checklist on a clean Windows machine to validate first-time installation.

## 1) Preconditions

- Windows user has local install permission.
- Revit 2026 is installed and has been opened once.
- No old MDR plugin is installed:
  - `%AppData%\Autodesk\Revit\Addins\2026\Mdr.Revit.addin` must not exist.
  - `%LocalAppData%\MDR\RevitPlugin\addin\2026\` must not exist.

## 2) Install MSI

- Run MSI (double-click) or silent install:

```powershell
msiexec /i "mdr-rvt-plugin-<version>-rvt2026.msi" /qn /norestart
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
msiexec /x "mdr-rvt-plugin-<version>-rvt2026.msi" /qn /norestart
```

- Verify removed:
  - `%AppData%\Autodesk\Revit\Addins\2026\Mdr.Revit.addin`
  - `%LocalAppData%\MDR\RevitPlugin\addin\2026\`

## 7) Pass/Fail Criteria

- Pass: install, startup, ribbon visibility, command launch, quick apply, uninstall all succeed.
- Fail: any missing file, Revit load error, command crash, or uninstall residue.
