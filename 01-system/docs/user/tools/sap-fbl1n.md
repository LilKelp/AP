# SAP FBL1N Export
**Category**: ops  
**Version**: v0.5 (Released: 2025-12-14)

## What it does
- Uses SAP GUI scripting to run FBL1N for specified company codes and exports the ALV grid to Excel.
- Supports Local File export to an Excel `.xls` saved to a target folder (fills `DY_PATH`/`DY_FILENAME`) with a Spreadsheet-export fallback to `.xlsx` when needed.

## Inputs
- `CompanyCode`: SAP company code (AU `8000`, NZ `8100`).
- `KeyDate`: Key date `dd/MM/yyyy` (e.g., `15/12/2025`).
- `OutputDir`: Target folder (recommended: `02-inputs/Payment run raw`). If `OutputDir` is relative, it resolves to the repo root based on the script location; AU/NZ subfolders auto-add when `OutputDir` ends with `Payment run raw`.
- `LayoutVariant` (optional): ALV variant name.
- `Mode`: `localfile` (default) saves `<dd.MM.yy>.xls`; `spreadsheet` saves `FBL1N_<bukrs>_<yyyymmdd>.xlsx`.
- `radio=<id>` (optional): Override the export radio control id if your SAP dialog differs.
- `dump` (optional): Print export dialog controls (useful to find the right `radio=` value).
- **Prereq**: SAP GUI session is logged in (use `sap-login` if needed).

## Steps (routine)
1. Ensure SAP GUI scripting is enabled (client and server).
2. If SAP is not logged in yet, run: `python 01-system/tools/ops/sap-login/sap_login.py`.
3. Run the VBScript export (local-file mode recommended for payment-list input):
   - AU: `cscript //Nologo 01-system/tools/ops/sap-fbl1n/sap_fbl1n_export.vbs 8000 15/12/2025 "02-inputs/Payment run raw" mode=localfile`
   - NZ: `cscript //Nologo 01-system/tools/ops/sap-fbl1n/sap_fbl1n_export.vbs 8100 15/12/2025 "02-inputs/Payment run raw" mode=localfile`
4. If local-file export fails, rerun with `mode=spreadsheet` (saves `FBL1N_<bukrs>_<yyyymmdd>.xlsx`).
5. Collect the exported files from the target folder (for example `02-inputs/Payment run raw/AU/15.12.25.xls`).

## Outputs
- Local file export: `02-inputs/Payment run raw/<REGION>/<dd.MM.yy>.xls` (or under `OutputDir` if you set a different folder).
- Spreadsheet export: `02-inputs/<dir>/FBL1N_<bukrs>_<yyyymmdd>.xlsx`.

## Notes
- Field IDs are based on standard FBL1N. If a control is not found, record a short SAP GUI script and update the IDs inside `01-system/tools/ops/sap-fbl1n/sap_fbl1n_export.vbs`.
- Files are overwritten on each run.
- VBScript Local File mode sets `DY_PATH` and `DY_FILENAME` when SAP shows an internal save dialog; if your SAP opens a Windows "Save As" dialog instead, the script falls back to SendKeys, so keep the Save As window focused until the file saves.
- `01-system/tools/ops/sap-fbl1n/sap_fbl1n_export.ps1` remains as a deprecated helper (not the registered tool entrypoint).

## Troubleshooting
- If scripting is blocked, enable “Allow Scripting” in SAP GUI options and ensure the server permits scripting.
- If control IDs differ, use SAP GUI Script Recorder on FBL1N and adjust the IDs in `01-system/tools/ops/sap-fbl1n/sap_fbl1n_export.vbs` (or run with `dump`).

## Change Log
- v0.5 (2025-12-14): Switch tool registry entrypoint to the VBScript exporter (local-file default + spreadsheet fallback); keep PowerShell helper as deprecated.
- v0.4 (2025-12-14): Removed clipboard mode and defaulted VBScript to `mode=localfile`; relative OutputDir now resolves to repo root to avoid SAP work_dir saves.
- v0.3 (2025-12-12): Switch spreadsheet mode to Local File export, add `mode=localfile` alias, and fall back to SendKeys when Windows Save As is used.
- v0.2 (2025-12-11): Add Spreadsheet save mode with path/filename fill (`mode=spreadsheet`, optional `radio=` override).
- v0.1 (2025-12-02): initial version
